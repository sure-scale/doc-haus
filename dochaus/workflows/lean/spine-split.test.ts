import { afterEach, expect, test } from "bun:test"
import { readLeanEvents } from "./events"
import { enqueueLeanRequest } from "./requests"
import { runLean } from "./runner"
import { parseSpineSplit, SPINE_SPLIT_PERMISSIONS, spineSplitPrompt, type SpineEntry } from "./spine-split"
import { LEAN_MAX_RETRIES, LEAN_STAGE_AGENTS } from "./stages"
import { readLeanJson, readLeanText } from "./state"
import { AGENT, cleanupMatters, happyPath, matter, scriptedModel, SOURCE, type Reply } from "./test-support"

// The spine step's split of each entry into atoms, driven through the real
// coordinator with a scripted model: one job per entry, checked by code, its
// atoms written to spine.json in the entries' order and saved by content.

afterEach(cleanupMatters)

const DELIVERY = { source_id: SOURCE, quote: "within thirty days of the order date" }
const PAYMENT = { source_id: SOURCE, quote: "Payment is due on delivery." }
const ONLY: SpineEntry = { kind: "fact", detail: "The supply terms are the only source of the note.", evidence: [{ source_id: "request", quote: "Draft a short supply note" }], limitations: [] }
const COMPOUND: SpineEntry = { kind: "requirement", detail: "The note states that the supplier delivers within thirty days of the order date and that payment is due on delivery.", evidence: [DELIVERY, PAYMENT], limitations: [] }
const SILENT: SpineEntry = { kind: "open-item", detail: "The terms give no remedy for late delivery.", evidence: [], limitations: ["The supply terms are silent on late delivery."] }
const ATOMS = [
  { kind: "requirement", detail: "The note states that the supplier delivers within thirty days of the order date.", evidence: [0], limitations: [] },
  { kind: "requirement", detail: "The note states that payment is due on delivery.", evidence: [1], limitations: [] },
]
const SPLIT = [
  { ...COMPOUND, detail: ATOMS[0]!.detail, evidence: [DELIVERY] },
  { ...COMPOUND, detail: ATOMS[1]!.detail, evidence: [PAYMENT] },
]

const spineReply = (...entries: SpineEntry[]) => () => JSON.stringify({ entries, limitations: [] })
/** Splits the compound entry and finds every other entry atomic. */
const splitting: Reply = prompt => JSON.stringify(prompt.includes(COMPOUND.detail) ? { atoms: ATOMS } : { atomic: true })
const splitPrompts = (model: ReturnType<typeof scriptedModel>) => model.prompts.filter(prompt => prompt.agent === AGENT.split)
const spineEntries = (directory: string) => (readLeanJson(directory, "spine.json") as { entries: SpineEntry[] }).entries

test("the prompt shows one entry, its evidence and limitations by index in tagged blocks, and needs no lookup", () => {
  const prompt = spineSplitPrompt(COMPOUND)
  for (const tag of ["entry", "evidence", "limitations", "output_contract"]) expect(prompt).toContain(`<${tag}>`)
  expect(prompt).toContain(`<entry>\nkind: requirement\ndetail: ${COMPOUND.detail}\n</entry>`)
  expect(prompt).toContain(`[0] ${SOURCE}: "${DELIVERY.quote}"\n[1] ${SOURCE}: "${PAYMENT.quote}"`)
  expect(prompt).toContain("<limitations>\nNone.\n</limitations>")
  expect(prompt).toContain("no lookup is needed")
})

test("an atomic entry is itself; a compound entry becomes atoms carrying the entry's own evidence", () => {
  expect(parseSpineSplit('{"atomic":true}', COMPOUND)).toEqual({ issues: [], atoms: [COMPOUND] })
  expect(parseSpineSplit(JSON.stringify({ atoms: ATOMS }), COMPOUND)).toEqual({ issues: [], atoms: SPLIT })
  // A limitation goes with the atom it qualifies; an entry without evidence splits into atoms that each carry one.
  const unproven: SpineEntry = { kind: "open-item", detail: "Neither the goods nor the price are stated.", evidence: [], limitations: ["The terms list no goods.", "The terms state no price."] }
  expect(parseSpineSplit(JSON.stringify({ atoms: [
    { kind: "open-item", detail: "The goods are not stated.", evidence: [], limitations: [0] },
    { kind: "open-item", detail: "The price is not stated.", evidence: [], limitations: [1, 1] },
  ] }), unproven)).toEqual({ issues: [], atoms: [
    { ...unproven, detail: "The goods are not stated.", limitations: ["The terms list no goods."] },
    { ...unproven, detail: "The price is not stated.", limitations: ["The terms state no price."] },
  ] })
})

test("the check refuses a kind, a reference or a detail the entry does not hold, and an atom or split that drops part of it", () => {
  const issues = (atoms: object[], parent = COMPOUND) => parseSpineSplit(JSON.stringify({ atoms }), parent).issues
  const [first, second] = ATOMS as [typeof ATOMS[0], typeof ATOMS[0]]
  expect(issues([first])).toEqual([expect.stringContaining('reply {"atomic":true}')])
  expect(issues([{ ...first, kind: "drafting-choice" }, second])).toEqual(["atoms[0] kind must be decision, recommendation, requirement, fact, open-item."])
  expect(issues([{ ...first, evidence: [0, 2] }, second])).toEqual(["atoms[0].evidence lists [2]: use evidence indexes 0 to 1."])
  expect(issues([first, { ...second, evidence: [] }])).toEqual(["atoms[1] cites no evidence; cite the index of the evidence that supports it.", "Evidence [1] goes to no atom; cite it in the atom it supports."])
  expect(issues([first, { ...second, limitations: [0] }])).toEqual(["atoms[1].limitations lists [0], but the entry has no limitation; leave it empty."])
  // A figure, a quoted passage or a capitalised term the entry does not hold is added detail.
  expect(issues([{ ...first, detail: "The note states that the supplier delivers within 30 days of the order date." }, second])).toEqual([
    'atoms[0] says "30", which the entry does not; use only the entry\'s own words.',
  ])
  expect(issues([first, { ...second, detail: 'The note states that "payment is due on invoice".' }])).toEqual([
    'atoms[1] says "payment is due on invoice", which the entry does not; use only the entry\'s own words.',
  ])
  expect(issues([first, { ...second, detail: "The note states that payment to Acme is due on delivery." }])).toEqual([
    'atoms[1] says "Acme", which the entry does not; use only the entry\'s own words.',
  ])
  // Words the entry or its quotations hold, in any case, and a sentence's opening word are its own.
  expect(issues([first, { ...second, detail: "Payment is due on delivery. The Supplier is paid then." }])).toEqual([])
  // A figure the entry states must stay in some atom.
  const dated: SpineEntry = { ...COMPOUND, detail: "Delivery is due within 30 days and payment is due on delivery." }
  expect(issues([{ ...first, detail: "Delivery is due within the period." }, second], dated)).toEqual(['The entry\'s figure "30" is in no atom; keep it in the atom that states it.'])
  expect(issues([first, first])).toEqual(["Evidence [1] goes to no atom; cite it in the atom it supports.", "atoms[1] repeats atoms[0]; each atom states a different point."])
  expect(parseSpineSplit("Two points.", COMPOUND).issues).toEqual([expect.stringContaining('Reply with the JSON object only')])
})

test("spine.json holds the atoms in the entries' order, each split in its entry's place, from tool-less split sessions", async () => {
  const directory = matter()
  // The first entry's job answers last, so the order is the entries', not the replies'.
  const model = scriptedModel({ ...happyPath, [AGENT.spine]: spineReply(ONLY, COMPOUND, SILENT), [AGENT.split]: splitting })
  const delayed = { ...model.port, async prompt(input: Parameters<typeof model.port.prompt>[0]) {
    if (input.agent === AGENT.split) await Bun.sleep(input.text.includes(ONLY.detail) ? 30 : input.text.includes(COMPOUND.detail) ? 15 : 0)
    return model.port.prompt(input)
  } }
  const state = await runLean(directory, { sessions: delayed })

  expect(state.status).toBe("completed")
  expect(spineEntries(directory)).toEqual([ONLY, ...SPLIT, SILENT])
  expect(splitPrompts(model)).toHaveLength(3)
  const sessions = model.sessions.filter(session => session.agent === AGENT.split)
  expect(sessions).toHaveLength(3)
  for (const session of sessions) expect(session).toMatchObject({ parentID: "ses_owner", permissions: SPINE_SPLIT_PERMISSIONS })
  expect(LEAN_STAGE_AGENTS.split).toBe("lean-spine-split")
  // The step reports its split sessions beside its own.
  expect(state.stages.spine?.session_ids).toHaveLength(3)
  expect(readLeanEvents(directory).filter(event => event.event === "spine_split")).toEqual([expect.objectContaining({
    stage: "spine", counts: { entries: 3, reused: 0, atomic: 2, split: 1, kept: 0, atoms: 4 },
  })])
})

test("a split whose reference never passes its check is retried in its session, then keeps its entry", async () => {
  const directory = matter()
  const bad = JSON.stringify({ atoms: [ATOMS[0], { ...ATOMS[1], evidence: [1, 5] }] })
  const model = scriptedModel({ ...happyPath, [AGENT.spine]: spineReply(ONLY, COMPOUND), [AGENT.split]: prompt => prompt.includes(COMPOUND.detail) || prompt.includes("did not pass") ? bad : JSON.stringify({ atomic: true }) })
  const state = await runLean(directory, { sessions: model.port })

  expect(state.status).toBe("completed")
  expect(state.stages.spine?.status).toBe("done")
  expect(spineEntries(directory)).toEqual([ONLY, COMPOUND])
  const compound = splitPrompts(model).filter(prompt => prompt.sessionID === splitPrompts(model).find(row => row.text.includes(COMPOUND.detail))!.sessionID)
  expect(compound).toHaveLength(LEAN_MAX_RETRIES + 1)
  for (const retry of compound.slice(1)) expect(retry.text).toContain("- atoms[1].evidence lists [5]: use evidence indexes 0 to 1.")
  const checks = readLeanEvents(directory).filter(event => event.event === "checker" && event.step === "split entry 2")
  expect(checks.map(event => event.attempt)).toEqual([0, 1, 2, 3])
  expect(readLeanEvents(directory).find(event => event.event === "spine_split")?.counts).toMatchObject({ atomic: 1, split: 0, kept: 1, atoms: 2 })
})

test("a resumed spine splits only the entries it has not split before", async () => {
  const directory = matter()
  let failed = false
  const failing: Reply = (prompt, index) => {
    if (prompt.includes(SILENT.detail) && !failed) { failed = true; throw new Error("Scripted split failure") }
    return splitting(prompt, index)
  }
  const replies = { ...happyPath, [AGENT.spine]: spineReply(ONLY, COMPOUND, SILENT) }
  const first = scriptedModel({ ...replies, [AGENT.split]: failing })
  const stopped = await runLean(directory, { sessions: first.port })
  expect(stopped.status).toBe("failed")
  expect(stopped.stages.spine?.status).not.toBe("done")
  expect(readLeanText(directory, "spine.json")).toBeUndefined()

  enqueueLeanRequest(directory, { kind: "continue" })
  const resumed = scriptedModel({ ...replies, [AGENT.split]: splitting })
  expect((await runLean(directory, { sessions: resumed.port })).status).toBe("completed")
  expect(splitPrompts(resumed).map(prompt => prompt.text.includes(SILENT.detail))).toEqual([true])
  expect(spineEntries(directory)).toEqual([ONLY, ...SPLIT, SILENT])
  expect(readLeanEvents(directory).findLast(event => event.event === "spine_split")?.counts).toMatchObject({ reused: 2, atomic: 1, split: 0, atoms: 4 })
})

test("the outline, review, change plan, part, notes and package review prompts read the atoms as spine.json holds them", async () => {
  const directory = matter()
  const model = scriptedModel({ ...happyPath, [AGENT.spine]: spineReply(ONLY, COMPOUND, SILENT), [AGENT.split]: splitting })
  expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")

  const spine = readLeanText(directory, "spine.json")!
  expect(JSON.parse(spine).entries).toEqual([ONLY, ...SPLIT, SILENT])
  const first = (agent: string) => model.prompts.find(prompt => prompt.agent === agent)!.text
  for (const agent of [AGENT.draft, AGENT.review, AGENT["change-plan"], AGENT.finalize, AGENT["package-review"]]) expect({ agent, spine: first(agent).includes(spine.trim()) }).toEqual({ agent, spine: true })
  // Each part prompt points the drafter at the same record.
  const parts = model.prompts.filter(prompt => prompt.agent === AGENT.draft).slice(1)
  expect(parts.length).toBeGreaterThan(0)
  for (const part of parts) expect(part.text).toContain(".dochaus/lean/spine.json")
  // No model job is shown the compound entry outside its own split.
  expect(model.prompts.filter(prompt => prompt.agent !== AGENT.split && prompt.agent !== AGENT.spine && prompt.text.includes(COMPOUND.detail))).toEqual([])
})
