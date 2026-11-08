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
// atoms written to spine-atoms.json in the entries' order and saved by content,
// while spine.json keeps the entries as the spine wrote them.

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
const spineAtoms = (directory: string) => (readLeanJson(directory, "spine-atoms.json") as { entries: SpineEntry[] }).entries
const spineParents = (directory: string) => (readLeanJson(directory, "spine.json") as { entries: SpineEntry[] }).entries

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

test("a reply of one atom keeps the entry as written, whatever the atom says", () => {
  for (const atom of [ATOMS[0], { kind: "fact", detail: "Delivery is due in 30 days to Acme.", evidence: [7], limitations: [0] }])
    expect(parseSpineSplit(JSON.stringify({ atoms: [atom] }), COMPOUND)).toEqual({ issues: [], atoms: [COMPOUND] })
  expect(parseSpineSplit(JSON.stringify({ atoms: [] }), COMPOUND).issues).toEqual([expect.stringContaining("Reply with the JSON object only")])
})

test("a split job that replies with one atom is atomic at once: no retry, the entry unchanged in spine-atoms.json", async () => {
  const directory = matter()
  const model = scriptedModel({ ...happyPath, [AGENT.spine]: spineReply(ONLY, COMPOUND), [AGENT.split]: () => JSON.stringify({ atoms: [ATOMS[0]] }) })
  expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
  expect(spineAtoms(directory)).toEqual([ONLY, COMPOUND])
  expect(splitPrompts(model)).toHaveLength(2)
  expect(readLeanEvents(directory).find(event => event.event === "spine_split")?.counts).toEqual({ entries: 2, reused: 0, atomic: 2, split: 0, kept: 0, atoms: 2 })
})

test("the check refuses a kind, a reference or a detail the entry does not hold, and an atom or split that drops part of it", () => {
  const issues = (atoms: object[], parent = COMPOUND) => parseSpineSplit(JSON.stringify({ atoms }), parent).issues
  const [first, second] = ATOMS as [typeof ATOMS[0], typeof ATOMS[0]]
  expect(issues([{ ...first, kind: "drafting-choice" }, second])).toEqual(["atoms[0] kind must be decision, recommendation, requirement, fact, open-item."])
  expect(issues([{ ...first, evidence: [0, 2] }, second])).toEqual(["atoms[0].evidence lists [2]: use evidence indexes 0 to 1."])
  const uncited = "atoms[1] cites no evidence; a limitation does not replace evidence. Cite the index of the evidence that supports it, 0 to 1."
  expect(issues([first, { ...second, evidence: [] }])).toEqual([uncited, "Evidence [1] goes to no atom; cite it in the atom it supports."])
  // Under an entry with evidence, an atom carrying only a limitation is still uncited: the prompt asks a limitation only of an entry without evidence.
  const limited = { ...COMPOUND, limitations: ["The order date is not stated."] }
  expect(issues([{ ...first, limitations: [0] }, { ...second, evidence: [], limitations: [0] }], limited)).toEqual([uncited, "Evidence [1] goes to no atom; cite it in the atom it supports."])
  expect(spineSplitPrompt(limited)).toContain("When the entry has no evidence, each atom carries at least one limitation.")
  expect(issues([first, { ...second, limitations: [0] }])).toEqual(["atoms[1].limitations lists [0], but the entry has no limitation; leave it empty."])
  // A figure, a quoted passage or a capitalised term the entry does not hold is added detail.
  expect(issues([{ ...first, detail: "The note states that the supplier delivers within thirty days of the order date, with 1.5% interest." }, second])).toEqual([
    'atoms[0] says "1.5", which the entry does not; use only the entry\'s own words.',
  ])
  // A bare number of up to four digits is not a figure.
  expect(issues([{ ...first, detail: "The note states that the supplier delivers within 30 days of the order date." }, second])).toEqual([])
  expect(issues([first, { ...second, detail: 'The note states that "payment is due on invoice".' }])).toEqual([
    'atoms[1] says "payment is due on invoice", which the entry does not; use only the entry\'s own words.',
  ])
  expect(issues([first, { ...second, detail: "The note states that payment to Acme is due on delivery." }])).toEqual([
    'atoms[1] says "Acme", which the entry does not; use only the entry\'s own words.',
  ])
  // Words the entry or its quotations hold, in any case, and a sentence's opening word are its own.
  expect(issues([first, { ...second, detail: "Payment is due on delivery. The Supplier is paid then." }])).toEqual([])
  // A figure the entry states must stay in some atom; a bare number of up to four digits is not a figure.
  const dated: SpineEntry = { ...COMPOUND, detail: "Delivery is due within 30 days and a 2.5% charge is due on delivery." }
  expect(issues([{ ...first, detail: "Delivery is due within the period." }, second], dated)).toEqual(['The entry\'s figure "2.5" is in no atom; keep it in the atom that states it.'])
  // A figure an atom keeps in another form is kept.
  expect(issues([{ ...first, detail: "Delivery is due within 30 days." }, { ...second, detail: "A 2.5 % charge is due on delivery." }], dated)).toEqual([])
  expect(issues([first, first])).toEqual(["Evidence [1] goes to no atom; cite it in the atom it supports.", "atoms[1] repeats atoms[0]; each atom states a different point."])
  expect(parseSpineSplit("Two points.", COMPOUND).issues).toEqual([expect.stringContaining('Reply with the JSON object only')])
})

test("an atom cites a quote stating each figure the entry's quotes state; a figure they never state is not asked of it", () => {
  const PRICE = { source_id: SOURCE, quote: "The price is $4,800,000." }
  const priced: SpineEntry = { kind: "requirement", detail: "The note states a price of 4,800,000 and that payment is due on delivery.", evidence: [PRICE, PAYMENT], limitations: [] }
  const atoms = (price: number[], payment: number[]) => JSON.stringify({ atoms: [
    { kind: "requirement", detail: "The note states a price of 4,800,000.", evidence: price, limitations: [] },
    { kind: "requirement", detail: "The note states that payment is due on delivery.", evidence: payment, limitations: [] },
  ] })
  expect(parseSpineSplit(atoms([1], [0]), priced).issues).toEqual(['atoms[0] states "4,800,000", which none of its evidence states; cite the index of the evidence that states it.'])
  expect(parseSpineSplit(atoms([0], [1]), priced).issues).toEqual([])
  // An entry the spine kept with a figure none of its quotes states splits as before.
  expect(parseSpineSplit(atoms([1], [0]), { ...priced, evidence: [DELIVERY, PAYMENT] }).issues).toEqual([])
})

test("spine-atoms.json holds the atoms in the entries' order, each split in its entry's place, from tool-less split sessions, and spine.json the entries", async () => {
  const directory = matter()
  // The first entry's job answers last, so the order is the entries', not the replies'.
  const model = scriptedModel({ ...happyPath, [AGENT.spine]: spineReply(ONLY, COMPOUND, SILENT), [AGENT.split]: splitting })
  const delayed = { ...model.port, async prompt(input: Parameters<typeof model.port.prompt>[0]) {
    if (input.agent === AGENT.split) await Bun.sleep(input.text.includes(ONLY.detail) ? 30 : input.text.includes(COMPOUND.detail) ? 15 : 0)
    return model.port.prompt(input)
  } }
  const state = await runLean(directory, { sessions: delayed })

  expect(state.status).toBe("completed")
  expect(spineAtoms(directory)).toEqual([ONLY, ...SPLIT, SILENT])
  expect(spineParents(directory)).toEqual([ONLY, COMPOUND, SILENT])
  expect(state.stages.spine?.outputs).toEqual(["spine.json", "spine-atoms.json"])
  expect(splitPrompts(model)).toHaveLength(3)
  const sessions = model.sessions.filter(session => session.agent === AGENT.split)
  expect(sessions).toHaveLength(3)
  for (const session of sessions) expect(session).toMatchObject({ parentID: "ses_owner", permissions: SPINE_SPLIT_PERMISSIONS })
  expect(LEAN_STAGE_AGENTS.split).toBe("lean-spine-split")
  // The step reports its split sessions beside its own, as it does its absence
  // classifiers': two, as the open-item entry is not classified.
  expect(state.stages.spine?.session_ids).toHaveLength(5)
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
  expect(spineAtoms(directory)).toEqual([ONLY, COMPOUND])
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
  expect(readLeanText(directory, "spine-atoms.json")).toBeUndefined()

  enqueueLeanRequest(directory, { kind: "continue" })
  const resumed = scriptedModel({ ...replies, [AGENT.split]: splitting })
  expect((await runLean(directory, { sessions: resumed.port })).status).toBe("completed")
  expect(splitPrompts(resumed).map(prompt => prompt.text.includes(SILENT.detail))).toEqual([true])
  expect(spineAtoms(directory)).toEqual([ONLY, ...SPLIT, SILENT])
  expect(readLeanEvents(directory).findLast(event => event.event === "spine_split")?.counts).toMatchObject({ reused: 2, atomic: 1, split: 0, atoms: 4 })
})

test("content stages read spine entries, focused notes retain their open points, and level and placement jobs read the atoms", async () => {
  const directory = matter()
  const model = scriptedModel({ ...happyPath, [AGENT.spine]: spineReply(ONLY, COMPOUND, SILENT), [AGENT.split]: splitting })
  expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")

  const spine = readLeanText(directory, "spine.json")!
  expect(JSON.parse(spine).entries).toEqual([ONLY, COMPOUND, SILENT])
  const first = (agent: string) => model.prompts.find(prompt => prompt.agent === agent)!.text
  for (const agent of [AGENT.draft, AGENT.review, AGENT["change-plan"], AGENT["package-review"]]) expect({ agent, spine: first(agent).includes(spine.trim()) }).toEqual({ agent, spine: true })
  const notes = model.prompts.filter(prompt => prompt.agent === AGENT.finalize)
  expect(notes.some(prompt => prompt.text.includes(SILENT.detail))).toBe(true)
  expect(notes.every(prompt => !prompt.text.includes(spine.trim()))).toBe(true)
  // Each part prompt points the drafter at the same record.
  // Retry feedback (the placement revision) is not a part prompt.
  const parts = model.prompts.filter(prompt => prompt.agent === AGENT.draft && !prompt.text.startsWith("Your reply did not pass")).slice(1)
  expect(parts.length).toBeGreaterThan(0)
  for (const part of parts) expect(part.text).toContain(".dochaus/lean/spine.json")
  // Each atom gets its own level and placement job; none of them is shown the compound entry.
  for (const agent of [AGENT.level, AGENT.placement]) {
    const jobs = model.prompts.filter(prompt => prompt.agent === agent)
    expect([...new Set(jobs.map(prompt => [ONLY, ...SPLIT, SILENT].findIndex(atom => prompt.text.includes(atom.detail))))].toSorted()).toEqual([0, 1, 2, 3])
    expect(jobs.filter(prompt => prompt.text.includes(COMPOUND.detail))).toEqual([])
  }
})
