import { afterAll, afterEach, beforeAll, expect, test } from "bun:test"
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { advisePrompt } from "./advise"
import { type GateResult } from "./gate"
import { researchPrompt } from "./research"
import { enqueueLeanRequest } from "./requests"
import { runLean } from "./runner"
import { LEAN_PROMPT_VERSION, researchBasis, stageInputs } from "./stages"
import { hashStageInputs, leanFile, readLeanJson, readLeanText, readQuestions, writeQuestions, type Question } from "./state"
import { AGENT, NO_FINDINGS, cleanupMatters, drafter, draftingNotes, happyPath, matter, noChanges, scriptedModel, SOURCE, SOURCE_TEXT, type Reply } from "./test-support"

// Drafting Auto mode is a file-only installation opt-in read per job from
// WORKSPACE_ROOT/.preferences/preferences.json. These tests drive the real
// coordinator with the shared scripted model; the setting file is the only
// substitution besides the model. Off must stay byte-identical to the base.

let previous: string | undefined
let workspace = ""
let preferences = ""
const setAutoMode = (enabled: boolean) => {
  mkdirSync(preferences, { recursive: true })
  writeFileSync(path.join(preferences, "preferences.json"), JSON.stringify({ draftingAutoMode: enabled }))
}
beforeAll(() => {
  previous = process.env.WORKSPACE_ROOT
  workspace = mkdtempSync(path.join(tmpdir(), "drafting-auto-mode-"))
  process.env.WORKSPACE_ROOT = workspace
  preferences = path.join(workspace, ".preferences")
})
afterEach(() => {
  cleanupMatters()
  rmSync(preferences, { recursive: true, force: true })
})
afterAll(() => {
  if (previous === undefined) delete process.env.WORKSPACE_ROOT
  else process.env.WORKSPACE_ROOT = previous
  rmSync(workspace, { recursive: true, force: true })
})

const quiet = { ...happyPath, [AGENT.review]: () => NO_FINDINGS, [AGENT.changes]: noChanges }
const AUTO = "Drafting Auto mode is enabled for this installation:"

// The shared smoke fixture: a request that needs a delivery location no source
// records. The gap rides the discrepancy record, the spine and the review; the
// body marks it and the notes name it. Only intake's decision differs off vs on.
const LOCATION_REQUEST = "Draft a short supply note from the supply terms. State the delivery location."
const LOCATION_MARKER = "[open item: confirm the delivery location]"
const LOCATION_NOTE = "The delivery location is not recorded in the supply terms; the lawyer must confirm it before the note is used."
const LOCATION_GAP = {
  detail: "The delivery location is not recorded; the note must state it.",
  requirement: { source_id: "request", quote: "State the delivery location." },
  evidence: [],
  limitations: ["No matter source states the delivery location."],
  material: true,
  lawyer: true,
}
const SPINE_FACT = { kind: "fact", detail: "The supplier shall deliver the goods within thirty days of the order date.", evidence: [{ source_id: SOURCE, quote: "within thirty days of the order date" }], limitations: [] }
const SPINE_OPEN = { kind: "open-item", detail: "The delivery location is not recorded; the note must state it as an open item.", evidence: [], limitations: ["No matter source states the delivery location."] }
const OPEN_DRAFT = `# Supply note\n\n## Delivery\n\nThe supplier shall deliver the goods within thirty days of the order date. Deliver to ${LOCATION_MARKER}.\n\n## Payment\n\nPayment is due on delivery.\n`

/** The discrepancy jobs told apart by session, so a retry answers as its job:
 * GAPS lists the location gap, its VERIFY finds no source supplies it, and
 * DIFFERENCES finds none. Each scripted model needs its own factory: session
 * ids repeat across models. */
function locationDiscrepancy(): Reply {
  const jobs = new Map<string | undefined, "gaps" | "verify" | "differences">()
  return (prompt, _index, sessionID) => {
    if (!jobs.has(sessionID)) jobs.set(sessionID, prompt.includes('Return JSON {"gaps"') ? "gaps" : prompt.includes('Return JSON {"supplied"') ? "verify" : "differences")
    switch (jobs.get(sessionID)) {
      case "gaps": return JSON.stringify({ gaps: [LOCATION_GAP], limitations: [] })
      case "verify": return JSON.stringify({ supplied: false, evidence: [], limitations: [] })
      default: return JSON.stringify({ entries: [], limitations: [] })
    }
  }
}
const locationPlacement: Reply = prompt => JSON.stringify(prompt.includes("<part>\n") ? { entries: [] } : { parts: ["d1/1"] })
const locationNotes: Reply = (prompt, index) => {
  const point = /<open_point>\n([\s\S]*?)\n<\/open_point>/.exec(prompt)?.[1]
  if (point && JSON.parse(point).kind === "placeholder") return LOCATION_NOTE
  return draftingNotes(prompt, index)
}
/** Fresh replies per model, so the discrepancy factory's sessions stay its own. */
const locationReplies = () => ({
  ...quiet,
  [AGENT.draft]: drafter(OPEN_DRAFT),
  [AGENT.discrepancy]: locationDiscrepancy(),
  [AGENT.spine]: () => JSON.stringify({ entries: [SPINE_FACT, SPINE_OPEN], limitations: [] }),
  [AGENT.placement]: locationPlacement,
  [AGENT.finalize]: locationNotes,
})

test("Auto readiness uses the original request and only changes the intake key", () => {
  const directory = matter(SOURCE_TEXT, LOCATION_REQUEST)
  const prompt = () => researchPrompt(directory, "intake", researchBasis(directory))
  const steps = ["discrepancy", "intake", "research", "research-review", "spine"] as const
  const keys = () => steps.map(step => hashStageInputs(step, LEAN_PROMPT_VERSION, stageInputs(directory, step)))
  const offPrompt = prompt()
  const offKeys = keys()
  expect(stageInputs(directory, "intake").map(row => row.name)).not.toContain("drafting-auto-mode")
  setAutoMode(true)
  const onPrompt = prompt()
  expect(onPrompt).toContain(LOCATION_REQUEST)
  expect(onPrompt).not.toBe(offPrompt)
  expect(stageInputs(directory, "intake")).toContainEqual({ name: "drafting-auto-mode", bytes: "review-draft-v2" })
  expect(keys().map((key, index) => key !== offKeys[index])).toEqual([false, true, false, false, false])
  for (const stage of ["research", "research-review"] as const)
    expect(researchPrompt(directory, stage, researchBasis(directory))).not.toContain(AUTO)
  setAutoMode(false)
  expect(prompt()).toBe(offPrompt)
  expect(keys()).toEqual(offKeys)
})

test("with Auto mode on, an unresolved delivery location rides the discrepancy, spine and review to a disclosed open item", async () => {
  const directory = matter(SOURCE_TEXT, LOCATION_REQUEST)
  setAutoMode(true)
  const model = scriptedModel(locationReplies())
  const state = await runLean(directory, { sessions: model.port })
  expect(state.status).toBe("completed")
  expect(readQuestions(directory)).toEqual([])
  // The gap stays unsupplied in the discrepancy record; intake proceeds without asking it.
  expect(readLeanJson(directory, "discrepancy.json")).toEqual({ entries: [], gaps: [LOCATION_GAP], limitations: [] })
  const intake = model.prompts.find(row => row.agent === AGENT.intake)!.text
  expect(intake).toContain(AUTO)
  // The spine carries the gap as an open-item entry beside the settled fact, both placed.
  expect((readLeanJson(directory, "spine.json") as { entries: { kind: string }[] }).entries.map(row => row.kind)).toEqual(["fact", "open-item"])
  expect(readLeanJson(directory, "placement.json")).toMatchObject({ status: "placed", issues: [] })
  // The body marks the location; the gate discloses it as a review draft, not a pass to execute.
  const gate = readLeanJson(directory, "gate.json") as GateResult
  expect(gate).toMatchObject({ status: "pass", delivery: "review_draft", blocking: [] })
  expect(gate.disclosed.map(item => item.kind)).toEqual(["placeholder"])
  const item = gate.disclosed[0]!
  expect(item.ref).toMatchObject({ document: "d1", placeholder: LOCATION_MARKER })
  // The notes name the open item the code references.
  const written = readFileSync(leanFile(directory, "drafting-notes.md"), "utf8")
  expect(written).toContain(`${LOCATION_NOTE} [ref: ${item.id}]`)
}, 120_000)

test("with Auto mode off, the same unresolved location is asked as a fact question and blocks before any plan", async () => {
  const directory = matter(SOURCE_TEXT, LOCATION_REQUEST)
  const model = scriptedModel({ ...locationReplies(), [AGENT.intake]: () => JSON.stringify({ questions: [
    { question: "What is the delivery location?", kind: "fact", context: "The supply terms do not record the delivery location the note must state." },
  ] }) })
  const state = await runLean(directory, { sessions: model.port })
  expect(state.status).toBe("blocked")
  // The gap is recorded the same way; only intake's decision differs under the base rule.
  expect(readLeanJson(directory, "discrepancy.json")).toEqual({ entries: [], gaps: [LOCATION_GAP], limitations: [] })
  const asked = readQuestions(directory)
  expect(asked).toHaveLength(1)
  expect(asked[0]).toMatchObject({ status: "open", kind: "fact" })
  expect(asked[0]!.options?.map(option => option.fact_action)).toEqual(["provisional", "wait"])
  expect(readLeanText(directory, "plan.json")).toBeUndefined()
  const intake = model.prompts.find(row => row.agent === AGENT.intake)!.text
  expect(intake).not.toContain(AUTO)
}, 120_000)

test("with Auto mode on, an explicit hold still asks and blocks before any plan or draft", async () => {
  const directory = matter()
  setAutoMode(true)
  const model = scriptedModel({ [AGENT.intake]: () => JSON.stringify({ questions: [
    { question: "Should the draft wait for the client's approval of the approach before proceeding?", kind: "instruction", context: "The lawyer's approval is outstanding." },
  ] }) })
  const state = await runLean(directory, { sessions: model.port })
  expect(state.status).toBe("blocked")
  expect(readQuestions(directory)).toHaveLength(1)
  expect(readQuestions(directory)[0]).toMatchObject({ status: "open", kind: "instruction" })
  expect(readLeanText(directory, "plan.json")).toBeUndefined()
}, 120_000)

test("with Auto mode on, a missing fact still asks with the provisional and wait choices, and its advice carries the guidance", async () => {
  const directory = matter()
  setAutoMode(true)
  const model = scriptedModel({ [AGENT.intake]: () => JSON.stringify({ questions: [
    { question: "What is the delivery location?", kind: "fact", context: "The source omits the location." },
  ] }) })
  const state = await runLean(directory, { sessions: model.port })
  expect(state.status).toBe("blocked")
  const asked = readQuestions(directory)[0]!
  expect(asked).toMatchObject({ status: "open", kind: "fact" })
  expect(asked.options?.map(option => option.fact_action)).toEqual(["provisional", "wait"])
  const advice = model.prompts.find(row => row.agent === AGENT.advise)!.text
  expect(advice).toContain(AUTO)
}, 120_000)

test("with Auto mode on, an answered wait still holds the run without any model call", async () => {
  const directory = matter()
  setAutoMode(true)
  writeQuestions(directory, [{ id: "q-delivery", stage: "intake", kind: "fact", question: "What is the delivery location?", context: "The source omits the location.",
    status: "answered", answer: "Wait for me to provide the location.", fact_status: "wait", answer_class: "answered" } as Question])
  const model = scriptedModel(quiet)
  const state = await runLean(directory, { sessions: model.port })
  expect(state.status).toBe("blocked")
  expect(state.stop_reason).toContain("evidence")
  expect(model.prompts).toHaveLength(0)
  expect(readQuestions(directory)).toHaveLength(1)
}, 120_000)

test("enabling Auto mode while blocked does not retract the asked question; after the answer, intake re-runs with the guidance", async () => {
  const directory = matter()
  const hold = scriptedModel({ [AGENT.intake]: () => JSON.stringify({ questions: [
    { question: "Should the draft wait for the client's approval of the approach before proceeding?", kind: "instruction", context: "The lawyer's approval is outstanding." },
  ] }) })
  expect((await runLean(directory, { sessions: hold.port })).status).toBe("blocked")
  const asked = readQuestions(directory)[0]!
  setAutoMode(true)
  enqueueLeanRequest(directory, { kind: "continue" })
  const still = scriptedModel(quiet)
  const state = await runLean(directory, { sessions: still.port })
  expect(state.status).toBe("blocked")
  expect(state.stop_reason).toContain(asked.id)
  expect(still.prompts).toHaveLength(0)
  expect(readQuestions(directory)).toEqual([asked])
  enqueueLeanRequest(directory, { kind: "answer", question_id: asked.id, answer: "Proceed; the missing contact detail may remain a placeholder." })
  const resumed = scriptedModel(quiet)
  expect((await runLean(directory, { sessions: resumed.port })).status).toBe("completed")
  expect(resumed.prompts.find(row => row.agent === AGENT.intake)!.text).toContain(AUTO)
  expect(readQuestions(directory)[0]).toMatchObject({ id: asked.id, status: "answered" })
}, 120_000)

test("advice gains only the Auto mode line for an intake fact; other questions keep the base prompt", () => {
  const directory = matter()
  const fact: Question = { id: "q-location", stage: "intake", kind: "fact", question: "What is the delivery location?", context: "The source omits the location.", status: "open" }
  const off = advisePrompt(directory, fact)
  expect(off).not.toContain(AUTO)
  setAutoMode(true)
  const on = advisePrompt(directory, fact)
  const start = on.indexOf(AUTO)
  const line = on.slice(start, on.indexOf("\n\n<question>"))
  expect(on.replace(`\n\n${line}`, "")).toBe(off)
  expect(advisePrompt(directory, { ...fact, kind: "instruction" })).not.toContain(AUTO)
  expect(advisePrompt(directory, { ...fact, stage: "draft" })).not.toContain(AUTO)
  setAutoMode(false)
  expect(advisePrompt(directory, fact)).toBe(off)
})
