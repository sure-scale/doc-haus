import { afterEach, expect, test } from "bun:test"
import { readFileSync } from "node:fs"
import { renderMarkdownToDocx } from "../../lib/render"
import { readDocxCommentParagraphs } from "../../lib/docx-comments"
import { checklistRepairItems, readRepairRound, LEAN_MAX_SUBSTANTIVE_REPAIRS } from "./repairs"
import { checklistRefinementSourceBasis } from "./checklist-flow"
import { repairBodyText } from "./change-pipeline"
import { changesRepairSections, startRepairRound } from "./repairs"
import { LEAN_PROMPT_VERSION, checklistOutputState, stageInputs, stageOutputValid } from "./stages"
import { stageInputSha } from "./runner"
import { drainLeanRequests, enqueueLeanRequest } from "./requests"
import { writeLeanPlan } from "./plan"
import { hashStageInputs, leanFile, readLeanState, writeLeanBytes, writeLeanText, writeQuestions, type LeanState } from "./state"
import { cleanupMatters, matter, SOURCE, SUPPORT } from "./test-support"
import type { ChecklistItem, ChecklistFinding } from "./checklist-contract"

afterEach(cleanupMatters)
const item: ChecklistItem = { check: "Preserve supplied delivery timing.", basis: "directive", qualifications: [],
  source_refs: [{ document: SOURCE, locator: "source paragraph", quote: SUPPORT[0]!.quote }],
  target: { document: "advice.docx", role: "Client advice", treatment: "advice" } }
async function setup() {
  const directory = matter()
  writeLeanPlan(directory, [
    { title: "Record", purpose: "Record supplied terms", file: "record.docx", parts: [{ heading: "Record", scope: "Record" }] },
    { title: "Advice", purpose: "Explain supplied terms", file: "advice.docx", parts: [{ heading: "Advice", scope: "Advice" }] },
  ])
  await writeLeanBytes(directory, "d2/body-applied.docx", (await renderMarkdownToDocx("Timing is missing.\n\nA separate unchanged provision.")).bytes)
  const paragraph = readDocxCommentParagraphs(readFileSync(leanFile(directory, "d2/body-applied.docx")))[0]!
  const initialPID = "p111111111111"
  writeLeanText(directory, "d2/paragraphs.json", JSON.stringify([{ pid: initialPID, sha256: "initial", text: "The source timing was present in the initial draft." }]))
  writeLeanText(directory, "d2/apply.json", JSON.stringify({ applied: [{ id: "prior", pid: initialPID, result_pids: paragraph.pids }], unapplied: [] }))
  const finding: ChecklistFinding = { item_id: "parent", status: "unsatisfied", reason: "Supplied timing is missing from the current text.",
    evidence: [{ document: "d2", locator: `native PIDs ${paragraph.pids.join(",")}`, text: paragraph.paragraph_text, kind: "body" }],
    source_refs: item.source_refs, inspection_scope: "Current affected paragraph and source.", limitation: "No extra approval inferred." }
  writeLeanText(directory, "checklist/extraction.json", JSON.stringify({ expected_sources: ["request"], sources: [{ key: "request", status: "succeeded", input_sha: "input", items: [{ id: "parent", item }], issues: [] }] }))
  const save = (value: ChecklistFinding) => writeLeanText(directory, "checklist/checks.json", JSON.stringify({ extraction_complete: true, expected_items: ["parent"], checks: [{ item_id: "parent", status: "succeeded", input_sha: "input", finding: value, issues: [] }] }))
  save(finding)
  return { directory, finding, save, initialPID }
}

test("effective unsatisfied finding carries exact source/item/current evidence with d2 and split lineage ownership", async () => {
  const { directory, finding, initialPID } = await setup()
  const repairs = checklistRepairItems(directory)
  expect(repairs).toHaveLength(1)
  expect(repairs[0]!.ref).toEqual({ document: "d2", repair_target: "body", pid: initialPID })
  const packet = JSON.parse(repairs[0]!.detail.split("\n")[1]!)
  expect(packet.item).toEqual(item)
  expect(packet.finding).toEqual(finding)
  expect(packet.current_target_paragraphs[0].paragraph_text).toBe("Timing is missing.")
  expect(repairs[0]!.detail).toContain("adverse verdict is not authority")
})

test("indeterminate, undeclared and ambiguous targets do not invent a body owner", async () => {
  const { directory, finding, save } = await setup()
  save({ ...finding, status: "indeterminate" })
  expect(checklistRepairItems(directory)).toEqual([])
  save(finding)
  const extraction = JSON.parse(readFileSync(leanFile(directory, "checklist/extraction.json"), "utf8"))
  extraction.sources[0].items[0].item.target.document = "not-planned.docx"
  writeLeanText(directory, "checklist/extraction.json", JSON.stringify(extraction))
  expect(checklistRepairItems(directory)).toEqual([])
})

test("archived original adverse findings are not effective repair targets", async () => {
  const { directory, finding } = await setup()
  const refined = { ...item, required_action: "flag" as const }
  writeLeanText(directory, "checklist/refinements.json", JSON.stringify({ parents: [{ parent_id: "parent", parent: item, scope: "cited_sources", source_basis: checklistRefinementSourceBasis(directory), expected_units: ["unit"], units: [{ unit_id: "unit", input_sha: "unit", status: "succeeded", result: { parent_id: "parent", unit_id: "unit", disposition: "refined", reason: "Focused check.", items: [refined], source_refs: item.source_refs, limitations: [] }, children: [{ id: "child", item: refined }], issues: [] }] }] }))
  writeLeanText(directory, "checklist/checks.json", JSON.stringify({ extraction_complete: true, expected_items: ["child"], checks: [
    { item_id: "parent", status: "succeeded", input_sha: "old", finding, issues: [] },
    { item_id: "child", status: "succeeded", input_sha: "new", finding: { ...finding, item_id: "child", status: "satisfied", action_status: "fulfilled", open_decisions: [] }, issues: [] },
  ] }))
  expect(checklistRepairItems(directory)).toEqual([])
})

test("explicit Continue admits one substantive round and refuses another after the existing ceiling", async () => {
  const { directory } = await setup()
  const state = readLeanState(directory)!
  state.status = "blocked"
  state.stages["checklist-check"] = { ...state.stages["checklist-check"]!, status: "blocked", input_sha: stageInputSha(directory, "checklist-check") }
  expect(readRepairRound(directory)).toBeUndefined()
  enqueueLeanRequest(directory, { kind: "continue" })
  expect(drainLeanRequests(directory, state).repairRound).toBe(true)
  expect(readRepairRound(directory)?.substantive_attempts).toBe(1)
  const exhausted = { ...readRepairRound(directory)!, substantive_attempts: LEAN_MAX_SUBSTANTIVE_REPAIRS }
  writeLeanText(directory, "repairs.json", JSON.stringify(exhausted))
  enqueueLeanRequest(directory, { kind: "continue" })
  expect(drainLeanRequests(directory, state).repairRound).toBe(false)
  expect(readRepairRound(directory)).toEqual(exhausted)
})

for (const change of ["Word", "answers"] as const) test(`Continue reinspects stale ${change} evidence without admitting a body repair`, async () => {
  const { directory } = await setup()
  const state = readLeanState(directory)!
  state.status = "blocked"
  state.stages["checklist-check"] = { ...state.stages["checklist-check"]!, status: "blocked", input_sha: stageInputSha(directory, "checklist-check") }
  if (change === "Word") await writeLeanBytes(directory, "d2/body-applied.docx", (await renderMarkdownToDocx("Timing is missing.\n\nNew current context.")).bytes)
  else writeLeanText(directory, "updated-notes.md", "A later instruction changes the applicable timing.")
  enqueueLeanRequest(directory, { kind: "continue" })
  expect(drainLeanRequests(directory, state).repairRound).toBe(false)
  expect(readRepairRound(directory)).toBeUndefined()
})

test("an indeterminate checklist does not revive a historical blocked gate", async () => {
  const { directory, finding, save } = await setup()
  save({ ...finding, status: "indeterminate" })
  const state = readLeanState(directory)!
  state.status = "blocked"
  state.stages["checklist-check"] = { ...state.stages["checklist-check"]!, status: "blocked", input_sha: stageInputSha(directory, "checklist-check") }
  writeLeanText(directory, "gate.json", JSON.stringify({ status: "blocked", blocking: [{ id: "old", stage: "gate", kind: "omission", detail: "Historical defect", ref: { document: "d2", repair_target: "body" } }] }))
  enqueueLeanRequest(directory, { kind: "continue" })
  expect(drainLeanRequests(directory, state).repairRound).toBe(false)
  expect(readRepairRound(directory)).toBeUndefined()
})

test("checklist adjudication receives affected actual Word paragraphs without an unrelated tail or established-defect instruction", async () => {
  const { directory, finding, save } = await setup()
  await writeLeanBytes(directory, "d2/body-applied.docx", (await renderMarkdownToDocx("Timing is missing.\n\nImmediate neighbor.\n\nUnrelated intermediate.\n\nUnrelated tail is excluded.")).bytes)
  const state = readLeanState(directory)!
  state.stages["d2.apply"] = { ...state.stages["d2.apply"]!, status: "done", input_sha: hashStageInputs("d2.apply", LEAN_PROMPT_VERSION, stageInputs(directory, "d2.apply")) }
  writeLeanText(directory, "state.json", JSON.stringify(state))
  const row = readDocxCommentParagraphs(readFileSync(leanFile(directory, "d2/body-applied.docx")))[0]!
  save({ ...finding, evidence: [{ document: "d2", locator: row.pids.join(","), text: row.paragraph_text, kind: "body" }] })
  startRepairRound(directory, checklistRepairItems(directory), true)
  const context = repairBodyText(directory, "d2", false, true)
  expect(context).toContain("Timing is missing.")
  expect(context).toContain("Immediate neighbor.")
  expect(context).not.toContain("Unrelated tail")
  expect(context).toContain("view=current_body")
  const sections = changesRepairSections(directory, "d2").join("\n")
  expect(sections).toContain("<checklist_findings>")
  expect(sections).not.toContain("Each item is a defect")
})


for (const action_status of ["not_fulfilled", "fulfilled", "indeterminate", undefined] as const) test(`explicit Continue routes refined open/${action_status ?? "missing"} to source adjudication only for established unfulfillment`, async () => {
  const { directory, finding, save } = await setup()
  const extraction = JSON.parse(readFileSync(leanFile(directory, "checklist/extraction.json"), "utf8"))
  extraction.sources[0].items[0].item.required_action = "propose"
  writeLeanText(directory, "checklist/extraction.json", JSON.stringify(extraction))
  save({ ...finding, status: "open", action_status, open_decisions: ["Approval"] })
  const state = readLeanState(directory)!
  state.status = "blocked"
  state.stages["checklist-check"] = { ...state.stages["checklist-check"]!, status: "blocked", input_sha: stageInputSha(directory, "checklist-check") }
  enqueueLeanRequest(directory, { kind: "continue" })
  expect(drainLeanRequests(directory, state).repairRound).toBe(action_status === "not_fulfilled")
  expect(readRepairRound(directory)?.substantive_attempts).toBe(action_status === "not_fulfilled" ? 1 : undefined)
  if (action_status === "not_fulfilled") {
    const packet = JSON.parse(readRepairRound(directory)!.items[0]!.detail.split("\n")[1]!)
    expect(packet.item.required_action).toBe("propose")
    expect(packet.finding.action_status).toBe("not_fulfilled")
    expect(packet.finding.status).toBe("open")
    expect(readRepairRound(directory)!.items[0]!.detail).toContain("Preserve proposals as proposals")
    writeLeanText(directory, "repairs.json", JSON.stringify({ ...readRepairRound(directory)!, substantive_attempts: LEAN_MAX_SUBSTANTIVE_REPAIRS }))
    enqueueLeanRequest(directory, { kind: "continue" })
    expect(drainLeanRequests(directory, state).repairRound).toBe(false)
    expect(readRepairRound(directory)!.substantive_attempts).toBe(LEAN_MAX_SUBSTANTIVE_REPAIRS)
  }
})


test("completed checklist validity cannot hide an unfulfilled or inconsistent required action behind open or satisfied", async () => {
  const { directory, finding, save } = await setup()
  const extraction = JSON.parse(readFileSync(leanFile(directory, "checklist/extraction.json"), "utf8"))
  extraction.sources[0].items[0].item.required_action = "propose"
  writeLeanText(directory, "checklist/extraction.json", JSON.stringify(extraction))
  writeLeanText(directory, "checklist/report.md", "Current findings")
  // Complete output is valid; the unfulfilled action is returned as an
  // unresolved finding for disclosure, never as an open or satisfied one.
  for (const status of ["open", "satisfied"] as const) {
    save({ ...finding, status, action_status: "not_fulfilled", open_decisions: ["Approval"] })
    expect(checklistOutputState(directory)).toMatchObject({ valid: true, openFindings: [], unresolved: [{ item_id: "parent", class: "omission", finding: { status } }] })
  }
  save({ ...finding, status: "open", action_status: "fulfilled", open_decisions: ["Approval"] })
  expect(stageOutputValid(directory, "checklist-check")).toBe(true)
  expect(checklistOutputState(directory)).toMatchObject({ valid: true, open: true, openFindings: [{ item_id: "parent" }], unresolved: [] })
})

// The focused adjudicator's exact-item disposition of the last admitted round.
function adjudicate(directory: string, item_id: string, options: { groups?: unknown[]; cited?: boolean } = {}) {
  writeLeanText(directory, "d2/repair-adjudications.json", JSON.stringify({ assessments: [{ item_id, value: { shared_decisions: [], groups: options.groups ?? [], dismissed: [], questions: [],
    findings: [{ disposition: options.cited === false ? "deferred" : "rejected", kind: "information", detail: "Decline the assessed timing repair.", reason: "The supplied terms already govern timing.", documents: ["d2"],
      sources: options.cited === false ? [] : [{ source_id: SOURCE, quote: SUPPORT[0]!.quote }], limitations: options.cited === false ? ["No source passage located."] : [] }],
    superseded_findings: [], superseded_decisions: [] } }] }))
}
function settled(directory: string, outputs = ["d2/change-plan.json", "d2/repair-adjudications.json"]): LeanState {
  const state = readLeanState(directory)!
  state.status = "blocked"
  state.stages["checklist-check"] = { ...state.stages["checklist-check"]!, status: "blocked", input_sha: stageInputSha(directory, "checklist-check") }
  state.stages["d2.change-plan"] = { status: "done", input_sha: "accepted", outputs, session_id: null, attempts: 1, calls: 1, feedback: [], started_at: null, finished_at: null }
  return state
}
async function declinedRound() {
  const fixture = await setup()
  startRepairRound(fixture.directory, checklistRepairItems(fixture.directory), true)
  return { ...fixture, round: readRepairRound(fixture.directory)! }
}

test("Continue resumes instead of resubmitting a repair the current adjudication declined for the exact item", async () => {
  const { directory, round } = await declinedRound()
  adjudicate(directory, "parent")
  enqueueLeanRequest(directory, { kind: "continue" })
  expect(drainLeanRequests(directory, settled(directory)).repairRound).toBe(false)
  expect(readRepairRound(directory)).toEqual(round)
  expect(round.substantive_attempts).toBe(1)
})

for (const variant of ["another item", "a correction group", "a superseded plan", "an uncited deferral"] as const) test(`Continue admits the repair round as before when the adjudication has ${variant}`, async () => {
  const { directory, initialPID } = await declinedRound()
  adjudicate(directory, variant === "another item" ? "other" : "parent", {
    ...(variant === "a correction group" ? { groups: [{ id: "timing", pids: [initialPID], goal: "Restore the supplied timing.", sources: SUPPORT, limitations: [] }] } : {}),
    ...(variant === "an uncited deferral" ? { cited: false } : {}) })
  enqueueLeanRequest(directory, { kind: "continue" })
  expect(drainLeanRequests(directory, settled(directory, variant === "a superseded plan" ? ["d2/change-plan.json"] : undefined)).repairRound).toBe(true)
  expect(readRepairRound(directory)).toMatchObject({ round: 2, substantive_attempts: 2 })
})

for (const kind of ["clarification", "update", "answer"] as const) test(`${kind === "clarification" ? "a" : "an"} ${kind} accepted after the declined round lets Continue resubmit the repair`, async () => {
  const { directory } = await declinedRound()
  adjudicate(directory, "parent")
  if (kind === "answer") writeQuestions(directory, [{ id: "q1", stage: "change-plan", question: "Does the supplied timing apply to the advice?", context: "Timing", status: "open" }])
  await Bun.sleep(5)
  const text = "The supplied timing applies to the advice."
  enqueueLeanRequest(directory, kind === "answer" ? { kind, question_id: "q1", answer: text } : { kind, text })
  expect(drainLeanRequests(directory, settled(directory))).toMatchObject({ repairRound: false })
  // The checklist stage settles again on the new inputs before the Continue.
  enqueueLeanRequest(directory, { kind: "continue" })
  expect(drainLeanRequests(directory, settled(directory)).repairRound).toBe(true)
  expect(readRepairRound(directory)).toMatchObject({ round: 2, substantive_attempts: 2 })
})
