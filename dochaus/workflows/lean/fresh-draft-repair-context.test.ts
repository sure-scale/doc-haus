import { renderMarkdownToDocx } from "../../lib/render"
import { stagePrompt } from "./prompts"
import { afterEach, expect, test } from "bun:test"
import { freshDraftRepairInputs, reviewSections } from "./fresh-draft-repair-context"
import { REPAIR_ADJUDICATION_OUTPUT, repairAdjudicationIssues, mergeRepairAdjudications, runFocusedRepairAdjudication, type Assessment, type RepairAdjudication } from "./checklist-repair-adjudication"
import { readChangePlan } from "./change-pipeline"
import { reconciliationWork } from "./change-reconciliation"
import { reviewStages, writeLeanPlan } from "./plan"
import { readLeanState, readLeanText, updateLeanStage, writeLeanBytes, writeLeanText } from "./state"
import { openItem, quoteFound } from "./quotes"
import { readRepairRound, startRepairRound } from "./repairs"
import { cleanupMatters, matter, scriptedModel, SOURCE, SUPPORT } from "./test-support"

afterEach(cleanupMatters)
const old = "p111111111111", current = "p222222222222"
function setup() {
  const directory = matter()
  writeLeanPlan(directory, [{ title: "Advice", purpose: "Explain timing", file: "advice.docx", parts: [{ heading: "Timing", scope: "Timing" }] }])
  writeLeanText(directory, "d1/paragraphs.json", JSON.stringify([{ pid: current, sha256: "current", text: "Timing needs clarification." }]))
  updateLeanStage(directory, readLeanState(directory)!, "d1.draft", { status: "done" })
  for (const stage of reviewStages(directory, "d1")) {
    writeLeanText(directory, `d1/${stage}.md`, `# ${stage}\n\nCURRENT_${stage}: Reassess timing against the source.`)
    updateLeanStage(directory, readLeanState(directory)!, `d1.${stage}`, { status: "done" })
  }
  writeLeanText(directory, "d1/change-plan.json", JSON.stringify({ shared_decisions: ["Earlier generated timing interpretation."], groups: [{ id: "prior", pids: [old], goal: "Retain delivery timing.", sources: SUPPORT, limitations: [] }], findings: [{ disposition: "deferred", kind: "information", detail: "Prior unresolved issue", reason: "Unknown fact", documents: ["d1"], sources: [], limitations: ["Unanswered"] }], dismissed: [], questions: [] }))
  writeLeanText(directory, "d1/changes.json", JSON.stringify({ amendments: [{ id: "earlier", pid: old, existing_sha256: "old", updated: "Earlier supported wording.", reason: "Timing correction", sources: SUPPORT }], withheld: [{ amendment: { id: "withheld", pid: old, existing_sha256: "old", updated: "WITHHELD_INTENT", reason: "Earlier proposed correction", sources: SUPPORT }, issues: ["Not applied"] }], explanations: [], questions: [] }))
  const assessment: Assessment = { item_id: "timing", item: { check: "Explain timing", basis: "directive", source_refs: [{ document: SOURCE, locator: "source", quote: SUPPORT[0]!.quote }], qualifications: ["Do not assume approval"], target: { document: "d1", role: "Advice", treatment: "advice" } }, finding: { item_id: "timing", status: "unsatisfied", reason: "Timing missing from earlier draft", evidence: [{ document: "d1", locator: old, kind: "body", text: "Earlier stale wording" }], source_refs: [], inspection_scope: "Earlier body", limitation: "Needs current inspection" } }
  writeLeanText(directory, "checklist/checks.json", JSON.stringify({ checks: [assessment] }))
  startRepairRound(directory, [openItem("checklist-check", "omission", `Reassess.\n${JSON.stringify(assessment)}`, { document: "d1", repair_target: "body", pid: old })], true)
  return { directory, assessment }
}
const disposition: RepairAdjudication = { shared_decisions: [], groups: [], findings: [{ disposition: "deferred", kind: "information", detail: "Timing source reassessment", reason: "Current support needs inspection", documents: ["d1"], sources: [], limitations: ["Unresolved"] }], dismissed: [], questions: [], superseded_decisions: [], superseded_findings: [] }

test("new draft accounts for every current review and prior correction without accepting historical targets", () => {
  const { directory, assessment } = setup()
  const before = readLeanText(directory, "d1/changes.json")
  const context = freshDraftRepairInputs(directory, "d1", [assessment], REPAIR_ADJUDICATION_OUTPUT)
  expect(context.previous.groups).toEqual([])
  const prompts = context.inputs.map(row => row.prompt).join("\n")
  for (const stage of reviewStages(directory, "d1")) expect(prompts).toContain(`CURRENT_${stage}`)
  expect(prompts).toContain("Earlier supported wording.")
  expect(prompts).toContain("WITHHELD_INTENT")
  expect(prompts).toContain("Prior unresolved issue")
  expect(prompts).toContain("Earlier generated timing interpretation.")
  expect(prompts).toContain("Do not assume approval")
  expect(prompts).not.toContain("Earlier stale wording")
  expect(prompts).toContain("/checks/0/finding")
  expect(context.inputs.filter(row => row.assessment.item_id.startsWith("historical-plan:groups"))).toHaveLength(0)
  expect(repairAdjudicationIssues(directory, "d1", context.inputs[0]!, { ...disposition, groups: [{ id: "stale", pids: [old], goal: "Reuse", sources: SUPPORT, limitations: [] }] }).join(" ")).toContain("not an initial draft paragraph")
  expect(mergeRepairAdjudications(context.previous, [disposition]).groups).toEqual([])
  expect(readLeanText(directory, "d1/changes.json")).toBe(before)
})

test("source-backed retained-initial history survives changed PIDs only as a separate reassessment", () => {
  const { directory, assessment } = setup()
  const intent = { pid: current, reason: "Timing: the supplied delivery period supports restoring the initial wording instead of the earlier removal.", sources: SUPPORT }
  const saved = { amendments: [], retained_initial: [intent], explanations: [], withheld: [], questions: [] }
  writeLeanText(directory, "d1/changes.json", JSON.stringify(saved))
  expect(quoteFound(directory, SUPPORT[0]!)).toBe(true)
  const emptyPlan = { shared_decisions: [], groups: [], findings: [], dismissed: [], questions: [] }
  expect(reconciliationWork(directory, "d1", emptyPlan, { groups: [] }).direct.retained_initial).toEqual([intent])

  const fresh = "p333333333333"
  writeLeanText(directory, "d1/paragraphs.json", JSON.stringify([{ pid: fresh, sha256: "fresh", text: "Timing needs clarification." }]))
  expect(reconciliationWork(directory, "d1", emptyPlan, { groups: [] }).direct.retained_initial).toEqual([])
  const before = readLeanText(directory, "d1/changes.json")
  const result = freshDraftRepairInputs(directory, "d1", [assessment], REPAIR_ADJUDICATION_OUTPUT)
  const assignments = result.inputs.filter(row => row.assessment.item_id.startsWith("historical-retained-initial:"))
  expect(assignments).toHaveLength(1)
  const assignment = assignments[0]!
  expect(assignment.prompt).toContain(JSON.stringify(intent))
  expect(assignment.prompt).toContain("/retained_initial/0")
  expect(assignment.prompt).toContain("earlier-draft restoration intent")
  expect(assignment.prompt).toContain("never adopt its old PID as a current target")
  expect(assignment.prompt).toContain(`[${fresh}] Timing needs clarification.`)
  expect(assignment.prompt).not.toContain(`[${current}] Timing needs clarification.`)
  expect(repairAdjudicationIssues(directory, "d1", assignment, { ...disposition, groups: [{ id: "old-target", pids: [current], goal: "Retain wording", sources: SUPPORT, limitations: [] }] }).join(" ")).toContain("not an initial draft paragraph")
  expect(result.previous.groups).toEqual([])
  expect(mergeRepairAdjudications(result.previous, [disposition]).groups).toEqual([])
  expect(readLeanText(directory, "d1/changes.json")).toBe(before)
})

test("oversized retained-initial history fails before admission rather than dropping its rationale", () => {
  const { directory } = setup()
  writeLeanText(directory, "d1/changes.json", JSON.stringify({ amendments: [], questions: [], retained_initial: [{ pid: old, reason: "Historical restoration rationale. ".repeat(250), sources: SUPPORT }] }))
  const before = readLeanText(directory, "d1/changes.json")
  expect(() => freshDraftRepairInputs(directory, "d1", [], REPAIR_ADJUDICATION_OUTPUT)).toThrow("historical-retained-initial:0 exceeds its bounded context")
  expect(readLeanText(directory, "d1/changes.json")).toBe(before)
})

test("new draft native dispatch uses bounded jobs, accepts only freshly adjudicated current targets and preserves budget", async () => {
  const { directory, assessment } = setup()
  await writeLeanBytes(directory, "d1/body.docx", (await renderMarkdownToDocx("Timing needs clarification.")).bytes)
  const budget = readRepairRound(directory)?.substantive_attempts
  const jobs = freshDraftRepairInputs(directory, "d1", [assessment], REPAIR_ADJUDICATION_OUTPUT).inputs.length
  let count = 0
  const model = scriptedModel({ "lean-change-plan": () => JSON.stringify({ ...disposition, findings: [{ ...disposition.findings![0], detail: `Disposition ${count++}` }], groups: [{ id: "current", pids: [current], goal: "Clarify supplied timing", sources: SUPPORT, limitations: [] }] }) })
  const result = await runFocusedRepairAdjudication(directory, readLeanState(directory)!, { sessions: model.port }, "d1", () => {})
  expect(result.status).toBe("done")
  expect(model.prompts).toHaveLength(jobs)
  expect(readChangePlan(directory, "d1")?.groups.flatMap(row => row.pids)).toEqual([current])
  expect(readRepairRound(directory)?.substantive_attempts).toBe(budget)
  const reconciliation = stagePrompt(directory, "d1.changes")
  expect(reconciliation).toContain("Clarify supplied timing")
  expect(reconciliation).not.toContain("<previous_changes>")
  expect(reconciliation).toContain(`[${current}] Timing needs clarification.`)
  expect(reconciliation).not.toContain("Earlier supported wording.")
  expect(reconciliation).not.toContain("WITHHELD_INTENT")
})

test("missing fresh review or corrupt prior plan fails before native sessions and leaves original artifacts", async () => {
  const { directory } = setup()
  updateLeanStage(directory, readLeanState(directory)!, "d1.review", { status: "pending" })
  const before = readLeanText(directory, "d1/change-plan.json")
  const model = scriptedModel({})
  expect((await runFocusedRepairAdjudication(directory, readLeanState(directory)!, { sessions: model.port }, "d1", () => {})).status).toBe("failed")
  expect(model.sessions).toEqual([])
  expect(readLeanText(directory, "d1/change-plan.json")).toBe(before)
  updateLeanStage(directory, readLeanState(directory)!, "d1.review", { status: "done" })
  const raw = JSON.parse(before!)
  raw.groups[0].sources = [{ source_id: SOURCE, quote: "Invented canonical quotation." }]
  writeLeanText(directory, "d1/change-plan.json", JSON.stringify(raw))
  expect(() => freshDraftRepairInputs(directory, "d1", [], REPAIR_ADJUDICATION_OUTPUT)).toThrow("invalid beyond")
})

test("bounded review chunks retain every section and fail explicitly on indivisible oversized text", () => {
  const review = Array.from({ length: 20 }, (_, i) => `## Issue ${i}\n\n${`Unique${i} `.repeat(100)}`).join("\n\n")
  const chunks = reviewSections(review)
  expect(chunks.length).toBeLessThan(20)
  expect(chunks.every(chunk => chunk.length <= 6000)).toBe(true)
  for (let i = 0; i < 20; i++) expect(chunks.join("\n")).toContain(`Unique${i}`)
  expect(() => reviewSections("x".repeat(7000))).toThrow("decompose")
})

test("oversized review tables retain exact rows, repeated column context and surrounding qualifications", () => {
  const header = "| Position | Evidence | Qualification |\n|---|---|---|"
  const rows = Array.from({ length: 24 }, (_, i) => `| Row ${i} | ${`Evidence${i} `.repeat(35)} | Subject to condition ${i}; proposal only. |`)
  const chunks = reviewSections(`## Coverage\n\nBefore: source meaning remains unverified.\n\n${header}\n${rows.join("\n")}\n\nAfter: missing approval remains unknown.`)
  expect(chunks.length).toBeGreaterThan(1)
  expect(chunks.every(chunk => chunk.length <= 6000)).toBe(true)
  const all = chunks.join("\n")
  expect(all).toContain("Before: source meaning remains unverified.")
  expect(all).toContain("After: missing approval remains unknown.")
  for (const row of rows) expect(chunks.flatMap(chunk => chunk.split("\n")).filter(line => line === row)).toHaveLength(1)
  for (const chunk of chunks.filter(chunk => rows.some(row => chunk.includes(row)))) {
    expect(chunk).toContain("## Coverage")
    expect(chunk).toContain(header)
  }
  expect(() => reviewSections(`${header}\n| Single | ${"x".repeat(7000)} | Unresolved |`)).toThrow("table row")
})

test("ordinary multiline review prose and small tables remain unchanged", () => {
  const prose = "## Scope\n\nA qualification continues on\nthe next line without becoming another decision.\n\nPosition | Evidence\n:--- | ---:\nProposal only | First line<br>Second line with an escaped \\| marker."
  expect(reviewSections(prose)).toEqual([prose])
  const header = "Position | Evidence\n:--- | ---:"
  const rows = Array.from({ length: 20 }, (_, i) => `Item ${i} | ${"Evidence<br>Qualified \\| text. ".repeat(20)}`)
  const chunks = reviewSections(`## Scope\n\n${header}\n${rows.join("\n")}`)
  expect(chunks.every(chunk => chunk.length <= 6000 && chunk.includes(header))).toBe(true)
  for (const row of rows) expect(chunks.flatMap(chunk => chunk.split("\n")).filter(line => line === row)).toHaveLength(1)
})
