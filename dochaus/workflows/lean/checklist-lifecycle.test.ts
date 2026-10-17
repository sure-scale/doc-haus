import { afterEach, expect, setSystemTime, test } from "bun:test"
import { existsSync, readdirSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import path from "node:path"
import { strFromU8, unzipSync } from "fflate"
import { renderMarkdownToDocx } from "../../lib/render"
import { runLean, stageInputSha } from "./runner"
import { stageOutputValid, type StageSessionPort } from "./stages"
import { exportDocxComments, readDocxCommentParagraphs } from "../../lib/docx-comments"
import { leanFile, readLeanJson, readLeanState, readOpenItems, updateLeanStage, writeLeanBytes, writeLeanState, writeLeanText } from "./state"
import { finalGate, type GateResult } from "./gate"
import { readLeanDocuments } from "./deliverables"
import { enqueueLeanRequest } from "./requests"
import { CHECKLIST_JOB_AGENTS } from "./checklist-job"
import { readRepairRound } from "./repairs"
import { AGENT, cleanupMatters, happyPath, matter, registering, scriptedModel, drafter, pidOf, DRAFT, SOURCE, SUPPORT } from "./test-support"
import { checklistTextMeasuresMeaning, type ChecklistCheckerInput, type ChecklistProducerInput } from "./checklist-contract"

afterEach(cleanupMatters)

// The focused repair adjudicator defers the supplied assessment without a
// supported correction; every other change plan proposes nothing.
const deferChecklistRepair = (text: string) => JSON.stringify({ shared_decisions: [], groups: [], dismissed: [], questions: [],
  ...(text.includes("ONE supplied checklist assessment") ? { superseded_findings: [], superseded_decisions: [], findings: [{ disposition: "deferred", kind: "information",
    detail: "The assessed requirement still needs a supported correction.", reason: "No supported correction is identified for this assessment.", documents: ["d1"], sources: [], limitations: ["No supported correction identified."] }] } : {}) })

// `broken` adds a second requirement whose check job always fails: an integrity
// failure, so the checklist stage blocks. `second` adds it satisfied instead,
// and `pair` with the same result as the first.
// `departure` cites current draft text as a deliberate departure.
function model(directory: string, status: "satisfied" | "open" | "unsatisfied" | "indeterminate" = "satisfied", extra: { open_decisions?: string[]; indeterminate_cause?: "missing_inspection" | "not_inspectable" } = {}, options: { broken?: boolean; second?: boolean; pair?: boolean; departure?: boolean } = {}) {
  let extractedBeforeBody = false
  let checkedAfterNotes = false
  const result = scriptedModel({
    ...happyPath,
    [AGENT["change-plan"]]: deferChecklistRepair,
    [CHECKLIST_JOB_AGENTS.extract]: text => {
      const input = JSON.parse(text.split("SOURCE_INPUT:\n\n")[1]!) as ChecklistProducerInput
      extractedBeforeBody = !existsSync(leanFile(directory, "d1/body.docx"))
      const item = { check: "The requested supply note is drafted.", basis: "directive", source_refs: input.passages, qualifications: [], target: { ...input.targets[0], treatment: "operative" } }
      return JSON.stringify({ items: input.source.document === "request" ? [item, ...(options.broken || options.second || options.pair ? [{ ...item, check: "Payment remains due on delivery." }] : [])] : [] })
    },
    [CHECKLIST_JOB_AGENTS.check]: text => {
      const input = JSON.parse(text.split("CHECK_INPUT:\n\n")[1]!) as ChecklistCheckerInput
      if (options.broken && input.item.check.startsWith("Payment")) throw new Error("Scripted read-only inspection failure")
      if (options.second && input.item.check.startsWith("Payment")) return JSON.stringify({ item_id: input.item_id, status: "satisfied", reason: "Focused result satisfied.",
        evidence: input.current_evidence, source_refs: input.item.source_refs, inspection_scope: "Current body paragraphs only.", limitation: "Comments not inspected." })
      checkedAfterNotes = existsSync(leanFile(directory, "drafting-notes.docx")) && existsSync(leanFile(directory, "d1/body-applied.docx"))
      return JSON.stringify({ item_id: input.item_id, status, reason: `Focused result ${status}.`, evidence: input.current_evidence,
        source_refs: input.item.source_refs, inspection_scope: "Current body paragraphs only.", limitation: "Comments not inspected.", ...extra,
        ...(options.departure ? { departure_disclosure: input.current_evidence.slice(0, 1) } : {}) })
    },
  })
  return { ...result, early: () => extractedBeforeBody, late: () => checkedAfterNotes }
}

const checklistIDs = (directory: string) => (readLeanJson(directory, "checklist/extraction.json") as any).sources.flatMap((row: any) => row.items).map((row: any) => row.id) as string[]
const focusedRepairs = (scripted: ReturnType<typeof scriptedModel>) => scripted.prompts.filter(row => row.agent === AGENT["change-plan"] && row.text.includes("ONE supplied checklist assessment")).length
const report = (directory: string) => readFileSync(leanFile(directory, "checklist/report.md"), "utf8")
const checked = (run: ReturnType<typeof scriptedModel>, id: string) => run.prompts.filter(row => row.agent === CHECKLIST_JOB_AGENTS.check && (JSON.parse(row.text.split("CHECK_INPUT:\n\n")[1]!) as ChecklistCheckerInput).item_id === id).length

// A run settled blocked at the checklist: by an integrity failure, or before
// adverse findings were disclosed. Continue, refinement and clarification
// requests are admitted only at such a stop point.
function settleBlockedChecklist(directory: string) {
  const state = readLeanState(directory)!
  writeLeanState(directory, { ...state, status: "blocked", stop_reason: "Checklist handoff blocked. See checklist/report.md.",
    stages: { ...state.stages, "checklist-check": { ...state.stages["checklist-check"]!, status: "blocked" } } })
}

test("coordinator extracts before body drafting and checks after body and notes; unchanged resume makes no calls", async () => {
  const directory = matter()
  const scripted = model(directory)
  expect((await runLean(directory, { sessions: scripted.port })).status).toBe("completed")
  expect(readLeanJson(directory, "gate.json")).toMatchObject({ status: "pass", delivery: "execution_ready" })
  expect(scripted.early()).toBe(true)
  expect(scripted.late()).toBe(true)
  expect(readLeanJson(directory, "checklist/checks.json")).toMatchObject({ extraction_complete: true, checks: [{ finding: { status: "satisfied" } }] })
  const idle = scriptedModel({})
  expect((await runLean(directory, { sessions: idle.port })).status).toBe("completed")
  expect(idle.prompts).toHaveLength(0)
  // A valid external current Word change invalidates late checks even when
  // the draft and finalize stage keys remain otherwise unchanged.
  writeLeanBytes(directory, "drafting-notes.docx", (await renderMarkdownToDocx("# Current notes\n\nChanged current Word content.")).bytes)
  const refreshed = model(directory)
  expect((await runLean(directory, { sessions: refreshed.port })).status).toBe("completed")
  expect(refreshed.count(CHECKLIST_JOB_AGENTS.check)).toBe(1)
  expect(refreshed.count(CHECKLIST_JOB_AGENTS.extract)).toBe(0)
  expect(refreshed.count(AGENT.draft)).toBe(0)
  expect(readLeanJson(directory, "gate.json")).toMatchObject({ delivery: "execution_ready" })
  writeLeanBytes(directory, "drafting-notes.docx", (await renderMarkdownToDocx("# Current notes\n\nA further current Word change.")).bytes)
  const unresolved = model(directory, "open")
  expect((await runLean(directory, { sessions: unresolved.port })).status).toBe("completed")
  expect(readLeanJson(directory, "gate.json")).toMatchObject({ status: "pass", delivery: "review_draft", blocking: [],
    disclosed: [{ stage: "checklist-check", kind: "checklist_open", status: "disclosed", detail: expect.stringContaining("Finding: Focused result open.") }] })
  expect(existsSync(leanFile(directory, "checklist/report.md"))).toBe(true)
}, 120000)

for (const status of ["unsatisfied", "indeterminate"] as const) test(`${status} is disclosed with the review draft rather than blocking delivery`, async () => {
  const directory = matter()
  const scripted = model(directory, status)
  const state = await runLean(directory, { sessions: scripted.port })
  expect(state.status).toBe("completed")
  expect(state.stages["checklist-check"]?.status).toBe("done")
  expect(readOpenItems(directory).some(item => item.stage === "checklist-check")).toBe(false)
  expect(scripted.count(AGENT["package-review"])).toBe(1)
  expect(readLeanJson(directory, "gate.json")).toMatchObject({ status: "pass", delivery: "review_draft_with_findings", blocking: [] })
  expect(report(directory)).toContain("Target: d1; treatment: operative")
  expect(report(directory)).toContain(`Disclosed, not satisfied: ${checklistIDs(directory)[0]}: `)
}, 120000)

test("an omission is repaired once without Continue; an adverse re-check is disclosed with the review draft", async () => {
  const directory = matter()
  const scripted = model(directory, "unsatisfied")
  const state = await runLean(directory, { sessions: scripted.port })
  expect(state.status).toBe("completed")
  expect(state.stages["checklist-check"]?.status).toBe("done")
  const [id] = checklistIDs(directory)
  // Exactly one checklist repair round, adjudicated once, then one re-check.
  expect(focusedRepairs(scripted)).toBe(1)
  expect(readRepairRound(directory)).toMatchObject({ round: 1, substantive_attempts: 1, items: [{ stage: "checklist-check", ref: { document: "d1", repair_target: "body" } }] })
  expect(scripted.count(CHECKLIST_JOB_AGENTS.check)).toBe(2)
  expect(scripted.count(AGENT["package-review"])).toBe(1)
  const gate = readLeanJson(directory, "gate.json") as any
  expect(gate).toMatchObject({ status: "pass", delivery: "review_draft_with_findings", blocking: [] })
  expect(gate.disclosed).toContainEqual(expect.objectContaining({ kind: "checklist_unresolved", ref: expect.objectContaining({ checklist_item: id, checklist_class: "omission" }) }))
  expect(report(directory)).toContain(`Disclosed, not satisfied: ${id}: omission not fixed by the repair round: Focused result unsatisfied.`)
  expect(report(directory)).not.toContain("Delivery blocked")
  expect(readOpenItems(directory).some(item => item.stage === "checklist-check")).toBe(false)
  const round = readLeanJson(directory, "repairs.json")
  enqueueLeanRequest(directory, { kind: "continue" })
  const idle = scriptedModel({})
  expect((await runLean(directory, { sessions: idle.port })).status).toBe("completed")
  expect(idle.prompts).toHaveLength(0)
  expect(readLeanJson(directory, "repairs.json")).toEqual(round)
}, 120000)

for (const [label, scripted] of [
  ["not_inspectable", (directory: string) => model(directory, "indeterminate", { indeterminate_cause: "not_inspectable" })],
  ["departure", (directory: string) => model(directory, "unsatisfied", {}, { departure: true })],
] as const) test(`a ${label} finding starts no repair and is disclosed`, async () => {
  const directory = matter()
  const run = scripted(directory)
  const state = await runLean(directory, { sessions: run.port })
  expect(state.status).toBe("completed")
  const [id] = checklistIDs(directory)
  expect(focusedRepairs(run)).toBe(0)
  expect(readRepairRound(directory)).toBeUndefined()
  expect(run.count(CHECKLIST_JOB_AGENTS.check)).toBe(1)
  expect(readLeanJson(directory, "gate.json")).toMatchObject({ status: "pass", delivery: "review_draft_with_findings", blocking: [] })
  expect(report(directory)).toContain(`Disclosed, not satisfied: ${id}: ${label === "departure" ? "deliberate departure" : "cannot be established from document text"}: Focused result `)
  expect(report(directory)).not.toContain("Delivery blocked")
}, 120000)

// Unresolved findings reach the lawyer in the delivered notes: the gate puts
// them first, then the drafting notes unchanged, and the notes deliverable
// carries that file under its usual name. With none, the drafting notes are
// delivered as written.
const notesRows = (directory: string) => readLeanDocuments(directory)!.documents.filter(row => row.document === null).map(row => [row.kind, row.path, row.deliverable])
const NOTES_NAME = "Supply note - drafting notes.docx"

test("unresolved findings lead the delivered notes, registered under the notes' own name; a gate re-run changes nothing", async () => {
  const directory = matter()
  const deliverables = registering()
  const run = model(directory, "indeterminate", { indeterminate_cause: "not_inspectable" }, { pair: true })
  expect((await runLean(directory, { sessions: run.port, registerDeliverable: deliverables.registerDeliverable })).status).toBe("completed")
  const gate = readLeanJson(directory, "gate.json") as GateResult
  const findings = gate.items.filter(item => item.kind === "checklist_unresolved")
  expect(gate.delivery).toBe("review_draft_with_findings")
  expect(findings.map(item => item.ref.checklist_item)).toEqual(checklistIDs(directory))
  expect(findings).toHaveLength(2)
  const notes = readFileSync(leanFile(directory, "drafting-notes.md"), "utf8")
  const delivered = readFileSync(leanFile(directory, "delivered-notes.md"), "utf8")
  expect(delivered.endsWith(notes)).toBe(true)
  const lead = delivered.slice(0, -notes.length)
  for (const item of findings) expect(lead).toContain(`- ${item.ref.checklist_item}: ${item.detail} Source: `)
  expect(lead).toContain("(cannot be established from document text)")
  const docx = readFileSync(leanFile(directory, "delivered-notes.docx"))
  const text = readDocxCommentParagraphs(docx).map(row => row.paragraph_text)
  const at = (needle: string) => text.findIndex(line => line.includes(needle))
  expect(findings.map(item => at(item.detail)).every(index => index >= 0 && index < at("The draft states the delivery period from the supply terms."))).toBe(true)
  expect(notesRows(directory)).toEqual([["drafting_notes", ".dochaus/lean/delivered-notes.docx", NOTES_NAME], ["drafting_notes_markdown", ".dochaus/lean/delivered-notes.md", null]])
  expect(readFileSync(path.join(directory, NOTES_NAME))).toEqual(docx)
  expect(run.prompts.find(row => row.agent === AGENT["package-review"])!.text).toContain('a "Disclosed, not satisfied" line is an unresolved finding')
  // The same inputs give the same bytes; registration reuses the same name.
  const summary = readFileSync(leanFile(directory, "documents.json"), "utf8")
  updateLeanStage(directory, readLeanState(directory)!, "gate", { status: "pending", input_sha: null })
  const idle = scriptedModel({})
  expect((await runLean(directory, { sessions: idle.port, registerDeliverable: deliverables.registerDeliverable })).status).toBe("completed")
  expect(idle.prompts).toHaveLength(0)
  expect(readLeanState(directory)!.stages.gate?.input_sha).not.toBeNull()
  expect(readFileSync(leanFile(directory, "delivered-notes.md"), "utf8")).toBe(delivered)
  expect(readFileSync(leanFile(directory, "delivered-notes.docx"))).toEqual(docx)
  expect(readFileSync(leanFile(directory, "documents.json"), "utf8")).toBe(summary)
  expect(readdirSync(directory).filter(name => name.includes("drafting notes"))).toEqual([NOTES_NAME])
  // Missing delivered notes make the gate's output invalid: it runs again
  // rather than delivering the drafting notes without the findings.
  rmSync(leanFile(directory, "delivered-notes.docx"))
  expect((await runLean(directory, { sessions: scriptedModel({}).port, registerDeliverable: deliverables.registerDeliverable })).status).toBe("completed")
  expect(notesRows(directory)[0]).toEqual(["drafting_notes", ".dochaus/lean/delivered-notes.docx", NOTES_NAME])
  expect(readFileSync(path.join(directory, NOTES_NAME))).toEqual(readFileSync(leanFile(directory, "delivered-notes.docx")))
}, 120000)

test("with no unresolved finding the drafting notes are delivered as written; cleared findings leave no delivered notes", async () => {
  const directory = matter()
  const deliverables = registering()
  const run = model(directory, "indeterminate", { indeterminate_cause: "not_inspectable" })
  expect((await runLean(directory, { sessions: run.port, registerDeliverable: deliverables.registerDeliverable })).status).toBe("completed")
  expect(existsSync(leanFile(directory, "delivered-notes.docx"))).toBe(true)
  // A current Word change to the notes re-runs the check, which now finds the requirement met.
  writeLeanBytes(directory, "drafting-notes.docx", (await renderMarkdownToDocx("# Current notes\n\nChanged current Word content.")).bytes)
  const clean = model(directory)
  expect((await runLean(directory, { sessions: clean.port, registerDeliverable: deliverables.registerDeliverable })).status).toBe("completed")
  expect(readLeanJson(directory, "gate.json")).toMatchObject({ status: "pass", delivery: "execution_ready" })
  expect(["delivered-notes.md", "delivered-notes.docx"].some(name => existsSync(leanFile(directory, name)))).toBe(false)
  expect(notesRows(directory)).toEqual([["drafting_notes", ".dochaus/lean/drafting-notes.docx", NOTES_NAME], ["drafting_notes_markdown", ".dochaus/lean/drafting-notes.md", null]])
  expect(readFileSync(path.join(directory, NOTES_NAME))).toEqual(readFileSync(leanFile(directory, "drafting-notes.docx")))
  expect(readdirSync(directory).filter(name => name.includes("drafting notes"))).toEqual([NOTES_NAME])
}, 120000)

// An indeterminate finding is a missing inspection whether or not the checker
// recorded that cause: one reinspection in the stage, limited to that item,
// then disclosure. It is never repaired and needs no Continue.
for (const [label, extra] of [["missing_inspection", { indeterminate_cause: "missing_inspection" }], ["causeless indeterminate", {}]] as const) test(`a ${label} finding is reinspected once in the stage, then disclosed without a repair round`, async () => {
  const directory = matter()
  const run = model(directory, "indeterminate", extra, { second: true })
  const state = await runLean(directory, { sessions: run.port })
  expect(state.status).toBe("completed")
  expect(state.stages["checklist-check"]?.status).toBe("done")
  const [id, other] = checklistIDs(directory)
  expect([checked(run, id!), checked(run, other!)]).toEqual([2, 1])
  const rows = (readLeanJson(directory, "checklist/checks.json") as any).checks
  expect(rows.find((row: any) => row.item_id === id)).toMatchObject({ finding: { status: "indeterminate" }, reinspection: { original_finding: { status: "indeterminate" }, status: "succeeded" } })
  expect(rows.find((row: any) => row.item_id === other).reinspection).toBeUndefined()
  expect(focusedRepairs(run)).toBe(0)
  expect(readRepairRound(directory)).toBeUndefined()
  const gate = readLeanJson(directory, "gate.json") as any
  expect(gate).toMatchObject({ status: "pass", delivery: "review_draft_with_findings", blocking: [] })
  expect(gate.disclosed).toContainEqual(expect.objectContaining({ kind: "checklist_unresolved", ref: expect.objectContaining({ checklist_item: id, checklist_class: "missing_inspection" }),
    detail: expect.stringContaining("(not established by the check after reinspection)") }))
  expect(report(directory)).toContain(`Disclosed, not satisfied: ${id}: not established by the check after reinspection: Focused result indeterminate.`)
  expect(report(directory)).not.toContain("Delivery blocked")
  enqueueLeanRequest(directory, { kind: "continue" })
  const idle = scriptedModel({})
  expect((await runLean(directory, { sessions: idle.port })).status).toBe("completed")
  expect(idle.prompts).toHaveLength(0)
  // The reinspection is spent: even Continue at a blocked checklist inspects nothing again.
  settleBlockedChecklist(directory)
  enqueueLeanRequest(directory, { kind: "continue" })
  const blocked = model(directory, "satisfied", {}, { second: true })
  expect((await runLean(directory, { sessions: blocked.port })).status).toBe("completed")
  expect(blocked.count(CHECKLIST_JOB_AGENTS.check)).toBe(0)
  expect(focusedRepairs(blocked)).toBe(0)
}, 120000)

test("an omission with the repair allowance used starts no round and is disclosed as such", async () => {
  const directory = matter()
  const used = JSON.stringify({ round: 5, substantive_attempts: 5, requested_at: "saved", items: [] })
  writeLeanText(directory, "repairs.json", used)
  const run = model(directory, "unsatisfied")
  expect((await runLean(directory, { sessions: run.port })).status).toBe("completed")
  const [id] = checklistIDs(directory)
  expect(focusedRepairs(run)).toBe(0)
  expect(run.count(CHECKLIST_JOB_AGENTS.check)).toBe(1)
  expect(readFileSync(leanFile(directory, "repairs.json"), "utf8")).toBe(used)
  const gate = readLeanJson(directory, "gate.json") as any
  expect(gate).toMatchObject({ status: "pass", delivery: "review_draft_with_findings", blocking: [] })
  expect(gate.disclosed).toContainEqual(expect.objectContaining({ kind: "checklist_unresolved", ref: expect.objectContaining({ checklist_item: id, checklist_class: "omission" }),
    detail: expect.stringContaining("(omission; the repair allowance was used up)") }))
  expect(report(directory)).toContain(`Disclosed, not satisfied: ${id}: omission; the repair allowance was used up: Focused result unsatisfied.`)
  expect(report(directory)).not.toContain("not fixed by the repair round")
}, 120000)

test("an integrity failure still blocks the checklist with its finding and starts no repair", async () => {
  const directory = matter()
  const scripted = model(directory, "unsatisfied", {}, { broken: true })
  const state = await runLean(directory, { sessions: scripted.port })
  const [adverse, failed] = checklistIDs(directory)
  expect(state.status).toBe("blocked")
  expect(state.stages["checklist-check"]?.status).toBe("blocked")
  expect(state.stop_reason).toBe(`Checklist handoff blocked: ${failed}: failed: Scripted read-only inspection failure. See checklist/report.md.`)
  expect(readOpenItems(directory)).toContainEqual(expect.objectContaining({ stage: "checklist-check", status: "open", detail: `${failed}: failed: Scripted read-only inspection failure` }))
  expect(report(directory)).toContain(`Delivery blocked: ${failed}: failed: Scripted read-only inspection failure`)
  expect(report(directory)).toContain(`Disclosed, not satisfied: ${adverse}: omission; no repair round ran: Focused result unsatisfied.`)
  // The omission waits: no round starts while an integrity failure blocks.
  expect(focusedRepairs(scripted)).toBe(0)
  expect(readRepairRound(directory)).toBeUndefined()
  expect(scripted.count(AGENT["package-review"])).toBe(0)
}, 120000)

test("a declined checklist repair is disclosed with the adjudicator's reason, not blocked", async () => {
  const { directory, state } = await declinedRepairRun()
  expect(state.status).toBe("completed")
  const [id] = checklistIDs(directory)
  expect(readLeanJson(directory, "gate.json")).toMatchObject({ status: "pass", delivery: "review_draft_with_findings", blocking: [] })
  expect(report(directory)).toContain(`Disclosed, not satisfied: ${id}: omission; source adjudication declined the repair: Focused result unsatisfied. Adjudicator's reason: DECLINED_REASON: the supplied terms already state the delivery period.`)
  expect(report(directory)).not.toContain("Delivery blocked")
}, 120000)

test("informational open finding remains open in report while permitting review draft workflow", async () => {
  const directory = matter()
  const scripted = model(directory, "open")
  expect((await runLean(directory, { sessions: scripted.port })).status).toBe("completed")
  expect(readLeanJson(directory, "checklist/checks.json")).toMatchObject({ checks: [{ finding: { status: "open" } }] })
  // The finding stays open, non-blocking and unchanged; it is disclosed rather
  // than coerced to a pass or to a generic blocking omission.
  const gate = readLeanJson(directory, "gate.json") as any
  expect(gate).toMatchObject({ status: "pass", delivery: "review_draft", blocking: [] })
  expect(gate.disclosed).toHaveLength(1)
  expect(gate.disclosed[0]).toMatchObject({ stage: "checklist-check", kind: "checklist_open", status: "disclosed", ref: { document: "d1" },
    detail: "Checklist requirement left open: The requested supply note is drafted. Finding: Focused result open. Limitation: Comments not inspected." })
  expect(gate.items).toContainEqual(gate.disclosed[0])
  expect(readOpenItems(directory).some(item => item.kind === "checklist_open")).toBe(false)
  expect(readFileSync(leanFile(directory, "checklist/report.md"), "utf8")).toContain("Status: open")
}, 120000)

// After inspection, a requirement whose only open question is whether the
// request's scope covers it goes to the lawyer as a disclosed decision instead
// of an indeterminate finding that permanently blocks once reinspected.
const scopeDecision = "Whether the request's scope requires this deliverable to cover the requirement"
// An indeterminate finding whose one reinspection waits for an explicit Continue.
const notInspectable = { indeterminate_cause: "not_inspectable" } as const

test("an inspected applicability decision left open completes as a review draft disclosing that decision", async () => {
  const directory = matter()
  const scripted = model(directory, "open", { open_decisions: [scopeDecision] })
  const state = await runLean(directory, { sessions: scripted.port })
  expect(state.status).toBe("completed")
  expect(state.stages["checklist-check"]?.status).toBe("done")
  expect(scripted.prompts.find(row => row.agent === CHECKLIST_JOB_AGENTS.check)?.text).toContain("report open, name that decision in open_decisions")
  expect(readLeanJson(directory, "checklist/checks.json")).toMatchObject({ checks: [{ finding: { status: "open", open_decisions: [scopeDecision] } }] })
  const gate = readLeanJson(directory, "gate.json") as any
  expect(gate).toMatchObject({ status: "pass", delivery: "review_draft", blocking: [] })
  expect(gate.disclosed).toHaveLength(1)
  expect(gate.disclosed[0]).toMatchObject({ stage: "checklist-check", kind: "checklist_open", status: "disclosed", ref: { document: "d1" } })
  expect(gate.disclosed[0].detail).toContain(`Open decisions: ${scopeDecision}.`)
  expect(readFileSync(leanFile(directory, "checklist/report.md"), "utf8")).not.toContain("Delivery blocked")
  const idle = scriptedModel({})
  expect((await runLean(directory, { sessions: idle.port })).status).toBe("completed")
  expect(idle.prompts).toHaveLength(0)
}, 120000)

test("the one reinspection of an indeterminate finding may settle on an open applicability decision without redrafting", async () => {
  const directory = matter()
  expect((await runLean(directory, { sessions: model(directory, "indeterminate", notInspectable).port })).status).toBe("completed")
  settleBlockedChecklist(directory)
  enqueueLeanRequest(directory, { kind: "continue" })
  const review = model(directory, "open", { open_decisions: [scopeDecision] })
  expect((await runLean(directory, { sessions: review.port })).status).toBe("completed")
  expect(review.count(CHECKLIST_JOB_AGENTS.check)).toBe(1)
  expect(review.count(CHECKLIST_JOB_AGENTS.extract)).toBe(0)
  expect(review.count(AGENT.draft)).toBe(0)
  expect(readLeanJson(directory, "checklist/checks.json")).toMatchObject({ checks: [{ finding: { status: "open", open_decisions: [scopeDecision] },
    reinspection: { original_finding: { status: "indeterminate" }, status: "succeeded" } }] })
  expect(readLeanJson(directory, "gate.json")).toMatchObject({ status: "pass", delivery: "review_draft",
    disclosed: [{ kind: "checklist_open", detail: expect.stringContaining(`Open decisions: ${scopeDecision}.`) }] })
}, 120000)


test("redline-only native comment edits invalidate completed checklist checks", async () => {
  const directory = matter()
  expect((await runLean(directory, { sessions: model(directory).port })).status).toBe("completed")
  const before = stageInputSha(directory, "checklist-check")
  const original = readFileSync(leanFile(directory, "d1/redline.docx"))
  const paragraph = readDocxCommentParagraphs(original)[0]!
  const annotated = exportDocxComments(original, [{ id: "review-note", text: "Review this wording.", author: "Reviewer", anchor: { paragraph_index: paragraph.paragraph_index, paragraph_text: paragraph.paragraph_text } }])
  writeLeanBytes(directory, "d1/redline.docx", annotated.bytes)
  expect(stageInputSha(directory, "checklist-check")).not.toBe(before)
  const refreshed = model(directory)
  expect((await runLean(directory, { sessions: refreshed.port })).status).toBe("completed")
  expect(refreshed.count(CHECKLIST_JOB_AGENTS.check)).toBe(1)
  expect(refreshed.count(CHECKLIST_JOB_AGENTS.extract)).toBe(0)
  expect(refreshed.count(AGENT.draft)).toBe(0)
}, 120000)

test("completed checklist validation uses authoritative extraction IDs and full finding shape", async () => {
  const directory = matter()
  expect((await runLean(directory, { sessions: model(directory).port })).status).toBe("completed")
  const saved = readLeanJson(directory, "checklist/checks.json") as any
  const save = (value: unknown) => writeLeanText(directory, "checklist/checks.json", JSON.stringify(value))
  expect(stageOutputValid(directory, "checklist-check")).toBe(true)
  save({ ...saved, expected_items: [], checks: [] })
  expect(stageOutputValid(directory, "checklist-check")).toBe(false)
  const first = saved.checks[0]
  for (const finding of [
    { ...first.finding, status: "unknown" },
    { ...first.finding, item_id: "another-item" },
    { ...first.finding, inspection_scope: undefined },
  ]) {
    save({ ...saved, checks: [{ ...first, finding }] })
    expect(stageOutputValid(directory, "checklist-check")).toBe(false)
  }
  save({ ...saved, expected_items: [first.item_id, first.item_id], checks: [first, first] })
  expect(stageOutputValid(directory, "checklist-check")).toBe(false)
  save(saved)
  expect(stageOutputValid(directory, "checklist-check")).toBe(true)
}, 120000)


test("at an integrity block, only explicit Continue reinspects an adverse finding, once, without redrafting", async () => {
  const directory = matter()
  // The second requirement's check job always fails: the stage stays blocked
  // and that failed check re-runs on every pass.
  expect((await runLean(directory, { sessions: model(directory, "indeterminate", {}, { broken: true }).port })).status).toBe("blocked")
  const [adverse, failed] = checklistIDs(directory)
  const idle = model(directory, "satisfied", {}, { broken: true })
  expect((await runLean(directory, { sessions: idle.port })).status).toBe("blocked")
  expect([checked(idle, adverse!), checked(idle, failed!)]).toEqual([0, 1])
  enqueueLeanRequest(directory, { kind: "continue" })
  const review = model(directory, "indeterminate", {}, { broken: true })
  const stillBlocked = await runLean(directory, { sessions: review.port })
  expect(stillBlocked.status).toBe("blocked")
  expect(stillBlocked.stop_reason).toBe(`Checklist handoff blocked: ${failed}: failed: Scripted read-only inspection failure. See checklist/report.md.`)
  expect([checked(review, adverse!), checked(review, failed!)]).toEqual([1, 1])
  expect(review.count(AGENT.draft)).toBe(0)
  expect(review.count(CHECKLIST_JOB_AGENTS.extract)).toBe(0)
  expect(report(directory)).toContain("Original finding: indeterminate")
  expect(report(directory)).toContain(`Disclosed, not satisfied: ${adverse}: `)
  enqueueLeanRequest(directory, { kind: "continue" })
  const exhausted = model(directory, "satisfied", {}, { broken: true })
  expect((await runLean(directory, { sessions: exhausted.port })).status).toBe("blocked")
  expect([checked(exhausted, adverse!), checked(exhausted, failed!)]).toEqual([0, 1])
}, 120000)

test("a pause during the one reinspection spends nothing: Continue runs it again and uses its finding; the late reply is never recorded", async () => {
  const directory = matter()
  expect((await runLean(directory, { sessions: model(directory, "indeterminate", notInspectable).port })).status).toBe("completed")
  settleBlockedChecklist(directory)
  const checks = () => readLeanJson(directory, "checklist/checks.json") as { checks: { finding: unknown; reinspection?: { status: string; sessionID?: string } }[] }
  const original = checks().checks[0]!.finding
  enqueueLeanRequest(directory, { kind: "continue" })
  const late = model(directory, "satisfied")
  const asked = Promise.withResolvers<string>(), hold = Promise.withResolvers<void>()
  const interrupts: string[] = []
  const sessions: StageSessionPort = {
    createSession: input => late.port.createSession(input),
    async prompt(input) {
      if (input.agent === CHECKLIST_JOB_AGENTS.check) {
        asked.resolve(input.sessionID)
        await hold.promise
      }
      return late.port.prompt(input)
    },
    interrupt: async input => { interrupts.push(input.sessionID) },
  }
  const pause = new AbortController()
  const working = runLean(directory, { sessions, pause: pause.signal })
  const session = await asked.promise
  pause.abort()
  const paused = await working
  expect(paused).toMatchObject({ status: "blocked", stop_reason: "Paused at your request.", stages: { "checklist-check": { status: "pending" } } })
  expect(paused.paused_at).toBeDefined()
  expect(interrupts).toEqual([session])
  const admitted = checks()
  expect(admitted.checks[0]).toMatchObject({ finding: original, reinspection: { status: "failed" } })
  expect(admitted.checks[0]!.reinspection!.sessionID).toBeUndefined()
  // The held reply arrives after the pause and is never recorded.
  hold.resolve()
  await Bun.sleep(20)
  expect(late.count(CHECKLIST_JOB_AGENTS.check)).toBe(1)
  expect(checks()).toEqual(admitted)
  enqueueLeanRequest(directory, { kind: "continue" })
  const resumed = model(directory, "satisfied")
  expect((await runLean(directory, { sessions: resumed.port })).status).toBe("completed")
  expect(resumed.count(CHECKLIST_JOB_AGENTS.check)).toBe(1)
  expect(resumed.count(CHECKLIST_JOB_AGENTS.extract)).toBe(0)
  expect(resumed.count(AGENT.draft)).toBe(0)
  expect(checks()).toMatchObject({ checks: [{ finding: { status: "satisfied" }, reinspection: { original_finding: { status: "indeterminate" }, status: "succeeded" } }] })
}, 120000)

test("a focused model review may resolve an adverse assessment without changing its item or Word bytes", async () => {
  const directory = matter()
  await runLean(directory, { sessions: model(directory, "indeterminate", notInspectable).port })
  settleBlockedChecklist(directory)
  const original = readLeanJson(directory, "checklist/extraction.json")
  const bytes = readFileSync(leanFile(directory, "d1/body-applied.docx"))
  enqueueLeanRequest(directory, { kind: "continue" })
  const review = model(directory, "satisfied")
  expect((await runLean(directory, { sessions: review.port })).status).toBe("completed")
  expect(review.count(CHECKLIST_JOB_AGENTS.check)).toBe(1)
  expect(review.count(AGENT.draft)).toBe(0)
  expect(readLeanJson(directory, "checklist/extraction.json")).toEqual(original)
  expect(Buffer.compare(readFileSync(leanFile(directory, "d1/body-applied.docx")), bytes)).toBe(0)
  expect(readLeanJson(directory, "checklist/checks.json")).toMatchObject({ checks: [{ finding: { status: "satisfied" }, reinspection: { original_finding: { status: "indeterminate" }, status: "succeeded" } }] })
}, 120000)


test("accepted focused refinement reruns late verification only and archived adverse parent does not block", async () => {
  const directory = matter()
  expect((await runLean(directory, { sessions: model(directory, "indeterminate").port })).status).toBe("completed")
  settleBlockedChecklist(directory)
  const extraction = readLeanJson(directory, "checklist/extraction.json") as any
  const parent = extraction.sources.flatMap((row: any) => row.items)[0]
  const prior = (readLeanJson(directory, "checklist/checks.json") as any).checks[0]
  const word = readFileSync(leanFile(directory, "d1/body-applied.docx"))
  const early = ["d1.draft", "d1.review", "d1.change-plan", "d1.apply", "finalize"] as const
  const keys = early.map(step => stageInputSha(directory, step))
  // This fixture has only a native source DB and a request extraction item.
  writeFileSync(`${directory}/source-map.json`, JSON.stringify({ documents: [] }))
  enqueueLeanRequest(directory, { kind: "refine_checklist", parents: [{ parent_id: parent.id, scope: "cited_sources" }] })
  const focused = scriptedModel({ ...happyPath,
    [CHECKLIST_JOB_AGENTS.extract]: text => {
      const packet = JSON.parse(text.split("REFINEMENT_INPUT:\n\n")[1]!)
      return JSON.stringify({ parent_id: packet.parent_id, unit_id: packet.source_unit.id, disposition: "refined", reason: "The instruction supports this focused action.", items: [{ ...parent.item, required_action: "propose" }], source_refs: parent.item.source_refs, limitations: [] })
    },
    [CHECKLIST_JOB_AGENTS.check]: text => {
      const packet = JSON.parse(text.split("CHECK_INPUT:\n\n")[1]!)
      return JSON.stringify({ item_id: packet.item_id, status: "open", action_status: "fulfilled", open_decisions: ["Approval remains open."], reason: "The requested action is evidenced; approval remains unresolved.", evidence: packet.current_evidence, source_refs: packet.item.source_refs, inspection_scope: "Current target passages.", limitation: "No approval inferred." })
    },
  })
  expect((await runLean(directory, { sessions: focused.port })).status).toBe("completed")
  expect(focused.count(CHECKLIST_JOB_AGENTS.extract)).toBe(1)
  expect(focused.count(CHECKLIST_JOB_AGENTS.check)).toBe(1)
  expect(focused.count(AGENT.draft)).toBe(0)
  expect(focused.count(AGENT.finalize)).toBe(0)
  expect(early.map(step => stageInputSha(directory, step))).toEqual(keys)
  expect(readLeanJson(directory, "checklist/extraction.json")).toEqual(extraction)
  expect(readFileSync(leanFile(directory, "d1/body-applied.docx"))).toEqual(word)
  expect((readLeanJson(directory, "checklist/checks.json") as any).checks).toContainEqual(prior)
  expect(stageOutputValid(directory, "checklist-check")).toBe(true)
  expect(finalGate(directory)).toMatchObject({ status: "pass", delivery: "review_draft" })
  const idle = scriptedModel({})
  expect((await runLean(directory, { sessions: idle.port })).status).toBe("completed")
  expect(idle.prompts).toHaveLength(0)
  const saved = readLeanJson(directory, "checklist/refinements.json") as any
  saved.parents[0].expected_units.push("unaccounted-source-unit")
  writeLeanText(directory, "checklist/refinements.json", JSON.stringify(saved))
  expect(stageOutputValid(directory, "checklist-check")).toBe(false)
  expect(finalGate(directory).status).toBe("blocked")
}, 120000)

// A document-wide parent split for "Check the full draft": range and process
// components defer it, and only an affirmative document-properties finding
// decides it; an undecided parent is disclosed with the review draft. Unmarked
// all-not_applicable children keep base semantics.
for (const [status, marked, disclosed] of [["satisfied", true, false], ["indeterminate", true, true], ["not_applicable", true, true], ["not_applicable", false, false]] as const) test(`full-draft check of a document-wide requirement: ${marked ? "deferred" : "undeferred"} ranges and ${status} document-properties finding ${disclosed ? "disclose the parent" : "complete the run"}`, async () => {
  const directory = matter()
  expect((await runLean(directory, { sessions: model(directory, "indeterminate").port })).status).toBe("completed")
  settleBlockedChecklist(directory)
  const parent = (readLeanJson(directory, "checklist/extraction.json") as any).sources.flatMap((row: any) => row.items)[0]
  enqueueLeanRequest(directory, { kind: "refine_checklist", parents: [{ parent_id: parent.id, scope: "target_body" }] })
  const inputs: ChecklistCheckerInput[] = []
  const full = scriptedModel({ ...happyPath,
    [CHECKLIST_JOB_AGENTS.check]: text => {
      const input = JSON.parse(text.split("CHECK_INPUT:\n\n")[1]!) as ChecklistCheckerInput
      inputs.push(input)
      const whole = input.assignment?.kind === "document_properties"
      return JSON.stringify({ item_id: input.item_id, status: whole ? status : "not_applicable",
        reason: whole ? `Document-wide result ${status}.` : "Only the document-properties assignment decides this document-wide property.",
        evidence: whole ? input.current_evidence : [], source_refs: input.item.source_refs, inspection_scope: JSON.stringify(input.assignment), limitation: "Rendered layout not inspected.",
        ...(!whole && marked ? { deferred_to: "document_properties" } : {}) })
    },
  })
  const state = await runLean(directory, { sessions: full.port })
  expect(state.status).toBe("completed")
  expect(state.stages["checklist-check"]?.status).toBe("done")
  // A causeless indeterminate finding is a missing inspection: its one reinspection runs in the stage.
  expect(inputs.filter(input => input.assignment?.kind === "document_properties")).toHaveLength(status === "indeterminate" ? 2 : 1)
  expect(inputs.filter(input => input.assignment?.kind === "process_records")).toHaveLength(1)
  expect(inputs.every(input => input.assignment)).toBe(true)
  const whole = inputs.find(input => input.assignment?.kind === "document_properties")!
  expect(whole.current_evidence.length).toBeGreaterThan(0)
  expect(whole.current_evidence.every(row => row.kind === "artifact" && "accepted_body_text_measures" in JSON.parse(row.text))).toBe(true)
  expect(full.count(CHECKLIST_JOB_AGENTS.extract)).toBe(0)
  expect(full.count(AGENT.draft)).toBe(0)
  expect(stageOutputValid(directory, "checklist-check")).toBe(true)
  expect(report(directory)).not.toContain("Delivery blocked")
  if (disclosed) {
    expect(report(directory)).toContain(`Disclosed, not satisfied: ${parent.id}: refined checks indeterminate: `)
    expect(readLeanJson(directory, "gate.json")).toMatchObject({ status: "pass", delivery: "review_draft_with_findings", blocking: [] })
  } else expect(report(directory)).not.toContain(`Disclosed, not satisfied: ${parent.id}: `)
  if (disclosed && status === "not_applicable") expect(report(directory)).toContain("no document-properties finding decided it")
}, 120000)

// Each full-draft component applies the same applicability rule: paragraph
// components left open on the scope decision with the other components not
// applicable reconcile the parent open; any indeterminate component leaves it
// indeterminate, disclosed with the review draft.
for (const [other, parentStatus] of [["not_applicable", "open"], ["indeterminate", "indeterminate"]] as const) test(`full-draft components left open on an applicability decision with ${other} other components reconcile the parent ${parentStatus}`, async () => {
  const directory = matter()
  expect((await runLean(directory, { sessions: model(directory, "indeterminate").port })).status).toBe("completed")
  settleBlockedChecklist(directory)
  const parent = (readLeanJson(directory, "checklist/extraction.json") as any).sources.flatMap((row: any) => row.items)[0]
  enqueueLeanRequest(directory, { kind: "refine_checklist", parents: [{ parent_id: parent.id, scope: "target_body" }] })
  const instructions: string[] = []
  const full = scriptedModel({ ...happyPath,
    [CHECKLIST_JOB_AGENTS.check]: text => {
      const input = JSON.parse(text.split("CHECK_INPUT:\n\n")[1]!) as ChecklistCheckerInput
      instructions.push(text.split("CHECK_INPUT:")[0]!)
      const body = input.assignment?.kind === "target_body"
      return JSON.stringify({ item_id: input.item_id, status: body ? "open" : other,
        reason: body ? "The assigned paragraphs state the arrangement generally; its stated conditions would be missing if the request covers them." : `Component result ${other}.`,
        evidence: body ? input.current_evidence : [], source_refs: input.item.source_refs, inspection_scope: JSON.stringify(input.assignment), limitation: "No scope decision is inferred.",
        ...(body ? { open_decisions: [scopeDecision] } : {}) })
    },
  })
  const state = await runLean(directory, { sessions: full.port })
  expect(instructions.length).toBeGreaterThanOrEqual(3)
  expect(instructions.every(text => text.includes("Every component applies the applicability rule below exactly as a whole-item check does."))).toBe(true)
  expect(readFileSync(leanFile(directory, "checklist/report.md"), "utf8")).toContain(`coverage: target_body; status: ${parentStatus}`)
  expect(full.count(AGENT.draft)).toBe(0)
  expect(state.status).toBe("completed")
  expect(state.stages["checklist-check"]?.status).toBe("done")
  expect(stageOutputValid(directory, "checklist-check")).toBe(true)
  const gate = readLeanJson(directory, "gate.json") as any
  if (parentStatus === "indeterminate") {
    expect(report(directory)).toContain(`Disclosed, not satisfied: ${parent.id}: refined checks indeterminate: `)
    expect(gate).toMatchObject({ status: "pass", delivery: "review_draft_with_findings", blocking: [] })
    return
  }
  expect(gate).toMatchObject({ status: "pass", delivery: "review_draft", blocking: [] })
  const open = gate.disclosed.filter((item: any) => item.kind === "checklist_open")
  expect(open.length).toBeGreaterThan(0)
  expect(open.every((item: any) => item.detail.includes(`Open decisions: ${scopeDecision}.`))).toBe(true)
}, 120000)

// Text measures can disprove but never establish a rendered-length directive:
// an approximate one stays open (disclosed in the report, review draft); a
// clearly exceeded one is unsatisfied and takes the one automatic repair round.
async function documentPropertiesRun(status: "open" | "unsatisfied") {
  const directory = matter()
  expect((await runLean(directory, { sessions: model(directory, "indeterminate").port })).status).toBe("completed")
  settleBlockedChecklist(directory)
  const parent = (readLeanJson(directory, "checklist/extraction.json") as any).sources.flatMap((row: any) => row.items)[0]
  enqueueLeanRequest(directory, { kind: "refine_checklist", parents: [{ parent_id: parent.id, scope: "target_body" }] })
  const seen = { measures: "", guided: false }
  const full = scriptedModel({ ...happyPath,
    [AGENT["change-plan"]]: deferChecklistRepair,
    [CHECKLIST_JOB_AGENTS.check]: text => {
      const input = JSON.parse(text.split("CHECK_INPUT:\n\n")[1]!) as ChecklistCheckerInput
      if (input.assignment?.kind !== "document_properties")
        return JSON.stringify({ item_id: input.item_id, status: "not_applicable", reason: "Only the document-properties assignment decides this document-wide property.", evidence: [], source_refs: input.item.source_refs, inspection_scope: JSON.stringify(input.assignment), limitation: "", deferred_to: "document_properties" })
      seen.guided = text.split("CHECK_INPUT:")[0]!.includes(checklistTextMeasuresMeaning)
      seen.measures = JSON.stringify(JSON.parse(input.current_evidence[0]!.text).accepted_body_text_measures)
      return JSON.stringify({ item_id: input.item_id, status, reason: `Accepted body text measures ${seen.measures} ${status === "open" ? "do not clearly exceed about one page" : "clearly exceed one page"}.`,
        evidence: input.current_evidence, source_refs: input.item.source_refs, inspection_scope: "Exact artifact properties.", limitation: "Rendered pagination was not verified." })
    },
  })
  return { directory, parent, state: await runLean(directory, { sessions: full.port }), seen, full }
}

test("an approximate rendered-length directive left open by text measures completes as a review draft with its measures in the report", async () => {
  const { directory, state, seen } = await documentPropertiesRun("open")
  expect(seen.guided).toBe(true)
  expect(seen.measures).toMatch(/^\{"paragraphs":\d+,"words":\d+,"characters":\d+\}$/)
  expect(state.status).toBe("completed")
  expect(state.stages["checklist-check"]?.status).toBe("done")
  expect(stageOutputValid(directory, "checklist-check")).toBe(true)
  expect(readLeanJson(directory, "gate.json")).toMatchObject({ status: "pass", delivery: "review_draft", blocking: [] })
  const report = readFileSync(leanFile(directory, "checklist/report.md"), "utf8")
  expect(report).toContain(`Status: open\nAccepted body text measures ${seen.measures} do not clearly exceed about one page.`)
  expect(report).toContain("Limitation: Rendered pagination was not verified.")
  expect(report).toContain("coverage: target_body; status: open")
  expect(report).not.toContain("Delivery blocked")
  // The lawyer sees the open requirement on the delivered review draft: the
  // gate discloses the measures and the pagination limitation.
  const gate = readLeanJson(directory, "gate.json") as any
  const open = gate.items.filter((item: any) => item.kind === "checklist_open")
  expect(open).toHaveLength(1)
  expect(gate.disclosed).toContainEqual(open[0])
  expect(open[0]).toMatchObject({ stage: "checklist-check", status: "disclosed", ref: { document: "d1" } })
  expect(open[0].detail).toContain(`Finding: Accepted body text measures ${seen.measures} do not clearly exceed about one page.`)
  expect(open[0].detail).toContain("Limitation: Rendered pagination was not verified.")
  // Derived at the gate from checklist results it already reads: nothing is
  // written to open-items.json, so completed stage keys hold and an unchanged
  // resume does no work.
  expect(readOpenItems(directory).some(item => item.kind === "checklist_open")).toBe(false)
  const keys = (["checklist-check", "package-review", "gate"] as const).map(step => [state.stages[step]?.input_sha, stageInputSha(directory, step)])
  expect(keys.every(([saved, current]) => saved === current)).toBe(true)
  const idle = scriptedModel({})
  expect((await runLean(directory, { sessions: idle.port })).status).toBe("completed")
  expect(idle.prompts).toHaveLength(0)
}, 120000)

test("a rendered-length directive clearly exceeded by text measures takes the one automatic focused repair, then is disclosed", async () => {
  const { directory, parent, state, seen, full } = await documentPropertiesRun("unsatisfied")
  expect(state.status).toBe("completed")
  expect(state.stages["checklist-check"]?.status).toBe("done")
  expect(focusedRepairs(full)).toBe(1)
  const round = readRepairRound(directory)!
  expect(round).toMatchObject({ round: 1, substantive_attempts: 1 })
  expect(round.items).toHaveLength(1)
  expect(round.items[0]).toMatchObject({ stage: "checklist-check", ref: { document: "d1", repair_target: "body" } })
  expect(round.items[0]!.detail).toContain(seen.measures.replaceAll('"', '\\"'))
  expect(readLeanJson(directory, "gate.json")).toMatchObject({ status: "pass", delivery: "review_draft_with_findings", blocking: [] })
  expect(report(directory)).toContain(`Disclosed, not satisfied: ${parent.id}`)
  expect(report(directory)).not.toContain("Delivery blocked")
}, 120000)

test("a pending checklist with real extraction remains required at the gate", () => {
  const directory = matter()
  // Default pending stages alone do not retrofit checks into manual legacy work.
  expect(finalGate(directory).status).toBe("pass")
  writeLeanText(directory, "checklist/extraction.json", JSON.stringify({ expected_sources: [], sources: [] }))
  expect(finalGate(directory).status).toBe("blocked")
})

test("actual recorded research context changes invalidate a completed checklist through normal runner inputs", async () => {
  const directory = matter()
  expect((await runLean(directory, { sessions: model(directory).port })).status).toBe("completed")
  const before = stageInputSha(directory, "checklist-check")
  writeLeanText(directory, "research-context.txt", "Retrieved legal source context\nA recorded context limitation has changed.\n")
  expect(stageInputSha(directory, "checklist-check")).not.toBe(before)
  const refreshed = model(directory)
  expect((await runLean(directory, { sessions: refreshed.port })).status).toBe("completed")
  expect(refreshed.count(CHECKLIST_JOB_AGENTS.check)).toBe(1)
  expect(refreshed.count(CHECKLIST_JOB_AGENTS.extract)).toBe(0)
  expect(refreshed.count(AGENT.draft)).toBe(0)
})


test("a source-backed omission is repaired automatically through cumulative changes while preserving initial drafting", async () => {
  const directory = matter()
  const initialDelivery = "The supplier shall deliver the goods within thirty days of the order date."
  const omitted = "The supplier shall deliver the goods."
  const payment = "Payment shall be due on delivery."
  const initialDraft = DRAFT.replace(omitted, initialDelivery)
  const keys = ["d1.draft", "d1.review", "research"] as const
  let repairing = false, sawActual = false
  let initial: { bytes: Buffer<ArrayBuffer>; extraction: unknown; before: string[]; counts: { changes: number; draft: number; research: number } } | undefined
  const scripted = scriptedModel({ ...happyPath,
    [AGENT.draft]: drafter(initialDraft),
    [AGENT.review]: () => "No further findings.",
    [AGENT["change-plan"]]: text => {
      repairing ||= text.includes("ONE supplied checklist assessment")
      if (repairing) {
        const assessment = JSON.parse(text.split("<assessment>\n")[1]!.split("\n</assessment>")[0]!)
        const actual = assessment.finding.evidence.map((row: { text: string }) => row.text).join("\n")
        sawActual = actual.includes(omitted) && !actual.includes(initialDelivery)
        expect(text).toContain('"item_id"')
        expect(text).toContain('"status":"unsatisfied"')
        expect(text).toContain("<related_generated_history_not_authority>")
        expect(text).toContain("ONE supplied checklist assessment")
        return JSON.stringify({ shared_decisions: [], groups: [{ id: "timing", pids: [pidOf(text, initialDelivery)], goal: "Drop the unsupported timing removal; retain original source-supported timing and all unrelated supported amendments.", sources: SUPPORT, limitations: [] }], dismissed: [], findings: [], questions: [], superseded_groups: [], superseded_findings: [], superseded_decisions: [] })
      }
      return JSON.stringify({ shared_decisions: [], groups: [], dismissed: [], questions: [] })
    },
    [AGENT.changes]: text => repairing ? JSON.stringify({amendments:[], questions:[], retained_initial:[{pid:JSON.parse(text.split("<initial_targets>\n")[1]!.split("\n</initial_targets>")[0]!)[0].pid,reason:"The initial source-supported timing must be restored; the prior shortening was unsupported.",sources:SUPPORT}]}) : JSON.stringify({ amendments: [
      ...(!repairing ? [{ pid: pidOf(text, initialDelivery), updated: omitted, reason: "Prior mistaken shortening removes supplied timing.", sources: SUPPORT }] : []),
      { pid: pidOf(text, "Payment is due on delivery."), updated: payment, reason: "Preserve the supported payment wording correction.", sources: [{ source_id: SOURCE, quote: "Payment is due on delivery." }] },
    ], questions: [] }),
    [CHECKLIST_JOB_AGENTS.extract]: text => {
      const packet = JSON.parse(text.split("SOURCE_INPUT:\n\n")[1]!)
      return JSON.stringify({ items: packet.source.document === "request" ? [{ check: "The supplied delivery timing is preserved.", basis: "directive", source_refs: [{ document: SOURCE, locator: "source paragraph", quote: SUPPORT[0]!.quote }], qualifications: [], target: { ...packet.targets[0], treatment: "operative" } }] : [] })
    },
    [CHECKLIST_JOB_AGENTS.check]: text => {
      const packet = JSON.parse(text.split("CHECK_INPUT:\n\n")[1]!)
      // The first check sees the mistaken cumulative change; record the initial
      // drafting before the automatic repair round starts.
      if (!repairing && !initial) {
        initial = { bytes: readFileSync(leanFile(directory, "d1/body.docx")), extraction: readLeanJson(directory, "checklist/extraction.json"),
          before: keys.map(step => stageInputSha(directory, step)), counts: { changes: scripted.count(AGENT.changes), draft: scripted.count(AGENT.draft), research: scripted.count(AGENT.research) } }
        // Deliberately stale accepted-text sidecar must not substitute for current Word.
        const oldRows = readLeanJson(directory, "d1/paragraphs-applied.json") as { text: string }[]
        writeLeanText(directory, "d1/paragraphs-applied.json", JSON.stringify(oldRows.map(row => ({ ...row, text: row.text === omitted ? initialDelivery : row.text }))))
      }
      return JSON.stringify({ item_id: packet.item_id, status: repairing ? "satisfied" : "unsatisfied", reason: repairing ? "Current operative timing is present." : "The cumulative amendment removed the supplied timing.", evidence: packet.current_evidence, source_refs: packet.item.source_refs, inspection_scope: "Current native Word target and exact source.", limitation: "No new approval inferred." })
    },
  })
  expect((await runLean(directory, { sessions: scripted.port })).status).toBe("completed")
  expect(focusedRepairs(scripted)).toBe(1)
  expect(sawActual).toBe(true)
  expect(scripted.count(AGENT.changes)).toBe(initial!.counts.changes + 1)
  expect(scripted.count(AGENT.draft)).toBe(initial!.counts.draft)
  expect(scripted.count(AGENT.research)).toBe(initial!.counts.research)
  expect(keys.map(step => stageInputSha(directory, step))).toEqual(initial!.before)
  expect(readFileSync(leanFile(directory, "d1/body.docx"))).toEqual(initial!.bytes)
  expect(readLeanJson(directory, "checklist/extraction.json")).toEqual(initial!.extraction)
  const current = readDocxCommentParagraphs(readFileSync(leanFile(directory, "d1/body-applied.docx"))).map(row => row.paragraph_text)
  expect(current).toContain(initialDelivery)
  expect(current).toContain(payment)
  expect(current).not.toContain(omitted)
  expect(readLeanJson(directory, "repairs.json")).toMatchObject({ substantive_attempts: 1, items: [{ ref: { document: "d1", repair_target: "body" } }] })
}, 120000)

// A source adjudication that declines the one automatic checklist repair is not
// resubmitted: the re-checked finding is disclosed with the adjudicator's reason.
async function declinedRepairRun(reinspected: "unsatisfied" | "open" = "unsatisfied") {
  const directory = matter()
  const checks: string[] = []
  const outputs = () => ["d1/redline.docx", "d1/body-applied.docx", "drafting-notes.docx"].map(file => readFileSync(leanFile(directory, file)))
  const initial: { delivered?: ReturnType<typeof outputs>; applied?: string } = {}
  const scripted = scriptedModel({ ...happyPath,
    [AGENT["change-plan"]]: text => {
      const focused = text.includes("ONE supplied checklist assessment")
      if (focused) Object.assign(initial, { delivered: outputs(), applied: readLeanState(directory)!.stages["d1.apply"]!.started_at! })
      return JSON.stringify({ shared_decisions: [], groups: [], dismissed: [], questions: [],
        ...(focused ? { superseded_findings: [], superseded_decisions: [], findings: [{ disposition: "rejected", kind: "information",
          detail: "The assessed timing repair is not required.", reason: "DECLINED_REASON: the supplied terms already state the delivery period.", documents: ["d1"], sources: SUPPORT, limitations: [] }] } : {}) })
    },
    [CHECKLIST_JOB_AGENTS.extract]: text => {
      const input = JSON.parse(text.split("SOURCE_INPUT:\n\n")[1]!) as ChecklistProducerInput
      return JSON.stringify({ items: input.source.document === "request" ? [{ check: "The supplied delivery timing is stated.", basis: "directive",
        source_refs: [{ document: SOURCE, locator: "source paragraph", quote: SUPPORT[0]!.quote }], qualifications: [], target: { ...input.targets[0], treatment: "operative" } }] : [] })
    },
    [CHECKLIST_JOB_AGENTS.check]: text => {
      checks.push(text)
      const input = JSON.parse(text.split("CHECK_INPUT:\n\n")[1]!) as ChecklistCheckerInput
      const status = text.includes("GENERATED, UNVERIFIED COUNTEREVIDENCE") ? reinspected : "unsatisfied"
      return JSON.stringify({ item_id: input.item_id, status, reason: `Focused result ${status}.`, evidence: input.current_evidence,
        source_refs: input.item.source_refs, inspection_scope: "Current body paragraphs only.", limitation: "Comments not inspected." })
    },
  })
  // The unsatisfied finding starts the one checklist round without a Continue;
  // the adjudicator declines it with a source-cited reason and the re-check
  // stays unsatisfied.
  const state = await runLean(directory, { sessions: scripted.port })
  expect(focusedRepairs(scripted)).toBe(1)
  const round = readLeanJson(directory, "repairs.json")
  expect(round).toMatchObject({ round: 1, substantive_attempts: 1 })
  // The declined repair applies nothing; re-applying and re-rendering the
  // unchanged notes keep every Word output byte for byte.
  expect(outputs()).toEqual(initial.delivered!)
  expect(readLeanJson(directory, "d1/repair-adjudications.json")).toMatchObject({ assessments: [{ value: { groups: [], findings: [{ disposition: "rejected" }] } }] })
  expect(checks.some(text => text.includes("GENERATED, UNVERIFIED COUNTEREVIDENCE"))).toBe(false)
  return { directory, scripted, state, round, applied: initial.applied! }
}

test("a disclosed declined repair is not resubmitted and an unchanged Continue does nothing", async () => {
  const { directory, scripted, state, round } = await declinedRepairRun()
  expect(state.status).toBe("completed")
  expect(state.stages["checklist-check"]?.status).toBe("done")
  const settled = { prompts: scripted.prompts.length, report: readFileSync(leanFile(directory, "checklist/report.md"), "utf8"), gate: readLeanJson(directory, "gate.json") }
  enqueueLeanRequest(directory, { kind: "continue" })
  const again = await runLean(directory, { sessions: scripted.port })
  expect(scripted.prompts.length).toBe(settled.prompts)
  expect(again.status).toBe("completed")
  expect(readLeanJson(directory, "repairs.json")).toEqual(round)
  expect(readFileSync(leanFile(directory, "checklist/report.md"), "utf8")).toBe(settled.report)
  expect(readLeanJson(directory, "gate.json")).toEqual(settled.gate)
}, 120000)

test("an apply re-run from unchanged inputs keeps the Word outputs, first-applied revision dates and checklist records without a model call", async () => {
  const { directory, scripted, applied } = await declinedRepairRun()
  const files = ["d1/redline.docx", "d1/body-applied.docx", "drafting-notes.docx"]
  const before = Object.fromEntries(files.map(file => [file, readFileSync(leanFile(directory, file))]))
  const checks = readLeanJson(directory, "checklist/checks.json")
  // Tracked changes show when they were first applied, through the declined no-op apply.
  const dates = [...strFromU8(unzipSync(before["d1/redline.docx"]!)["word/document.xml"]!).matchAll(/<w:(?:ins|del)\b[^>]*\bw:date="([^"]*)"/g)].map(match => match[1])
  expect(dates.length).toBeGreaterThan(0)
  expect(new Set(dates)).toEqual(new Set([new Date(applied).toISOString().replace(/\.\d{3}Z$/, "Z")]))
  // An apply interrupted after writing its outputs re-runs from its inputs, a day later.
  updateLeanStage(directory, readLeanState(directory)!, "d1.apply", { status: "running", input_sha: null, outputs: [], finished_at: null })
  const prompts = scripted.prompts.length
  try {
    setSystemTime(new Date(Date.parse(applied) + 86_400_000))
    enqueueLeanRequest(directory, { kind: "continue" })
    const again = await runLean(directory, { sessions: scripted.port })
    expect(Date.parse(again.stages["d1.apply"]!.started_at!)).toBe(Date.parse(applied) + 86_400_000)
    expect(again.status).toBe("completed")
  } finally {
    setSystemTime()
  }
  expect(scripted.prompts.length).toBe(prompts)
  for (const file of files) expect(readFileSync(leanFile(directory, file))).toEqual(before[file]!)
  expect(readLeanJson(directory, "checklist/checks.json")).toEqual(checks)
}, 120000)

for (const reinspected of ["unsatisfied", "open"] as const) test(`at a blocked checklist, Continue reinspects a declined repair once instead of resubmitting it: ${reinspected}`, async () => {
  const { directory, scripted, round } = await declinedRepairRun(reinspected)
  settleBlockedChecklist(directory)
  const before = { prompts: scripted.prompts.length, plans: scripted.count(AGENT["change-plan"]) }
  enqueueLeanRequest(directory, { kind: "continue" })
  const state = await runLean(directory, { sessions: scripted.port })
  // Exactly one reinspection and no repair or drafting work.
  const added = scripted.prompts.slice(before.prompts)
  expect(added[0]!.agent).toBe(CHECKLIST_JOB_AGENTS.check)
  expect(added.filter(row => row.agent === CHECKLIST_JOB_AGENTS.check)).toHaveLength(1)
  expect(scripted.count(AGENT["change-plan"])).toBe(before.plans)
  expect(added[0]!.text).toContain("GENERATED, UNVERIFIED COUNTEREVIDENCE")
  expect(added[0]!.text).toContain("DECLINED_REASON")
  expect(readLeanJson(directory, "repairs.json")).toEqual(round)
  expect(readLeanJson(directory, "checklist/checks.json")).toMatchObject({ checks: [{ reinspection: { original_finding: { status: "unsatisfied" }, status: "succeeded" } }] })
  const done = readdirSync(leanFile(directory, "requests/done")).toSorted().map(name => readLeanJson(directory, `requests/done/${name}`) as { outcome: { status: string; detail?: string } })
  expect(done.at(-1)!.outcome).toEqual({ status: "resume", detail: "1 item(s) already declined by source adjudication are reinspected instead of resubmitted." })
  expect(state.status).toBe("completed")
  expect(readLeanJson(directory, "gate.json")).toMatchObject({ status: "pass", delivery: reinspected === "open" ? "review_draft" : "review_draft_with_findings", blocking: [] })
  if (reinspected === "unsatisfied")
    expect(readFileSync(leanFile(directory, "checklist/report.md"), "utf8")).toContain(": omission; source adjudication declined the repair: Focused result unsatisfied. Adjudicator's reason: DECLINED_REASON")
  else expect((readLeanJson(directory, "gate.json") as any).disclosed).toContainEqual(expect.objectContaining({ stage: "checklist-check", kind: "checklist_open", status: "disclosed", detail: expect.stringContaining("Finding: Focused result open.") }))
  // The reinspection is spent: a further Continue at a blocked checklist makes no call.
  settleBlockedChecklist(directory)
  const spent = scripted.prompts.length
  enqueueLeanRequest(directory, { kind: "continue" })
  await runLean(directory, { sessions: scripted.port })
  expect(scripted.prompts.slice(spent).filter(row => row.agent === CHECKLIST_JOB_AGENTS.check)).toHaveLength(0)
  expect(readLeanJson(directory, "repairs.json")).toEqual(round)
}, 120000)
