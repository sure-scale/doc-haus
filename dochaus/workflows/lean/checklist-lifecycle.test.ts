import { afterEach, expect, test } from "bun:test"
import { existsSync, readFileSync, writeFileSync } from "node:fs"
import { renderMarkdownToDocx } from "../../lib/render"
import { runLean, stageInputSha } from "./runner"
import { stageOutputValid } from "./stages"
import { exportDocxComments, readDocxCommentParagraphs } from "../../lib/docx-comments"
import { leanFile, readLeanJson, readOpenItems, writeLeanBytes, writeLeanText } from "./state"
import { finalGate } from "./gate"
import { drainLeanRequests, enqueueLeanRequest } from "./requests"
import { CHECKLIST_JOB_AGENTS } from "./checklist-job"
import { checklistRepairItems, readRepairRound } from "./repairs"
import { AGENT, cleanupMatters, happyPath, matter, scriptedModel, drafter, pidOf, DRAFT, SOURCE, SUPPORT } from "./test-support"
import { checklistTextMeasuresMeaning, type ChecklistCheckerInput, type ChecklistProducerInput } from "./checklist-contract"

afterEach(cleanupMatters)

function model(directory: string, status: "satisfied" | "open" | "unsatisfied" | "indeterminate" = "satisfied") {
  let extractedBeforeBody = false
  let checkedAfterNotes = false
  const result = scriptedModel({
    ...happyPath,
    [CHECKLIST_JOB_AGENTS.extract]: text => {
      const input = JSON.parse(text.split("SOURCE_INPUT:\n\n")[1]!) as ChecklistProducerInput
      extractedBeforeBody = !existsSync(leanFile(directory, "d1/body.docx"))
      return JSON.stringify({ items: input.source.document === "request" ? [{
        check: "The requested supply note is drafted.", basis: "directive", source_refs: input.passages,
        qualifications: [], target: { ...input.targets[0], treatment: "operative" },
      }] : [] })
    },
    [CHECKLIST_JOB_AGENTS.check]: text => {
      const input = JSON.parse(text.split("CHECK_INPUT:\n\n")[1]!) as ChecklistCheckerInput
      checkedAfterNotes = existsSync(leanFile(directory, "drafting-notes.docx")) && existsSync(leanFile(directory, "d1/body-applied.docx"))
      return JSON.stringify({ item_id: input.item_id, status, reason: `Focused result ${status}.`, evidence: input.current_evidence,
        source_refs: input.item.source_refs, inspection_scope: "Current body paragraphs only.", limitation: "Comments not inspected." })
    },
  })
  return { ...result, early: () => extractedBeforeBody, late: () => checkedAfterNotes }
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

for (const status of ["unsatisfied", "indeterminate"] as const) test(`${status} blocks delivery with visible findings rather than omission disclosure`, async () => {
  const directory = matter()
  const scripted = model(directory, status)
  const state = await runLean(directory, { sessions: scripted.port })
  expect(state.status).toBe("blocked")
  expect(state.stages["checklist-check"]?.status).toBe("blocked")
  expect(state.stop_reason).toContain(status)
  expect(readOpenItems(directory).some(item => item.stage === "checklist-check" && item.status === "open")).toBe(true)
  expect(scripted.count(AGENT["package-review"])).toBe(0)
  expect(readFileSync(leanFile(directory, "checklist/report.md"), "utf8")).toContain("Target: d1; treatment: operative")
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


test("only explicit Continue reinspects blocked findings once without redrafting", async () => {
  const directory = matter()
  expect((await runLean(directory, { sessions: model(directory, "indeterminate").port })).status).toBe("blocked")
  const idle = model(directory, "satisfied")
  expect((await runLean(directory, { sessions: idle.port })).status).toBe("blocked")
  expect(idle.count(CHECKLIST_JOB_AGENTS.check)).toBe(0)
  enqueueLeanRequest(directory, { kind: "continue" })
  const review = model(directory, "indeterminate")
  const stillBlocked = await runLean(directory, { sessions: review.port })
  expect(stillBlocked.status).toBe("blocked")
  expect(stillBlocked.stop_reason).toContain("One reinspection has been used")
  expect(review.count(CHECKLIST_JOB_AGENTS.check)).toBe(1)
  expect(review.count(AGENT.draft)).toBe(0)
  expect(review.count(CHECKLIST_JOB_AGENTS.extract)).toBe(0)
  expect(readFileSync(leanFile(directory, "checklist/report.md"), "utf8")).toContain("Original finding: indeterminate")
  enqueueLeanRequest(directory, { kind: "continue" })
  const exhausted = model(directory, "satisfied")
  expect((await runLean(directory, { sessions: exhausted.port })).status).toBe("blocked")
  expect(exhausted.count(CHECKLIST_JOB_AGENTS.check)).toBe(0)
}, 120000)

test("a focused model review may resolve an adverse assessment without changing its item or Word bytes", async () => {
  const directory = matter()
  await runLean(directory, { sessions: model(directory, "indeterminate").port })
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
  expect((await runLean(directory, { sessions: model(directory, "indeterminate").port })).status).toBe("blocked")
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
// decides it. Unmarked all-not_applicable children keep base semantics.
for (const [status, marked, outcome] of [["satisfied", true, "completed"], ["indeterminate", true, "blocked"], ["not_applicable", true, "blocked"], ["not_applicable", false, "completed"]] as const) test(`full-draft check of a document-wide requirement: ${marked ? "deferred" : "undeferred"} ranges and ${status} document-properties finding leave the run ${outcome}`, async () => {
  const directory = matter()
  expect((await runLean(directory, { sessions: model(directory, "indeterminate").port })).status).toBe("blocked")
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
  expect(state.status).toBe(outcome)
  expect(state.stages["checklist-check"]?.status).toBe(outcome === "completed" ? "done" : "blocked")
  expect(inputs.filter(input => input.assignment?.kind === "document_properties")).toHaveLength(1)
  expect(inputs.filter(input => input.assignment?.kind === "process_records")).toHaveLength(1)
  expect(inputs.every(input => input.assignment)).toBe(true)
  const whole = inputs.find(input => input.assignment?.kind === "document_properties")!
  expect(whole.current_evidence.length).toBeGreaterThan(0)
  expect(whole.current_evidence.every(row => row.kind === "artifact" && "accepted_body_text_measures" in JSON.parse(row.text))).toBe(true)
  expect(full.count(CHECKLIST_JOB_AGENTS.extract)).toBe(0)
  expect(full.count(AGENT.draft)).toBe(0)
  if (outcome === "blocked") expect(state.stop_reason).toContain(`${parent.id}: indeterminate`)
  else expect(stageOutputValid(directory, "checklist-check")).toBe(true)
  if (outcome === "blocked" && status === "not_applicable") expect(state.stop_reason).toContain("no document-properties finding decided it")
}, 120000)

// Text measures can disprove but never establish a rendered-length directive:
// an approximate one stays open (disclosed in the report, review draft); a
// clearly exceeded one is unsatisfied and takes the existing repair route.
async function documentPropertiesRun(status: "open" | "unsatisfied") {
  const directory = matter()
  expect((await runLean(directory, { sessions: model(directory, "indeterminate").port })).status).toBe("blocked")
  const parent = (readLeanJson(directory, "checklist/extraction.json") as any).sources.flatMap((row: any) => row.items)[0]
  enqueueLeanRequest(directory, { kind: "refine_checklist", parents: [{ parent_id: parent.id, scope: "target_body" }] })
  const seen = { measures: "", guided: false }
  const full = scriptedModel({ ...happyPath,
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
  return { directory, parent, state: await runLean(directory, { sessions: full.port }), seen }
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

test("a rendered-length directive clearly exceeded by text measures blocks and offers the existing focused repair", async () => {
  const { directory, parent, state, seen } = await documentPropertiesRun("unsatisfied")
  expect(state.status).toBe("blocked")
  expect(state.stages["checklist-check"]?.status).toBe("blocked")
  expect(state.stop_reason).toContain(`${parent.id}: unsatisfied`)
  const repairs = checklistRepairItems(directory)
  expect(repairs).toHaveLength(1)
  expect(repairs[0]).toMatchObject({ stage: "checklist-check", ref: { document: "d1", repair_target: "body" } })
  expect(repairs[0]!.detail).toContain(seen.measures.replaceAll('"', '\\"'))
  enqueueLeanRequest(directory, { kind: "continue" })
  drainLeanRequests(directory, state)
  expect(readRepairRound(directory)?.items).toEqual(repairs)
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


test("Continue repairs a source-backed omission through cumulative changes while preserving initial drafting", async () => {
  const directory = matter()
  const initialDelivery = "The supplier shall deliver the goods within thirty days of the order date."
  const omitted = "The supplier shall deliver the goods."
  const payment = "Payment shall be due on delivery."
  const initialDraft = DRAFT.replace(omitted, initialDelivery)
  let repairing = false, sawActual = false
  const scripted = scriptedModel({ ...happyPath,
    [AGENT.draft]: drafter(initialDraft),
    [AGENT.review]: () => "No further findings.",
    [AGENT["change-plan"]]: text => {
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
      return JSON.stringify({ item_id: packet.item_id, status: repairing ? "satisfied" : "unsatisfied", reason: repairing ? "Current operative timing is present." : "The cumulative amendment removed the supplied timing.", evidence: packet.current_evidence, source_refs: packet.item.source_refs, inspection_scope: "Current native Word target and exact source.", limitation: "No new approval inferred." })
    },
  })
  expect((await runLean(directory, { sessions: scripted.port })).status).toBe("blocked")
  const initialBytes = readFileSync(leanFile(directory, "d1/body.docx"))
  const extraction = readLeanJson(directory, "checklist/extraction.json")
  const keys = ["d1.draft", "d1.review", "research"] as const
  const before = keys.map(step => stageInputSha(directory, step))
  // Deliberately stale accepted-text sidecar must not substitute for current Word.
  const oldRows = readLeanJson(directory, "d1/paragraphs-applied.json") as { text: string }[]
  writeLeanText(directory, "d1/paragraphs-applied.json", JSON.stringify(oldRows.map(row => ({ ...row, text: row.text === omitted ? initialDelivery : row.text }))))
  const counts = { changes: scripted.count(AGENT.changes), draft: scripted.count(AGENT.draft), research: scripted.count(AGENT.research) }
  repairing = true
  enqueueLeanRequest(directory, { kind: "continue" })
  expect((await runLean(directory, { sessions: scripted.port })).status).toBe("completed")
  expect(sawActual).toBe(true)
  expect(scripted.count(AGENT.changes)).toBe(counts.changes + 1)
  expect(scripted.count(AGENT.draft)).toBe(counts.draft)
  expect(scripted.count(AGENT.research)).toBe(counts.research)
  expect(keys.map(step => stageInputSha(directory, step))).toEqual(before)
  expect(readFileSync(leanFile(directory, "d1/body.docx"))).toEqual(initialBytes)
  expect(readLeanJson(directory, "checklist/extraction.json")).toEqual(extraction)
  const current = readDocxCommentParagraphs(readFileSync(leanFile(directory, "d1/body-applied.docx"))).map(row => row.paragraph_text)
  expect(current).toContain(initialDelivery)
  expect(current).toContain(payment)
  expect(current).not.toContain(omitted)
  expect(readLeanJson(directory, "repairs.json")).toMatchObject({ substantive_attempts: 1, items: [{ ref: { document: "d1", repair_target: "body" } }] })
}, 120000)
