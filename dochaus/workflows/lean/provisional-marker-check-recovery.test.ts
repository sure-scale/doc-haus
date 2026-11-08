import { afterEach, expect, test } from "bun:test"
import { existsSync, readFileSync } from "node:fs"
import path from "node:path"
import { readDocxCommentParagraphs } from "../../lib/docx-comments"
import { renderMarkdownToDocx } from "../../lib/render"
import { withParagraphPointers } from "../../lib/drafting-amendment"
import type { ChecklistCheckerInput } from "./checklist-contract"
import type { ChecklistDocumentRecord } from "./checklist-document-context"
import { runChecklistCheckFlow, runChecklistExtractionFlow, type ChecklistCheckResult } from "./checklist-flow"
import { CHECKLIST_JOB_AGENTS } from "./checklist-job"
import { readLeanDocuments } from "./deliverables"
import { currentFactDisposition, withFactOptions } from "./fact-answers"
import { writeLeanPlan } from "./plan"
import { enqueueLeanRequest } from "./requests"
import { LEAN_MAX_SUBSTANTIVE_REPAIRS, readRepairRound, startRepairRound } from "./repairs"
import { runLean } from "./runner"
import { leanFile, readLeanJson, readLeanText, readQuestions, writeLeanBytes, writeLeanText, writeQuestions, type Question } from "./state"
import { AGENT, cleanupMatters, drafter, happyPath, matter, NO_FINDINGS, noChanges, registering, scriptedModel } from "./test-support"

afterEach(cleanupMatters)

const questionText = "What batch label belongs to the rehearsal equipment?"
const permission = "Prepare a review draft with the batch label left unknown and the outstanding check disclosed. This does not authorise dispatch or signature."
const receiving = "The receiving record awaits the batch label [open item: obtain the batch label before using this record]."
const dispatch = "The dispatch reference also awaits the batch label [open item: verify the batch label before dispatch]."
type CheckPacket = ChecklistCheckerInput & { verification?: "provisional_fact_marker" }
type Variant = "marked" | "missing dependent marker" | "marker for another unknown" | "wait"

for (const variant of ["marked", "missing dependent marker", "marker for another unknown", "wait"] satisfies Variant[]) {
  test(`saved final-notes recovery checks provisional handling before registering the review draft (${variant})`, async () => {
    const directory = matter("The receiving and dispatch records refer to the rehearsal batch label, which has not been supplied.",
      "Prepare a rehearsal-equipment handover recording the receiving and dispatch references from the supplied information.")
    const offered = withFactOptions({ id: "q-rehearsal-label", stage: "intake", status: "open", kind: "fact", question: questionText, context: "The batch label is absent from the supplied information." })
    const question: Question = { ...offered, status: "answered", answer: permission, answer_class: "answered", fact_status: "provisional" }
    writeQuestions(directory, [question])
    for (let i = 0; i < LEAN_MAX_SUBSTANTIVE_REPAIRS; i++) startRepairRound(directory, [], true)
    const finalDispatch = variant === "missing dependent marker" ? "The dispatch reference also awaits the batch label."
      : variant === "marker for another unknown" ? "The dispatch reference also awaits the batch label [open item: obtain the archive room number]." : dispatch
    const draft = `# Rehearsal equipment handover\n\n## Receiving\n\n${receiving}\n\n## Dispatch\n\n${finalDispatch}\n`
    const registration = registering()
    let recovering = false
    const packets: CheckPacket[] = []
    const firstChecks = new Map<string, CheckPacket>()
    const model = scriptedModel({ ...happyPath,
      [AGENT.draft]: drafter(draft),
      [AGENT.review]: () => NO_FINDINGS,
      [AGENT.changes]: noChanges,
      [AGENT.finalize]: () => {
        if (!recovering) throw new Error("Synthetic interruption before final notes were saved")
        return "The batch label remains unknown. Obtain it before relying on either the receiving or dispatch reference. Permission covers preparation of a review draft only; dispatch and signature remain unapproved."
      },
      [CHECKLIST_JOB_AGENTS.check]: (text, _index, sessionID) => {
        const raw = text.split("CHECK_INPUT:\n\n")[1]
        if (raw) firstChecks.set(sessionID!, JSON.parse(raw))
        const packet = firstChecks.get(sessionID!)!
        expect(packet).toBeDefined()
        packets.push(packet)
        // This source-duty interpretation reproduces the real contract defect:
        // permission to prepare does not itself impose a per-passage product rule.
        if (packet.verification !== "provisional_fact_marker") return JSON.stringify({
          item_id: packet.item_id, status: "not_applicable", action_status: "not_applicable",
          reason: "The quoted preparation permission does not impose the product's per-dependent-passage marker rule.",
          open_decisions: [questionText], evidence: packet.current_evidence.filter(row => row.kind === "body"),
          source_refs: packet.item.source_refs, inspection_scope: "The current receiving and dispatch references.", limitation: "The factual label remains unknown.",
        })
        // Simulate targeted read-record inspection of the actual current DOCX
        // record, not a pass selected by a reserved item ID or a marker count.
        const record = readLeanJson(directory, `checklist-context/${packet.item.target.document}.json`) as ChecklistDocumentRecord
        const dependent = record.paragraphs.filter(row => /^(The receiving record|The dispatch reference)/.test(row.paragraph_text))
        expect(dependent).toHaveLength(2)
        const marked = dependent.every(row => /\[open item: [^\]]*\bbatch label\b[^\]]*\]/.test(row.paragraph_text))
        const evidence = dependent.map(row => ({ document: packet.item.target.document,
          locator: `native PIDs ${row.pids.join(", ")}; accepted paragraph_index ${row.paragraph_index}`,
          text: row.paragraph_text, kind: "body" }))
        return JSON.stringify({ item_id: packet.item_id, status: marked ? "satisfied" : "unsatisfied",
          action_status: marked ? "fulfilled" : "not_fulfilled", open_decisions: [], evidence, source_refs: packet.item.source_refs,
          reason: marked ? "Both fact-dependent references visibly preserve this exact unknown label. Its truth and release approval remain unsettled."
            : "A dependent reference lacks an explicit marker for the unknown batch label; another unknown's marker cannot substitute.",
          inspection_scope: "Every receiving and dispatch paragraph in the current accepted body.",
          limitation: "This assesses treatment of an unknown fact, not its answer or authority to dispatch or sign.",
        })
      },
    })

    const failed = await runLean(directory, { sessions: model.port, ...registration })
    expect(failed.status).toBe("failed")
    expect(failed.stages.finalize?.status).toBe("failed")
    expect(failed.stages["d1.apply"]?.status).toBe("done")
    expect(registration.names).toEqual([])
    const prefix = Object.fromEntries(Object.entries(failed.stages).filter(([, value]) => value?.status === "done")
      .map(([step, value]) => [step, { input_sha: value!.input_sha, finished_at: value!.finished_at }]))
    const body = readFileSync(leanFile(directory, "d1/body-applied.docx"))
    const paragraphs = readFileSync(leanFile(directory, "d1/paragraphs-applied.json"))
    const savedQuestions = readLeanText(directory, "questions.json")
    const repairs = readRepairRound(directory)
    expect(repairs?.substantive_attempts).toBe(LEAN_MAX_SUBSTANTIVE_REPAIRS)
    expect(readDocxCommentParagraphs(body).some(row => row.paragraph_text === receiving)).toBe(true)
    expect(existsSync(leanFile(directory, "gate.json"))).toBe(false)

    recovering = true
    if (variant === "wait") enqueueLeanRequest(directory, { kind: "answer", question_id: question.id, option_id: "o2" })
    else enqueueLeanRequest(directory, { kind: "continue" })
    const resumed = await runLean(directory, { sessions: model.port, ...registration })
    expect(readRepairRound(directory)).toEqual(repairs)
    expect(readFileSync(leanFile(directory, "d1/body-applied.docx"))).toEqual(body)
    expect(readFileSync(leanFile(directory, "d1/paragraphs-applied.json"))).toEqual(paragraphs)
    const current = readQuestions(directory).find(row => row.id === question.id)!
    expect(currentFactDisposition(directory, current)).toBe(variant === "wait" ? "wait" : "provisional")
    if (variant === "wait") {
      expect(resumed.status).toBe("blocked")
      expect(registration.names).toEqual([])
      expect(packets).toHaveLength(0)
      return
    }
    expect(readLeanText(directory, "questions.json")).toBe(savedQuestions)
    for (const [step, value] of Object.entries(prefix)) expect(resumed.stages[step as keyof typeof resumed.stages]).toMatchObject(value)
    expect(packets.length).toBeGreaterThan(0)
    if (variant !== "marked") {
      expect(resumed.status).toBe("blocked")
      expect(registration.names).toEqual([])
      expect(readLeanJson(directory, "gate.json")).toMatchObject({ status: "blocked" })
      return
    }
    expect(resumed.status).toBe("completed")
    expect(packets.every(packet => packet.verification === "provisional_fact_marker")).toBe(true)
    expect(readLeanJson(directory, "gate.json")).toMatchObject({ status: "pass", delivery: "review_draft", blocking: [] })
    const checks = readLeanJson(directory, "checklist/checks.json") as { checks: { finding?: { status: string; open_decisions?: string[]; action_status?: string } }[] }
    expect(checks.checks.every(row => row.finding?.status === "satisfied" && row.finding.action_status === "fulfilled" && row.finding.open_decisions?.length === 0)).toBe(true)
    expect(readLeanText(directory, "drafting-notes.md")).toContain("batch label remains unknown")
    const delivered = readLeanDocuments(directory)!
    expect(delivered.status).toBe("completed")
    for (const document of delivered.documents) if (document.deliverable) {
      expect(registration.names).toContain(document.deliverable)
      const bytes = readFileSync(path.join(directory, document.deliverable))
      expect(bytes.subarray(0, 2).toString()).toBe("PK")
      const text = readDocxCommentParagraphs(bytes).map(row => row.paragraph_text).join("\n")
      if (document.kind === "body") {
        expect(text).toContain(receiving)
        expect(text).toContain(dispatch)
      }
      if (document.kind === "drafting_notes") {
        expect(text).toContain("batch label remains unknown")
        expect(text).toContain("dispatch and signature remain unapproved")
      }
    }
    expect(registration.names.length).toBeGreaterThan(0)
  }, 120_000)
}

for (const legacyInvalid of [true, false]) test(`marker contract recovery preserves a consumed reinspection (${legacyInvalid ? "invalid generic result gets one original check" : "valid adverse result gets no appeal"})`, async () => {
  const directory = matter("The rehearsal batch label has not been supplied.", "Prepare a receiving and dispatch reference record.")
  writeLeanPlan(directory, [{ title: "Reference record", purpose: "Record receiving and dispatch references.", file: "reference-record.docx", parts: [{ heading: "References", scope: "References dependent on the batch label." }] }])
  const offered = withFactOptions({ id: "q-rehearsal-label", stage: "intake", status: "open", kind: "fact", question: questionText, context: "The label is not supplied." })
  writeQuestions(directory, [{ ...offered, status: "answered", answer: permission, answer_class: "answered", fact_status: "provisional" }])
  const bytes = withParagraphPointers((await renderMarkdownToDocx(`${receiving}\n\n${legacyInvalid ? dispatch : "The dispatch reference awaits the batch label."}`)).bytes)
  writeLeanBytes(directory, "d1/body.docx", bytes)
  const questions = readLeanText(directory, "questions.json")
  let recovery = false
  const first = new Map<string, CheckPacket>()
  const model = scriptedModel({
    [CHECKLIST_JOB_AGENTS.check]: (text, _index, sessionID) => {
      const json = text.split("CHECK_INPUT:\n\n")[1]
      if (json) first.set(sessionID!, JSON.parse(json))
      const input = first.get(sessionID!)!
      const record = readLeanJson(directory, "checklist-context/d1.json") as ChecklistDocumentRecord
      const dependent = record.paragraphs.filter(row => /^(The receiving record|The dispatch reference)/.test(row.paragraph_text))
      const marked = dependent.length === 2 && dependent.every(row => /\[open item: [^\]]*\bbatch label\b[^\]]*\]/.test(row.paragraph_text))
      const fulfilled = recovery && input.verification === "provisional_fact_marker" && marked
      return JSON.stringify({ item_id: input.item_id, status: fulfilled ? "satisfied" : legacyInvalid ? "indeterminate" : "unsatisfied",
        action_status: fulfilled ? "fulfilled" : legacyInvalid ? "indeterminate" : "not_fulfilled", open_decisions: [],
        ...(!fulfilled && legacyInvalid ? { indeterminate_cause: "missing_inspection" } : {}),
        evidence: fulfilled ? input.current_evidence.filter(row => row.kind === "body") : [], source_refs: input.item.source_refs,
        reason: fulfilled ? "Both dependent references visibly mark this unknown label." : legacyInvalid ? "The marker coverage inspection is incomplete." : "The dispatch paragraph has no marker for the unknown label.",
        inspection_scope: "The current receiving and dispatch reference paragraphs.", limitation: "The factual label itself remains unknown; no release is approved.",
      })
    },
  })
  const deps = { sessions: model.port, model: null }
  await runChecklistExtractionFlow(directory, deps)
  expect((await runChecklistCheckFlow(directory, deps)).status).toBe("succeeded")
  const reviewed = await runChecklistCheckFlow(directory, deps, { reinspectAdverse: true })
  expect(reviewed.checks).toHaveLength(1)
  const consumed = structuredClone(reviewed.checks[0]!)
  expect(consumed.reinspection?.status).toBe("succeeded")
  if (legacyInvalid) {
    consumed.finding = { ...consumed.finding!, status: "not_applicable", action_status: "not_applicable",
      open_decisions: [questionText], reason: "The source permission does not itself require per-passage markers." }
    consumed.reinspection!.original_finding = structuredClone(consumed.finding)
  }
  consumed.input_sha = "prior-generic-check-contract"
  const inspection = structuredClone(consumed.reinspection)
  writeLeanText(directory, "checklist/checks.json", JSON.stringify({ extraction_complete: true, expected_items: [consumed.item_id], checks: [consumed] }))
  recovery = true
  const before = model.count(CHECKLIST_JOB_AGENTS.check)
  const recovered = await runChecklistCheckFlow(directory, deps, { reinspectAdverse: true })
  expect(model.count(CHECKLIST_JOB_AGENTS.check) - before).toBe(legacyInvalid ? 1 : 0)
  expect(recovered.checks[0]!.reinspection).toEqual(inspection)
  expect(recovered.checks[0]!.finding?.status).toBe(legacyInvalid ? "satisfied" : "unsatisfied")
  const saved = readLeanJson(directory, "checklist/checks.json") as { checks: ChecklistCheckResult[] }
  const repeated = await runChecklistCheckFlow(directory, deps, { reinspectAdverse: true })
  expect(model.count(CHECKLIST_JOB_AGENTS.check) - before).toBe(legacyInvalid ? 1 : 0)
  expect(repeated.checks).toEqual(saved.checks)
  expect(readLeanText(directory, "questions.json")).toBe(questions)
  expect(readFileSync(leanFile(directory, "d1/body.docx"))).toEqual(Buffer.from(bytes))
})
