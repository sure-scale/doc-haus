import { Database } from "bun:sqlite"
import { afterEach, expect, test } from "bun:test"
import { readFileSync, readdirSync } from "node:fs"
import path from "node:path"
import { readDocxCommentParagraphs } from "../../lib/docx-comments"
import type { ChecklistCheckerInput } from "./checklist-contract"
import type { ChecklistDocumentRecord } from "./checklist-document-context"
import type { ChecklistCheckResult } from "./checklist-flow"
import { CHECKLIST_JOB_AGENTS } from "./checklist-job"
import { readLeanDocuments } from "./deliverables"
import { currentFactDisposition, withFactOptions } from "./fact-answers"
import { enqueueLeanRequest } from "./requests"
import { LEAN_MAX_SUBSTANTIVE_REPAIRS, readRepairRound, startRepairRound } from "./repairs"
import { runLean } from "./runner"
import { LEAN_PROMPT_VERSION, stageInputs } from "./stages"
import { hashStageInputs, leanFile, readLeanJson, readLeanState, readLeanText, readQuestions, updateLeanStage, writeLeanText, writeQuestions, type Question } from "./state"
import { AGENT, cleanupMatters, drafter, happyPath, matter, NO_FINDINGS, noChanges, registering, scriptedModel } from "./test-support"

afterEach(cleanupMatters)

const questionText = "What calibration window applies to the demonstration sensor?"
const permission = "Prepare the review record with the calibration window left unknown and the outstanding check disclosed. Do not issue or sign it."
const intakeText = "The intake entry awaits the calibration window [open item: obtain the calibration window before relying on this entry]."
const releaseText = "The release entry also awaits the calibration window [open item: verify the calibration window before release]."
const changedSource = "The new dispatch limit has not been supplied; obtain it before preparing the dispatch entry."
type Variant = "stale marked" | "current gate stale checklist" | "stale missing marker" | "current adverse" | "new wait answer" | "changed source"

for (const variant of ["stale marked", "current gate stale checklist", "stale missing marker", "current adverse", "new wait answer", "changed source"] satisfies Variant[]) {
  test(`Continue rechecks an obsolete blocked gate before body repair (${variant})`, async () => {
    const directory = matter("Both sensor entries depend on its calibration window, which has not been supplied.",
      "Prepare the demonstration-sensor intake and release record from the supplied information.")
    const offered = withFactOptions({ id: "q-sensor-window", stage: "intake", status: "open", kind: "fact", question: questionText, context: "The calibration window is not supplied." })
    const question: Question = { ...offered, status: "answered", answer: permission, answer_class: "answered", fact_status: "provisional" }
    writeQuestions(directory, [question])
    for (let i = 0; i < LEAN_MAX_SUBSTANTIVE_REPAIRS; i++) startRepairRound(directory, [], true)
    const missing = variant === "stale missing marker" || variant === "current adverse"
    const release = missing ? "The release entry also awaits the calibration window." : releaseText
    let recovering = false
    const first = new Map<string, ChecklistCheckerInput>()
    const registration = registering()
    const model = scriptedModel({ ...happyPath,
      [AGENT.draft]: drafter(`# Demonstration sensor record\n\n## Intake\n\n${intakeText}\n\n## Release\n\n${release}\n`),
      [AGENT.review]: () => NO_FINDINGS,
      [AGENT.changes]: noChanges,
      [AGENT.intake]: () => {
        const db = new Database(path.join(directory, ".dochaus/legal.db"), { readonly: true })
        let newLimit = false
        try { newLimit = !!db.query("SELECT 1 FROM chunks WHERE text = ?").get(changedSource) } finally { db.close() }
        if (newLimit) return JSON.stringify({ questions: [{ question: "What is the newly required dispatch limit?", kind: "fact", context: "The changed source requires a limit that has not been supplied." }] })
        return JSON.stringify({ ready: true, jurisdiction: "No external legal determination is needed for this supplied-source record", scope: "Supplied sensor information only", acting_for: { party: "", basis: "none" }, dates: { document_date: { date: "", basis: "today" }, operative_dates: [] } })
      },
      [AGENT.finalize]: () => "The calibration window remains unknown. Obtain it before relying on either sensor entry. Permission covers preparation of a review record only; issue and signature remain unapproved.",
      [CHECKLIST_JOB_AGENTS.check]: (text, _index, sessionID) => {
        const json = text.split("CHECK_INPUT:\n\n")[1]
        if (json) first.set(sessionID!, JSON.parse(json))
        const input = first.get(sessionID!)!
        const record = readLeanJson(directory, `checklist-context/${input.item.target.document}.json`) as ChecklistDocumentRecord
        const dependent = record.paragraphs.filter(row => /^(The intake entry|The release entry)/.test(row.paragraph_text))
        expect(dependent).toHaveLength(2)
        const marked = dependent.every(row => /\[open item: [^\]]*\bcalibration window\b[^\]]*\]/.test(row.paragraph_text))
        const checked = recovering || variant === "current adverse"
        return JSON.stringify({ item_id: input.item_id,
          status: checked ? marked ? "satisfied" : "unsatisfied" : "indeterminate",
          action_status: checked ? marked ? "fulfilled" : "not_fulfilled" : "indeterminate", open_decisions: [],
          ...(!checked ? { indeterminate_cause: "not_inspectable" } : {}),
          reason: !checked ? "The earlier assessment did not establish per-dependent-entry marker coverage."
            : marked ? "Both current dependent entries explicitly mark this unknown calibration window." : "The release entry depends on the unknown window but has no marker for it.",
          evidence: dependent.map(row => ({ document: input.item.target.document,
            locator: `native PIDs ${row.pids.join(", ")}; accepted paragraph_index ${row.paragraph_index}`, text: row.paragraph_text, kind: "body" })),
          source_refs: input.item.source_refs, inspection_scope: "All current intake and release entries dependent on this calibration window.",
          limitation: "The window remains unknown; preparation supplies no issue or signature approval.",
        })
      },
    })

    const firstBlocked = await runLean(directory, { sessions: model.port, ...registration })
    expect(firstBlocked.status).toBe("blocked")
    // The first gate records its factual-item resolution after notes were
    // keyed. Settle that existing bookkeeping from accepted section caches,
    // without a Continue/repair, before taking a coherent saved checkpoint.
    const blocked = await runLean(directory, { sessions: model.port, ...registration })
    expect(blocked.status).toBe("blocked")
    expect(blocked.stages.gate?.status).toBe("blocked")
    expect(blocked.stages["checklist-check"]?.status).toBe("done")
    expect(blocked.stages.finalize?.status).toBe("done")
    expect(blocked.stages.finalize?.input_sha).toBe(hashStageInputs("finalize", LEAN_PROMPT_VERSION, stageInputs(directory, "finalize")))
    expect(readLeanJson(directory, "gate.json")).toMatchObject({ status: "blocked", blocking: expect.any(Array) })
    expect(registration.names).toEqual([])

    // Emulate the actual upgrade entry: all Word/body work remains accepted,
    // while the old source-duty verdict and marker-only input keys are saved.
    // This does not turn a failed finalize into the blocked-gate case.
    if (variant !== "current adverse") {
      const checks = readLeanJson(directory, "checklist/checks.json") as { checks: ChecklistCheckResult[] }
      for (const row of checks.checks) if (row.finding) {
        row.input_sha = "prior-source-duty-check"
        row.finding = { ...row.finding, status: "not_applicable", action_status: "not_applicable", open_decisions: [questionText],
          reason: "The preparation permission did not itself impose the product's per-entry marker rule." }
      }
      writeLeanText(directory, "checklist/checks.json", JSON.stringify(checks))
      for (const step of ["checklist-check", "gate"] as const) {
        const inputs = stageInputs(directory, step).map(part => part.name === "provisional-fact-marker-contract"
          ? { ...part, bytes: "question-specific-provisional-markers-1" } : part)
        const oldKey = hashStageInputs(step, LEAN_PROMPT_VERSION, inputs)
        expect(oldKey).not.toBe(hashStageInputs(step, LEAN_PROMPT_VERSION, stageInputs(directory, step)))
        updateLeanStage(directory, readLeanState(directory)!, step, { input_sha: oldKey })
      }
      if (variant === "current gate stale checklist") {
        updateLeanStage(directory, readLeanState(directory)!, "gate", {
          input_sha: hashStageInputs("gate", LEAN_PROMPT_VERSION, stageInputs(directory, "gate")),
        })
      }
    }
    const saved = readLeanState(directory)!
    const prefix = Object.fromEntries(Object.entries(saved.stages).filter(([step, stage]) => stage?.status === "done" && !["checklist-check", "package-review"].includes(step))
      .map(([step, stage]) => [step, structuredClone(stage)]))
    const repairBytes = readLeanText(directory, "repairs.json")
    expect(readRepairRound(directory)).toMatchObject({ round: 5, substantive_attempts: 5 })
    const answers = readLeanText(directory, "questions.json")
    const body = readFileSync(leanFile(directory, "d1/body-applied.docx"))
    const paragraphs = readFileSync(leanFile(directory, "d1/paragraphs-applied.json"))
    const notes = readFileSync(leanFile(directory, "drafting-notes.docx"))
    const sections = new Map(readdirSync(leanFile(directory, "notes-sections")).map(file => [file, readFileSync(leanFile(directory, `notes-sections/${file}`))]))
    recovering = true
    if (variant === "new wait answer") enqueueLeanRequest(directory, { kind: "answer", question_id: question.id, option_id: "o2" })
    if (variant === "changed source") {
      const db = new Database(path.join(directory, ".dochaus/legal.db"))
      try { db.query("UPDATE chunks SET text = ?, char_end = ?").run(changedSource, changedSource.length) } finally { db.close() }
    }
    const before = model.prompts.length
    const request = enqueueLeanRequest(directory, { kind: "continue" })
    const resumed = await runLean(directory, { sessions: model.port, ...registration })
    const later = model.prompts.slice(before)
    if (variant === "current adverse") {
      expect(resumed.status).toBe("blocked")
      expect(registration.names).toEqual([])
      expect(readLeanJson(directory, `requests/done/${request.id}.json`)).toMatchObject({ outcome: { status: "repair_round" } })
      expect(later.some(prompt => prompt.agent === AGENT["change-plan"])).toBe(true)
      expect(readRepairRound(directory)?.substantive_attempts).toBe(5)
      return
    }
    expect(readLeanText(directory, "repairs.json")).toBe(repairBytes)
    expect(readLeanJson(directory, `requests/done/${request.id}.json`)).toMatchObject({ outcome: { status: "resume" } })
    expect(later.some(prompt => prompt.agent === AGENT["change-plan"])).toBe(false)
    expect(readFileSync(leanFile(directory, "d1/body-applied.docx"))).toEqual(body)
    expect(readFileSync(leanFile(directory, "d1/paragraphs-applied.json"))).toEqual(paragraphs)
    if (variant === "new wait answer" || variant === "changed source") {
      expect(resumed.status).toBe("blocked")
      expect(registration.names).toEqual([])
      if (variant === "new wait answer") expect(currentFactDisposition(directory, readQuestions(directory).find(row => row.id === question.id)!)).toBe("wait")
      else {
        expect(later.some(prompt => prompt.agent === AGENT.intake)).toBe(true)
        expect(readQuestions(directory).some(row => row.status === "open" && row.question === "What is the newly required dispatch limit?")).toBe(true)
      }
      return
    }
    expect(readLeanText(directory, "questions.json")).toBe(answers)
    expect(currentFactDisposition(directory, readQuestions(directory)[0]!)).toBe("provisional")
    for (const [step, value] of Object.entries(prefix)) expect(resumed.stages[step as keyof typeof resumed.stages]).toEqual(value)
    expect(readFileSync(leanFile(directory, "drafting-notes.docx"))).toEqual(notes)
    for (const [file, bytes] of sections) expect(readFileSync(leanFile(directory, `notes-sections/${file}`))).toEqual(bytes)
    expect(later.some(prompt => prompt.agent === CHECKLIST_JOB_AGENTS.check)).toBe(true)
    expect(later.every(prompt => new Set<string>([CHECKLIST_JOB_AGENTS.check, AGENT["package-review"]]).has(prompt.agent))).toBe(true)
    if (variant === "stale missing marker") {
      expect(resumed.status).toBe("blocked")
      expect(registration.names).toEqual([])
      return
    }
    expect(resumed.status).toBe("completed")
    expect(readLeanJson(directory, "gate.json")).toMatchObject({ status: "pass", delivery: "review_draft", blocking: [] })
    const documents = readLeanDocuments(directory)!
    expect(documents.status).toBe("completed")
    const deliveredBody = documents.documents.find(row => row.kind === "body")!.deliverable!
    expect(registration.names).toContain(deliveredBody)
    expect(readDocxCommentParagraphs(readFileSync(path.join(directory, deliveredBody))).map(row => row.paragraph_text)).toEqual(readDocxCommentParagraphs(body).map(row => row.paragraph_text))
    const deliveredNotes = documents.documents.find(row => row.kind === "drafting_notes")!.deliverable!
    const disclosed = readDocxCommentParagraphs(readFileSync(path.join(directory, deliveredNotes))).map(row => row.paragraph_text).join("\n")
    expect(disclosed).toContain("calibration window remains unknown")
    expect(disclosed).toContain("issue and signature remain unapproved")
  }, 120_000)
}
