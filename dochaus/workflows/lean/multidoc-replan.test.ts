import { afterEach, describe, expect, test } from "bun:test"
import { createHash } from "node:crypto"
import { appendFileSync, existsSync, readFileSync, writeFileSync } from "node:fs"
import path from "node:path"
import { checkChanges } from "./change-checks"
import { readParagraphs } from "./checks"
import { readLeanDocuments } from "./deliverables"
import { readLeanPlan, writeLeanPlan } from "./plan"
import { restoreLeanStage } from "./restore-stage"
import { runLean } from "./runner"
import { outlineBasis, planIsCurrent } from "./stages"
import { answerQuestion, leanFile, readLeanJson, readOpenItems, readQuestions, writeQuestions, writeOpenItems, replaceStageQuestions, replaceStageOpenItems } from "./state"
import {
  AGENT, APPLIED, DRAFT, DRAFTED, LETTER, QUESTION, SUPPORT, changesFirst, cleanupMatters, drafter, happyPath, listed, matter, noChanges, pidIn,
  registering, reviewFirst, scriptedModel, type Reply,
} from "./test-support"

// Runs whose earlier steps change after a document has finished: a question
// that a later document stops asking, lawyer's edits to the plan, and a re-plan
// that puts the documents in another order. Real coordinator, checkers,
// renderer, applier and gate; only the model is scripted.

const MARK = "[open item: whether late delivery carries a remedy]"
const CONTEXT = "The terms are silent."

afterEach(cleanupMatters)

/** Legacy runs could finish with one unresolved question shared by two
 * documents. Keep that state readable while the new runner now pauses it. */
async function askedTwice() {
  const directory = matter()
  const replies = { ...happyPath, [AGENT.draft]: drafter(DRAFT, LETTER), [AGENT.review]: reviewFirst, [AGENT.changes]: changesFirst }
  expect((await runLean(directory, { sessions: scriptedModel(replies).port })).status).toBe("completed")
  const notePid = readParagraphs(directory, "d1").find(row => row.text === DRAFTED)!.pid
  const letterPid = readParagraphs(directory, "d2")[0]!.pid
  const question = { id: "q-shared", stage: "changes" as const, document: "d2" as const, pid: letterPid, question: QUESTION, context: CONTEXT, status: "open" as const }
  writeQuestions(directory, [question])
  writeOpenItems(directory, [
    { id: "question-d1", stage: "changes", kind: "question", detail: QUESTION, ref: { document: "d1", pid: notePid, question_id: question.id }, status: "open" },
    { id: "question-d2", stage: "changes", kind: "question", detail: QUESTION, ref: { document: "d2", pid: letterPid, question_id: question.id }, status: "open" },
  ])
  return { directory, replies, question, notePid }
}

function sha(file: string) {
  return createHash("sha256").update(readFileSync(file)).digest("hex")
}

describe("lean drafting when a finished document's steps change", () => {
  test("a question the later document stops asking stays open with the earlier document that still asks it", async () => {
    const run = await askedTwice()
    replaceStageOpenItems(run.directory, "changes", "d2", [])
    replaceStageQuestions(run.directory, "changes", "d2", [])
    const model = scriptedModel(run.replies)
    const state = await runLean(run.directory, { sessions: model.port })

    expect(state.status).toBe("blocked")
    expect(model.prompts).toHaveLength(0)
    expect(readQuestions(run.directory)).toEqual([expect.objectContaining({ id: run.question.id, document: "d1", pid: run.notePid, status: "open" })])
    expect(state.stop_reason).toContain("Waiting for your answers")
    expect(answerQuestion(run.directory, run.question.id, "No remedy.")).toMatchObject({ status: "answered" })
  }, 120_000)

  test("restoring to before the later document's changes keeps the question with the earlier document", async () => {
    const run = await askedTwice()
    restoreLeanStage(run.directory, "d2.review")
    expect(readQuestions(run.directory)).toEqual([expect.objectContaining({ id: run.question.id, document: "d1", pid: run.notePid, status: "open" })])

    const model = scriptedModel(run.replies)
    const state = await runLean(run.directory, { sessions: model.port })
    expect(state.status).toBe("blocked")
    expect(model.prompts).toHaveLength(0)
    expect(readQuestions(run.directory)).toEqual([expect.objectContaining({ id: run.question.id, document: "d1", status: "open" })])
  }, 120_000)

  test("a lawyer's plan without the second document drops its steps, open items and summary rows; its registered files stay", async () => {
    const directory = matter()
    const deliverables = registering()
    const letter = `${LETTER.trimEnd()}\n\n## Address\n\nDeliver to [open item: confirm the delivery address].\n`
    const replies = { ...happyPath, [AGENT.draft]: drafter(DRAFT, letter), [AGENT.review]: reviewFirst, [AGENT.changes]: changesFirst }
    const first = await runLean(directory, { sessions: scriptedModel(replies).port, registerDeliverable: deliverables.registerDeliverable })
    expect(first.status).toBe("completed")
    expect(readOpenItems(directory).some(item => item.ref.document === "d2")).toBe(true)

    const plan = JSON.parse(readFileSync(leanFile(directory, "plan.json"), "utf8"))
    writeFileSync(leanFile(directory, "plan.json"), `${JSON.stringify({ ...plan, documents: plan.documents.slice(0, 1) }, null, 2)}\n`)
    const failing = scriptedModel({ ...replies, [AGENT.finalize]: () => "" })
    const failed = await runLean(directory, { sessions: failing.port, registerDeliverable: deliverables.registerDeliverable })
    expect(failed.status).toBe("failed")
    // d1's review and changes are scoped by the planned documents, so dropping d2 re-reviews d1; its draft is kept.
    expect(failing.counts()).toEqual([0, 1, 1, 4])
    expect(Object.keys(failed.stages)).toEqual(["discrepancy", "intake", "research", "research-review", "spine", "d1.draft", "d1.review", "d1.review-independent", "d1.review-wording", "d1.review-client-disadvantage", "d1.change-plan", "d1.change-prepare", "d1.changes", "d1.change-review", "d1.apply", "finalize", "checklist-check", "package-review", "gate"])
    expect(readOpenItems(directory).filter(item => item.ref.document === "d2")).toEqual([])
    expect(readLeanDocuments(directory)!.documents.map(row => row.document)).toEqual(["d1", "d1", null, null])
    expect(existsSync(path.join(directory, "Covering letter.docx"))).toBe(true)

    const resumed = await runLean(directory, { sessions: scriptedModel(replies).port, registerDeliverable: deliverables.registerDeliverable })
    expect(resumed.status).toBe("completed")
    expect(readLeanJson(directory, "gate.json")).toMatchObject({ status: "pass", blocking: [] })
    expect(readLeanDocuments(directory)!.documents.map(row => `${row.kind}:${row.document}:${row.deliverable}`)).toEqual([
      "body:d1:Supply note.docx", "redline:d1:Supply note - redline.docx", "drafting_notes:null:Supply note - drafting notes.docx", "drafting_notes_markdown:null:null",
    ])
  }, 180_000)

  test("a lawyer's plan that drops the first document and renumbers the second redrafts it from its entry under its own name", async () => {
    const directory = matter()
    const deliverables = registering()
    const replies = { ...happyPath, [AGENT.draft]: drafter(DRAFT, LETTER), [AGENT.review]: reviewFirst, [AGENT.changes]: changesFirst }
    const first = await runLean(directory, { sessions: scriptedModel(replies).port, registerDeliverable: deliverables.registerDeliverable })
    expect(first.status).toBe("completed")

    const plan = JSON.parse(readFileSync(leanFile(directory, "plan.json"), "utf8"))
    writeFileSync(leanFile(directory, "plan.json"), `${JSON.stringify({ ...plan, documents: [{ ...plan.documents[1], id: "d1" }] }, null, 2)}\n`)
    const model = scriptedModel(replies)
    const state = await runLean(directory, { sessions: model.port, registerDeliverable: deliverables.registerDeliverable })

    expect(state.status).toBe("completed")
    expect(model.counts()).toEqual([2, 1, 1, 1])
    expect(model.prompts.some(prompt => prompt.text.includes("Reply now with the outline only"))).toBe(false)
    expect(readLeanPlan(directory)?.documents.map(row => [row.id, row.title])).toEqual([["d1", "Covering letter"]])
    expect(readFileSync(leanFile(directory, "d1/draft.md"), "utf8")).toStartWith("# Covering letter\n")
    expect(readLeanDocuments(directory)!.documents.map(row => `${row.kind}:${row.document}:${row.title}:${row.deliverable}`)).toEqual([
      "body:d1:Covering letter:Covering letter.docx",
      "redline:d1:Covering letter:Covering letter - redline.docx",
      "drafting_notes:null:null:Covering letter - drafting notes.docx",
      "drafting_notes_markdown:null:null:null",
    ])
    expect(sha(path.join(directory, "Covering letter.docx"))).toBe(sha(leanFile(directory, "d1/body-applied.docx")))
    expect(existsSync(path.join(directory, "Covering letter (2).docx"))).toBe(false)
  }, 180_000)

  test("a re-plan that reorders the documents registers each body under its own title", async () => {
    const directory = matter()
    const deliverables = registering()
    const replies = { ...happyPath, [AGENT.review]: reviewFirst, [AGENT.changes]: changesFirst }
    const first = await runLean(directory, { sessions: scriptedModel({ ...replies, [AGENT.draft]: drafter(DRAFT, LETTER) }).port, registerDeliverable: deliverables.registerDeliverable })
    expect(first.status).toBe("completed")

    writeFileSync(leanFile(directory, "request.md"), "Draft a covering letter enclosing a short supply note from the supply terms.\n")
    const replanned = await runLean(directory, { sessions: scriptedModel({ ...replies, [AGENT.draft]: drafter(LETTER, DRAFT) }).port, registerDeliverable: deliverables.registerDeliverable })
    expect(replanned.status).toBe("completed")
    expect(readLeanPlan(directory)?.documents.map(row => [row.id, row.title])).toEqual([["d1", "Covering letter"], ["d2", "Supply note"]])
    expect(readLeanDocuments(directory)!.documents.slice(0, 4).map(row => `${row.kind}:${row.document}:${row.title}:${row.deliverable}`)).toEqual([
      "body:d1:Covering letter:Covering letter.docx",
      "redline:d1:Covering letter:Covering letter - redline.docx",
      "body:d2:Supply note:Supply note.docx",
      "redline:d2:Supply note:Supply note - redline.docx",
    ])
    expect(sha(path.join(directory, "Covering letter.docx"))).toBe(sha(leanFile(directory, "d1/body-applied.docx")))
    expect(sha(path.join(directory, "Supply note.docx"))).toBe(sha(leanFile(directory, "d2/body-applied.docx")))
  }, 180_000)

  test("a saved plan is followed until the request it was outlined from changes; a plan without its basis is followed as written", () => {
    const directory = matter()
    const parts = [{ heading: "Delivery", scope: "" }]
    expect(planIsCurrent(directory)).toBe(false)
    writeLeanPlan(directory, [{ title: "Supply note", purpose: "", parts }], outlineBasis(directory))
    expect(planIsCurrent(directory)).toBe(true)
    writeFileSync(leanFile(directory, "request.md"), "Draft a covering letter for the supply note.\n")
    expect(planIsCurrent(directory)).toBe(false)
    writeLeanPlan(directory, [{ title: "Supply note", purpose: "", parts }])
    expect(planIsCurrent(directory)).toBe(true)
  })

  test("a question repeated by id without a paragraph keeps the earlier paragraph only in the same document", () => {
    const directory = matter()
    writeQuestions(directory, [{ id: "q-earlier", stage: "changes", question: QUESTION, context: CONTEXT, pid: "p000000000001", document: "d1", status: "open" }])
    const reply = JSON.stringify({ amendments: [], questions: [{ id: "q-earlier", question: QUESTION }] })
    const rows = [{ pid: "p000000000002", sha256: "0".repeat(64), text: "Delivery follows the supply note." }]
    const asked = (document: "d1" | "d2") => checkChanges(directory, reply, rows, { withholdUnverified: true, document }).changes?.questions

    expect(asked("d1")).toEqual([expect.objectContaining({ id: "q-earlier", pid: "p000000000001" })])
    expect(asked("d2")?.map(question => [question.id, question.pid])).toEqual([["q-earlier", undefined]])
    expect(checkChanges(directory, reply, rows, { document: "d2" }).issues.join(" ")).toContain("is tied to no paragraph")
  })
})
