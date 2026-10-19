import { afterEach, expect, test } from "bun:test"
import { existsSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import path from "node:path"
import { readLeanDocuments } from "./deliverables"
import { readLeanEvents } from "./events"
import { writeLeanPlan } from "./plan"
import { enqueueLeanRequest, pendingLeanRequests } from "./requests"
import { restoreLeanStage } from "./restore-stage"
import { runLean } from "./runner"
import { LEAN_STAGE_AGENTS } from "./stages"
import { leanDir, leanFile, readLeanJson, readLeanState, readOpenItems, readQuestions } from "./state"
import { changes, happyLeanReplies, leanFixtureMatter, sameSession, scriptedLeanModel } from "./test-support"

const directories: string[] = []
afterEach(() => directories.splice(0).forEach(directory => rmSync(directory, { recursive: true, force: true })))

test("restoring to just after a step re-runs only the later steps, from that stage's saved outputs", async () => {
  const directory = leanFixtureMatter("lean-restore-")
  directories.push(directory)
  const first = await runLean(directory, { sessions: scriptedLeanModel().port })
  expect(first.status).toBe("completed")
  expect(readLeanDocuments(directory)?.status).toBe("completed")
  const draftedBody = readFileSync(leanFile(directory, "d1/body.docx"))
  const snapshots = path.join(leanDir(directory), "snapshots")

  const restored = restoreLeanStage(directory, "d1.changes")
  expect(restored.status).toBe("blocked")
  expect(restored.stop_reason).toBe("restored to just after step d1.changes; Continue re-runs the later steps")
  expect(Object.entries(restored.stages).map(([step, row]) => `${step}=${row?.status}`)).toEqual([
    "discrepancy=done", "intake=done", "research=done", "research-review=done", "spine=done", "d1.draft=done", "d1.review=done", "d1.review-independent=done", "d1.review-wording=done", "d1.review-client-disadvantage=done", "d1.change-plan=done", "d1.change-prepare=done", "d1.changes=done", "d1.change-review=pending", "d1.apply=pending", "finalize=pending", "checklist-check=pending", "package-review=pending", "gate=pending",
  ])
  ;["d1/body-applied.docx", "drafting-notes.md", "d1/redline.docx", "gate.json", "documents.json"].forEach(name =>
    expect(existsSync(leanFile(directory, name))).toBe(false))
  expect(readFileSync(leanFile(directory, "d1/body.docx"))).toEqual(draftedBody)
  expect(existsSync(snapshots)).toBe(true)

  // Projected-text review runs before apply; the draft, reviews and change list are reused.
  const resume = scriptedLeanModel({ [LEAN_STAGE_AGENTS.finalize]: happyLeanReplies()[LEAN_STAGE_AGENTS.finalize] })
  const resumed = await runLean(directory, { sessions: resume.port })
  expect(resumed.status).toBe("completed")
  expect(resume.agents()).toEqual(["lean-checklist-extract", LEAN_STAGE_AGENTS["change-review"], LEAN_STAGE_AGENTS.finalize, LEAN_STAGE_AGENTS["package-review"]])
  expect(resumed.stages["d1.draft"]?.finished_at).toBe(first.stages["d1.draft"]?.finished_at)
  expect(resumed.stages["d1.apply"]?.finished_at).not.toBe(first.stages["d1.apply"]?.finished_at)
  expect(readLeanState(directory)?.stages.gate?.status).toBe("done")

  const sinceRestore = readLeanEvents(directory).slice(readLeanEvents(directory).findIndex(event => event.event === "restored"))
  expect(sinceRestore.filter(event => ["skipped", "started"].includes(event.event)).map(event => `${event.event}:${event.document ?? ""}${event.stage}`)).toEqual([
    "skipped:discrepancy", "skipped:intake", "skipped:research", "skipped:research-review", "skipped:spine", "skipped:d1draft", "skipped:d1review", "skipped:d1review-independent", "skipped:d1review-wording", "skipped:d1review-client-disadvantage", "skipped:d1change-plan", "skipped:d1change-prepare", "skipped:d1changes", "started:d1change-review", "started:d1apply", "started:finalize", "started:checklist-check", "started:package-review", "started:gate",
  ])
}, 120_000)

test("restoring a step that never ran is refused without touching the run", async () => {
  const directory = leanFixtureMatter("lean-restore-")
  directories.push(directory)
  const before = readFileSync(leanFile(directory, "state.json"), "utf8")
  expect(() => restoreLeanStage(directory, "d1.review")).toThrow('No snapshot of step "d1.review" exists')
  expect(readFileSync(leanFile(directory, "state.json"), "utf8")).toBe(before)
})

test("restoring keeps a question whose answer is still queued, and the next run applies that answer", async () => {
  const directory = leanFixtureMatter("lean-restore-")
  directories.push(directory)
  const asking = sameSession((prompt, index) => JSON.stringify({
    ...JSON.parse(changes(prompt, index)), questions: [{ question: "Should late delivery carry a remedy?", kind: "judgment", context: "The terms are silent." }],
  }))
  const first = await runLean(directory, { sessions: scriptedLeanModel({ ...happyLeanReplies(), [LEAN_STAGE_AGENTS.changes]: asking }).port })
  expect(first.status).toBe("blocked")
  const question = readQuestions(directory)[0]!
  expect(question).toMatchObject({ stage: "changes", document: "d1", status: "open" })
  expect(readOpenItems(directory).some(item => !["draft", "review"].includes(item.stage))).toBe(true)
  // The lawyer answers after the run stopped; the answer waits in the queue.
  const request = enqueueLeanRequest(directory, { kind: "answer", question_id: question.id, answer: "No remedy." })

  expect(restoreLeanStage(directory, "d1.review").status).toBe("blocked")
  expect(pendingLeanRequests(directory)).toEqual([`${request.id}.json`])
  expect(readQuestions(directory).map(row => row.id)).toEqual([question.id])
  expect(readOpenItems(directory).filter(item => !["draft", "review"].includes(item.stage))).toEqual([])

  const resume = scriptedLeanModel({ ...happyLeanReplies() })
  const resumed = await runLean(directory, { sessions: resume.port })
  expect(resumed.status).toBe("completed")
  // The answer re-runs the steps that read it; the accepted draft is kept.
  expect(resume.agents()).toEqual([LEAN_STAGE_AGENTS.discrepancy, LEAN_STAGE_AGENTS.intake, LEAN_STAGE_AGENTS.research, LEAN_STAGE_AGENTS["research-review"], LEAN_STAGE_AGENTS.spine, "lean-checklist-extract", LEAN_STAGE_AGENTS["review-independent"], LEAN_STAGE_AGENTS["review-wording"], LEAN_STAGE_AGENTS["review-client-disadvantage"], LEAN_STAGE_AGENTS.review, LEAN_STAGE_AGENTS["change-plan"], LEAN_STAGE_AGENTS.changes, LEAN_STAGE_AGENTS["change-review"], LEAN_STAGE_AGENTS.finalize, LEAN_STAGE_AGENTS["package-review"]])
  expect(readLeanJson(directory, `requests/done/${request.id}.json`)).toMatchObject({ kind: "answer", outcome: { status: "answered" } })
  expect(readQuestions(directory)).toEqual([expect.objectContaining({ id: question.id, status: "answered", answer: "No remedy." })])
  const sinceRestore = readLeanEvents(directory).slice(readLeanEvents(directory).findIndex(event => event.event === "restored"))
  expect(sinceRestore.filter(event => event.event === "request_processed").map(event => event.outcome)).toEqual(["answered"])
}, 120_000)

test("restoring puts back the plan's coverage record with the plan, and restoring before the plan removes it", async () => {
  const directory = leanFixtureMatter("lean-restore-")
  directories.push(directory)
  expect((await runLean(directory, { sessions: scriptedLeanModel().port })).status).toBe("completed")
  const record = readFileSync(leanFile(directory, "coverage.json"), "utf8")
  expect(JSON.parse(record)).toMatchObject({ status: "covered" })
  // A later plan's record, as a later outline would leave it.
  writeFileSync(leanFile(directory, "coverage.json"), record.replace(/"plan": "[0-9a-f]+"/, `"plan": "${"0".repeat(64)}"`))

  restoreLeanStage(directory, "d1.changes")
  expect(readFileSync(leanFile(directory, "coverage.json"), "utf8")).toBe(record)
  const resume = scriptedLeanModel({ [LEAN_STAGE_AGENTS.finalize]: happyLeanReplies()[LEAN_STAGE_AGENTS.finalize] })
  expect((await runLean(directory, { sessions: resume.port })).status).toBe("completed")
  expect(resume.agents()).toEqual(["lean-checklist-extract", LEAN_STAGE_AGENTS["change-review"], LEAN_STAGE_AGENTS.finalize, LEAN_STAGE_AGENTS["package-review"]])

  restoreLeanStage(directory, "spine")
  expect(existsSync(leanFile(directory, "plan.json"))).toBe(false)
  expect(existsSync(leanFile(directory, "coverage.json"))).toBe(false)
}, 120_000)

test("restoring a plan written by hand removes a coverage record its snapshot does not have", async () => {
  const directory = leanFixtureMatter("lean-restore-")
  directories.push(directory)
  writeLeanPlan(directory, [{ title: "Supply note", purpose: "Record the supply terms", file: "supply-note.docx",
    parts: [{ heading: "Delivery", scope: "Delivery terms" }, { heading: "Payment", scope: "Payment terms" }] }])
  const first = scriptedLeanModel()
  expect((await runLean(directory, { sessions: first.port })).status).toBe("completed")
  expect(first.agents()).not.toContain(LEAN_STAGE_AGENTS.coverage)
  expect(existsSync(leanFile(directory, "coverage.json"))).toBe(false)
  writeFileSync(leanFile(directory, "coverage.json"), `${JSON.stringify({ schema: "dochaus.lean.coverage/1", plan: "0".repeat(64), status: "covered", acts: [], issues: [] })}\n`)

  restoreLeanStage(directory, "d1.changes")
  expect(existsSync(leanFile(directory, "coverage.json"))).toBe(false)
  const resume = scriptedLeanModel({ [LEAN_STAGE_AGENTS.finalize]: happyLeanReplies()[LEAN_STAGE_AGENTS.finalize] })
  expect((await runLean(directory, { sessions: resume.port })).status).toBe("completed")
  expect(resume.agents()).not.toContain(LEAN_STAGE_AGENTS.draft)
}, 120_000)
