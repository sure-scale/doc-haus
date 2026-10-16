import { afterEach, expect, test } from "bun:test"
import { Database } from "bun:sqlite"
import { readFileSync } from "node:fs"
import path from "node:path"
import { readParagraphs } from "./checks"
import { openItem } from "./quotes"
import { documentRepairs, readRepairRound, startRepairRound } from "./repairs"
import { runLean, stageInputSha } from "./runner"
import { leanFile, readLeanState, writeLeanText, writeQuestions } from "./state"
import { AGENT, APPLIED, DRAFT, LETTER, changesFirst, cleanupMatters, drafter, happyPath, matter, reviewFirst, scriptedModel } from "./test-support"

afterEach(cleanupMatters)

function finding(directory: string, document: "d1" | "d2", detail: string, target: "body" | "research" = "body") {
  return openItem("change-review", "omission", detail, { document, repair_target: target, pid: readParagraphs(directory, document)[0]!.pid })
}

async function completedRepair() {
  const directory = matter()
  const model = scriptedModel({ ...happyPath, [AGENT.draft]: drafter(DRAFT, LETTER), [AGENT.review]: reviewFirst,
    [AGENT.changes]: (prompt, index) => {
      const value = JSON.parse(changesFirst(prompt, index))
      if (prompt.includes("Add the supplied payment timing to delivery.")) value.amendments[0].updated = `${APPLIED}\n\nPayment is due on delivery.`
      return JSON.stringify(value)
    } })
  expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
  // Legacy repair records have no per-document context; round3 must migrate
  // without changing the input key accepted by its existing reader.
  writeLeanText(directory, "repairs.json", JSON.stringify({ round: 3, requested_at: "2026-01-01T00:00:00Z", substantive_attempts: 2,
    items: [finding(directory, "d1", "Reassess the supplied delivery timing.")] }))
  expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
  return { directory, model }
}

test("a memo-only round preserves completed upstream stages and draft, but repairs that memo", async () => {
  const { directory, model } = await completedRepair()
  const before = readLeanState(directory)!
  const outputs = Object.fromEntries(Object.entries(before.stages).filter(([step]) => step.startsWith("d1.") || step === "d2.draft")
    .flatMap(([, stage]) => (stage?.outputs ?? []).map(file => [file, readFileSync(leanFile(directory, file))])))
  const memoPlan = stageInputSha(directory, "d2.change-plan")
  const memoDraft = stageInputSha(directory, "d2.draft")
  startRepairRound(directory, [finding(directory, "d2", "Reassess the covering explanation.")], true)
  expect(readRepairRound(directory)).toMatchObject({ round: 4, substantive_attempts: 3, items: [expect.objectContaining({ ref: expect.objectContaining({ document: "d2" }) })] })
  expect(documentRepairs(directory, "d1")).toMatchObject({ round: 3 })
  for (const [step, row] of Object.entries(before.stages)) {
    if (step.startsWith("d1.")) expect(stageInputSha(directory, step as "d1.draft")).toBe(row!.input_sha!)
  }
  expect(stageInputSha(directory, "d2.draft")).toBe(memoDraft)
  expect(stageInputSha(directory, "d2.change-plan")).not.toBe(memoPlan)
  const calls = model.prompts.length
  expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
  const after = readLeanState(directory)!
  for (const [step, row] of Object.entries(before.stages)) {
    if (step.startsWith("d1.") || step === "d2.draft") expect(after.stages[step as "d1.draft"]).toEqual(row)
  }
  for (const [file, bytes] of Object.entries(outputs)) expect(readFileSync(leanFile(directory, file))).toEqual(bytes)
  expect(model.prompts.slice(calls).some(row => row.agent === AGENT.draft)).toBe(false)
  expect(after.stages["d2.change-plan"]?.input_sha).not.toBe(before.stages["d2.change-plan"]?.input_sha)

  // A real subsequent upstream correction replaces, rather than unions, its
  // retired input and propagates its changed final text to the memo draft.
  const upstreamKey = stageInputSha(directory, "d1.change-plan")
  startRepairRound(directory, [finding(directory, "d1", "Add the supplied payment timing to delivery.")], true)
  expect(documentRepairs(directory, "d1")?.items.map(row => row.detail)).toEqual(["Add the supplied payment timing to delivery."])
  expect(readRepairRound(directory)?.document_contexts?.d1).toBeUndefined()
  expect(stageInputSha(directory, "d1.change-plan")).not.toBe(upstreamKey)
  expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
  expect(stageInputSha(directory, "d2.draft")).not.toBe(memoDraft)
}, 120_000)

test("retained repair conditions expire on changed source or answered instructions", async () => {
  for (const change of ["source", "answer", "draft"]) {
    const { directory } = await completedRepair()
    startRepairRound(directory, [finding(directory, "d2", "Reassess advice.")])
    expect(documentRepairs(directory, "d1")?.round).toBe(3)
    if (change === "source") {
      const db = new Database(path.join(directory, ".dochaus", "legal.db"))
      try { db.run("UPDATE chunks SET text = text || ' Revised timing applies.'") } finally { db.close() }
    }
    if (change === "answer") writeQuestions(directory, [{ id: "timing-answer", question: "Which timing applies?", context: "Timing instruction", status: "answered", answer: "Use the newly supplied terms.", stage: "change-plan", document: "d1" }])
    if (change === "draft") writeLeanText(directory, "d1/paragraphs.json", JSON.stringify(readParagraphs(directory, "d1").map((row, index) => ({ ...row, pid: `p${String(index).padStart(12, "0")}` }))))
    expect(documentRepairs(directory, "d1")).toBeUndefined()
    expect(stageInputSha(directory, "d1.change-plan")).not.toBe(readLeanState(directory)!.stages["d1.change-plan"]!.input_sha!)
  }
}, 120_000)

test("addressing a document retires its old body repair even when the new repair is research-only", async () => {
  const { directory } = await completedRepair()
  startRepairRound(directory, [finding(directory, "d2", "Reassess advice.")])
  expect(documentRepairs(directory, "d1")?.round).toBe(3)
  startRepairRound(directory, [finding(directory, "d1", "Research the new proposition.", "research")])
  expect(documentRepairs(directory, "d1")).toBeUndefined()
  expect(readRepairRound(directory)?.document_contexts?.d1).toBeUndefined()
}, 120_000)

test("an already stale former admission is not retained under a new basis", async () => {
  const { directory } = await completedRepair()
  writeLeanText(directory, "updated-notes.md", "A new instruction changes the earlier assessment.")
  startRepairRound(directory, [finding(directory, "d2", "Reassess advice.")])
  expect(documentRepairs(directory, "d1")).toBeUndefined()
  expect(readRepairRound(directory)?.document_contexts?.d1).toBeUndefined()
}, 120_000)

test("run-wide notes repair does not retire the first document's completed body repair", async () => {
  const { directory } = await completedRepair()
  const key = stageInputSha(directory, "d1.change-plan")
  startRepairRound(directory, [openItem("package-review", "omission", "Clarify the run-wide drafting notes.", { repair_target: "memo" })])
  expect(documentRepairs(directory, "d1")?.round).toBe(3)
  expect(stageInputSha(directory, "d1.change-plan")).toBe(key)
}, 120_000)
