import { afterEach, expect, test } from "bun:test"
import { Database } from "bun:sqlite"
import path from "node:path"
import { reconcileAnswers } from "./answer-reconcile"
import { OutputLimitError } from "./native-stage-result"
import { runLean } from "./runner"
import { readLeanJson, readLeanState, writeLeanText, writeQuestions, type Question } from "./state"
import { AGENT, cleanupMatters, matter, scriptedModel } from "./test-support"

afterEach(cleanupMatters)

const question = (id: string, answer: string, context = "Delivery terms for the supply note."): Question => ({
  id, question: `What should the note say about ${id}?`, context, stage: "intake", kind: "instruction", status: "answered", answer,
})
const proposed = (ids: string[]) => JSON.stringify({ conflicts: [{ questions: ids, conflict: "The proposed directions cannot both be followed." }] })
const review = (rows: Question[], status: "conflict" | "compatible", extra: Record<string, unknown> = {}) => JSON.stringify({
  status, reason: status === "conflict" ? "The same operative point requires two incompatible terms." : "The future condition has not supplied a different fact; these directions can both be followed.",
  directions: rows.map(row => ({ question_id: row.id, quote: row.answer })),
  applicability: "The requested drafting and the actual scope and conditions of these directions were compared.", basis: [], ...extra,
})

function fixture(rows: Question[], verification: (prompt: string, index: number) => string) {
  const directory = matter()
  writeQuestions(directory, rows)
  let reviews = 0
  const reviewSessions = new Set<string>()
  const model = scriptedModel({ [AGENT.reconcile]: (prompt, _index, sessionID) => {
    if (prompt.includes("Verify ONE proposed answer conflict")) reviewSessions.add(sessionID!)
    return reviewSessions.has(sessionID!) ? verification(prompt, reviews++) : proposed(rows.slice(0, 2).map(row => row.id))
  } })
  return { directory, model, run: () => reconcileAnswers(directory, readLeanState(directory)!, { sessions: model.port }) }
}

test("an independently rejected hypothetical conflict never becomes a lawyer question", async () => {
  const rows = [question("q-particulars", "Use the particulars provisionally; if a later records check differs, refer the difference for a decision."),
    question("q-source", "Follow the supplied note as instructions for this work.")]
  const { directory, model, run } = fixture(rows, () => review(rows, "compatible"))
  expect(await run()).toEqual([])
  expect(model.prompts.filter(row => row.text.includes("Verify ONE proposed answer conflict"))).toHaveLength(1)
  expect(readLeanJson(directory, "answer-reconcile.json")).toMatchObject({ conflicts: [] })
})

test.each([
  ["an actual operative inconsistency", "Require payment on delivery.", "Require payment thirty days after delivery."],
  ["incompatible terms in the same future conditional branch", "If delivery is late, require a price reduction.", "If delivery is late, exclude every price reduction."],
])("a verified conflict remains a lawyer decision: %s", async (_name, first, second) => {
  const rows = [question("q-first", first), question("q-second", second)]
  const { run } = fixture(rows, () => review(rows, "conflict"))
  expect(await run()).toEqual([expect.objectContaining({ kind: "judgment", status: "open" })])
})

test("the focused review receives only the candidate answers and preserves source-authority context", async () => {
  const rows = [question("q-first", "Keep the details provisional.", "Whether to follow the uploaded instructions, not whether every fact is verified."),
    question("q-second", "Follow the supplied note as instructions."), question("q-unrelated", "An unrelated commercial preference.")]
  const { run } = fixture(rows, prompt => {
    expect(prompt).toContain(rows[0]!.context)
    expect(prompt).not.toContain(rows[2]!.answer!)
    expect(prompt).toContain("source authority")
    expect(prompt).toContain("conditional branch")
    return review(rows.slice(0, 2), "compatible")
  })
  expect(await run()).toEqual([])
})

test("different document and paragraph scopes remain visible to the conflict reviewer", async () => {
  const rows: Question[] = [{ ...question("q-one", "Use a fixed period in this document."), document: "d1", pid: "p123456789abc" },
    { ...question("q-two", "Use a variable period in this document."), document: "d2", pid: "pabcdef123456" }]
  const { run } = fixture(rows, prompt => {
    for (const row of rows) {
      expect(prompt).toContain(`"document":"${row.document}"`)
      expect(prompt).toContain(`"paragraph":"${row.pid}"`)
    }
    return review(rows, "compatible")
  })
  expect(await run()).toEqual([])
})

test("the governing request and later updates reach the focused reviewer", async () => {
  const rows = [question("q-one", "Use a fixed period."), question("q-two", "Leave the period open.")]
  const request = "Draft only the supplier's internal working note."
  const update = "The later direction to leave the period open replaces the earlier fixed-period direction."
  const { directory, run } = fixture(rows, prompt => {
    expect(prompt).toContain(request)
    expect(prompt).toContain(update)
    expect(prompt).toContain("<request>")
    expect(prompt).toContain("<user_updates>")
    return review(rows, "compatible")
  })
  writeLeanText(directory, "request.md", request)
  writeLeanText(directory, "updated-notes.md", update)
  expect(await run()).toEqual([])
})

test("an invented answer quotation is retried in the same focused session before any conflict is accepted", async () => {
  const rows = [question("q-first", "Use a fixed delivery date."), question("q-second", "Leave the delivery date open.")]
  const { run, model } = fixture(rows, (_prompt, index) => review(rows, "conflict", index === 0
    ? { directions: [{ question_id: rows[0]!.id, quote: "A completed check changed the delivery date." }, { question_id: rows[1]!.id, quote: rows[1]!.answer }] } : {}))
  expect(await run()).toHaveLength(1)
  const calls = model.prompts.filter(row => row.text.includes("Verify ONE proposed answer conflict") || row.text.includes("not in that answer"))
  expect(calls).toHaveLength(2)
  expect(calls[0]!.sessionID).toBe(calls[1]!.sessionID)
})

test.each(["unknown reference", "invalid source quote", "malformed output", "uncertain applicability"])("failed conflict verification is not cached as compatibility: %s", async failure => {
  const rows = [question("q-first", "Use a fixed delivery date."), question("q-second", "Leave the delivery date open.")]
  const { directory, run } = fixture(rows, () => failure === "malformed output" ? "unreadable"
    : review(rows, "conflict", failure === "uncertain applicability" ? { status: "unresolved" } : failure === "unknown reference"
      ? { directions: [{ question_id: "q-absent", quote: rows[0]!.answer }, { question_id: rows[1]!.id, quote: rows[1]!.answer }] }
      : { basis: [{ source_id: "supply-terms.docx", quote: "A completed check changed the delivery date." }] }))
  await expect(run()).rejects.toThrow("The proposed answer conflict could not be verified")
  expect(readLeanJson(directory, "answer-reconcile.json")).toMatchObject({ verification_failed: true })
  expect(readLeanJson(directory, "answer-reconcile.json")).not.toHaveProperty("conflicts")
})

test("verification failure stops the product walk before drafting and records no compatibility result", async () => {
  const rows = [question("q-first", "Keep the date fixed."), question("q-second", "Leave the date open.")]
  const { directory, model } = fixture(rows, () => review(rows, "conflict", { status: "unresolved" }))
  expect(await runLean(directory, { sessions: model.port })).toMatchObject({ status: "failed", stop_reason: expect.stringContaining("could not be verified") })
  expect(model.count(AGENT.draft)).toBe(0)
  expect(readLeanJson(directory, "answer-reconcile.json")).toMatchObject({ verification_failed: true })
  expect(readLeanJson(directory, "answer-reconcile.json")).not.toHaveProperty("conflicts")
})

test("an unchanged Continue does not reset exhausted verification; changed answers reopen it", async () => {
  const rows = [question("q-first", "Keep the delivery date fixed."), question("q-second", "Leave the delivery date open.")]
  let available = false
  const { directory, model, run } = fixture(rows, () => available ? review(rows, "compatible") : "unreadable")
  await expect(run()).rejects.toThrow("could not be verified")
  const calls = model.count(AGENT.reconcile)
  await expect(run()).rejects.toThrow("could not be verified")
  expect(model.count(AGENT.reconcile)).toBe(calls)
  available = true
  rows[1]!.answer = "Keep the same delivery date fixed."
  writeQuestions(directory, rows)
  expect(await run()).toEqual([])
  expect(model.count(AGENT.reconcile)).toBe(calls + 2)
  expect(readLeanJson(directory, "answer-reconcile.json")).toMatchObject({ conflicts: [] })
  expect(readLeanJson(directory, "answer-reconcile.json")).not.toHaveProperty("verification_failed")
})

test("a verifier output-limit stop preserves exhaustion on an unchanged Continue", async () => {
  const rows = [question("q-first", "Fix the period."), question("q-second", "Leave the period open.")]
  const { directory, model, run } = fixture(rows, () => { throw new OutputLimitError("The review reached its output limit.") })
  await expect(run()).rejects.toThrow("could not be verified")
  const calls = model.count(AGENT.reconcile)
  await expect(run()).rejects.toThrow("could not be verified")
  expect(model.count(AGENT.reconcile)).toBe(calls)
  expect(readLeanJson(directory, "answer-reconcile.json")).toMatchObject({ verification_failed: true })
  expect(readLeanJson(directory, "answer-reconcile.json")).not.toHaveProperty("conflicts")
})

test("an unchanged verified result is reused, but changed source evidence reopens the comparison", async () => {
  const rows = [question("q-first", "Keep the stated period provisional pending the source check."), question("q-second", "Follow the supplied source.")]
  const { directory, model, run } = fixture(rows, (_prompt, index) => review(rows, index ? "conflict" : "compatible"))
  expect(await run()).toEqual([])
  expect(await run()).toEqual([])
  expect(model.count(AGENT.reconcile)).toBe(2)
  const db = new Database(path.join(directory, ".dochaus", "legal.db"))
  try { db.run("UPDATE chunks SET text = ?", ["The revised supplier terms require an incompatible delivery period."]) } finally { db.close() }
  expect(await run()).toHaveLength(1)
  expect(model.count(AGENT.reconcile)).toBe(4)
})

test("separate candidates are reviewed separately and a concurrent runtime stop is not cached as an exhausted check", async () => {
  const rows = [question("q-date-a", "Fix the date."), question("q-date-b", "Leave the date open."),
    question("q-price-a", "Fix the price."), question("q-price-b", "Leave the price open.")]
  const directory = matter()
  writeQuestions(directory, rows)
  let reviewed = 0
  const model = scriptedModel({ [AGENT.reconcile]: prompt => {
    if (!prompt.includes("Verify ONE proposed answer conflict")) return JSON.stringify({ conflicts: [
      { questions: rows.slice(0, 2).map(row => row.id), conflict: "The delivery-date directions differ." },
      { questions: rows.slice(2).map(row => row.id), conflict: "The price directions differ." },
    ] })
    reviewed++
    if (prompt.includes('"question_id":"q-date-a"')) {
      expect(prompt).not.toContain(rows[2]!.answer!)
      return review(rows.slice(0, 2), "conflict", { status: "unresolved" })
    }
    expect(prompt).not.toContain(rows[0]!.answer!)
    throw new Error("The native review session stopped.")
  } })
  await expect(reconcileAnswers(directory, readLeanState(directory)!, { sessions: model.port })).rejects.toThrow("native review session stopped")
  expect(reviewed).toBe(2)
  expect(readLeanJson(directory, "answer-reconcile.json")).toBeUndefined()
})
