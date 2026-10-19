import { afterEach, expect, test } from "bun:test"
import { Database } from "bun:sqlite"
import path from "node:path"
import { acceptEvidenceReply, checkEvidence, evidencePrompt } from "./evidence"
import { enqueueLeanRequest, drainLeanRequests } from "./requests"
import { leanSteps } from "./plan"
import { stageInputSha } from "./runner"
import { answerQuestion, writeQuestions, writeLeanText } from "./state"
import { cleanupMatters, matter, SOURCE, SOURCE_TEXT } from "./test-support"

afterEach(cleanupMatters)

test("discrepancy evidence must match canonical text and uncertainty stays explicit", () => {
  const directory = matter()
  const entry = { detail: "Delivery term appears in the source", evidence: [{ source_id: SOURCE, quote: SOURCE_TEXT }], limitations: [], material: false, status: "benign", resolution: "Delivery and payment are distinct events." }
  expect(checkEvidence(directory, "discrepancy", { entries: [entry], limitations: [] })).toEqual([])
  expect(checkEvidence(directory, "discrepancy", { entries: [{ ...entry, evidence: [{ source_id: SOURCE, quote: "Invented delivery period" }] }], limitations: [] }).length).toBeGreaterThan(0)
  expect(checkEvidence(directory, "discrepancy", { entries: [{ ...entry, evidence: [], limitations: [] }], limitations: [] }).length).toBeGreaterThan(0)
  expect(checkEvidence(directory, "discrepancy", { entries: [{ ...entry, evidence: [], limitations: ["Original unavailable"], status: "unresolved", resolution: "" }], limitations: [] })).toEqual([])
})

test("later answers and same-name source revisions invalidate dependent work", () => {
  const directory = matter()
  const before = stageInputSha(directory, "discrepancy")
  const db = new Database(path.join(directory, ".dochaus/legal.db"))
  db.run("UPDATE chunks SET text = ?", [SOURCE_TEXT + " Revised instruction."])
  db.close()
  expect(stageInputSha(directory, "discrepancy")).not.toBe(before)
  const review = stageInputSha(directory, "d1.review")
  writeQuestions(directory, [{ id: "q-choice", stage: "intake", question: "Which delivery term applies?", context: "Versions conflict", status: "open" }])
  answerQuestion(directory, "q-choice", "Use the later expressly agreed term.")
  expect(stageInputSha(directory, "d1.review")).not.toBe(review)
  expect(evidencePrompt(directory, "discrepancy")).toContain("Use the later expressly agreed term.")
  const spine = stageInputSha(directory, "spine")
  writeLeanText(directory, "updated-notes.md", "Use the earlier term instead.")
  expect(stageInputSha(directory, "spine")).not.toBe(spine)
})

test("spine is source-linked and followed by four independent reviews", () => {
  const directory = matter()
  expect(acceptEvidenceReply(directory, "spine", JSON.stringify({ entries: [{ kind: "requirement", detail: "State payment timing", evidence: [{ source_id: SOURCE, quote: "Payment is due on delivery." }], limitations: [] }], limitations: [] })).issues).toEqual([])
  expect(leanSteps(directory)).toEqual(["discrepancy", "intake", "research", "research-review", "spine", "d1.draft", "d1.review", "d1.review-independent", "d1.review-wording", "d1.review-client-disadvantage", "d1.change-plan", "d1.change-prepare", "d1.changes", "d1.change-review", "d1.apply", "finalize", "checklist-check", "package-review", "gate"])
  expect(evidencePrompt(directory, "spine")).toContain("future stages must retrieve underlying text")
})


test("user instruction evidence uses the canonical request identity and rejects handoff aliases", () => {
  const directory = matter()
  const reply = (source_id: string, quote: string) => ({ entries: [{ kind: "decision", detail: "Follow the supplied instruction", evidence: [{ source_id, quote }], limitations: [] }], limitations: [] })
  enqueueLeanRequest(directory, { kind: "update", text: "Use the revised delivery provision." })
  drainLeanRequests(directory)
  writeQuestions(directory, [{ id: "q-evidence", stage: "intake", question: "Which provision?", context: "Material choice", status: "open" }])
  answerQuestion(directory, "q-evidence", "Use the expressly agreed later provision.")
  expect(checkEvidence(directory, "spine", reply("request", "Use the revised delivery provision."))).toEqual([])
  expect(checkEvidence(directory, "spine", reply("request", "Use the expressly agreed later provision."))).toEqual([])
  for (const alias of ["Request text", "Human answer", "Intake", "Verified research", "update:unavailable"]) {
    const issues = checkEvidence(directory, "spine", reply(alias, "Use the revised delivery provision."))
    expect(issues[0]).toContain("Unknown evidence source_id")
    expect(issues[0]).toContain(SOURCE)
    expect(issues[0]).toContain('source_id "request"')
    expect(issues[0]).toContain("derived records")
  }
  expect(checkEvidence(directory, "spine", reply(SOURCE, "Unsupported words"))[0]).toContain("not found in that canonical source")
  expect(evidencePrompt(directory, "spine")).toContain('source_id "request"')
})

test("a rejected quotation identifies the exact evidence to correct without accepting display labels", () => {
  const directory = matter()
  const answer = "Use the expressly agreed later provision."
  writeQuestions(directory, [{ id: "q-feedback", stage: "intake", question: "Which provision?", context: "Material choice", status: "open" }])
  answerQuestion(directory, "q-feedback", answer)
  const reply = (quote: string) => ({ entries: [{ kind: "decision", detail: "Follow the answer", evidence: [{ source_id: "request", quote }], limitations: [] }], limitations: [] })
  expect(checkEvidence(directory, "spine", reply(answer))).toEqual([])
  const rejected = `Answer: ${answer}`
  const issues = checkEvidence(directory, "spine", reply(rejected))
  expect(issues).toHaveLength(1)
  expect(issues[0]).toContain(JSON.stringify(rejected))
  expect(issues[0]).toContain('in "request"')
  expect(issues[0]).toContain("not found in that canonical source")
})
