import { afterEach, expect, test } from "bun:test"
import { Database } from "bun:sqlite"
import path from "node:path"
import { acceptEvidenceReply, checkEvidence, evidencePrompt } from "./evidence"
import { finalGate } from "./gate"
import { enqueueLeanRequest, drainLeanRequests } from "./requests"
import { leanSteps } from "./plan"
import { stagePrompt } from "./prompts"
import { documentRepairs, startRepairRound } from "./repairs"
import { runLean, stageInputSha } from "./runner"
import { answerQuestion, readLeanText, readOpenItems, writeQuestions, writeLeanText } from "./state"
import { AGENT, cleanupMatters, happyPath, matter, scriptedModel, SOURCE, SOURCE_TEXT } from "./test-support"

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
  const draft = stageInputSha(directory, "d1.draft")
  writeQuestions(directory, [{ id: "q-choice", stage: "intake", question: "Which delivery term applies?", context: "Versions conflict", status: "open" }])
  answerQuestion(directory, "q-choice", "Use the later expressly agreed term.")
  expect(stageInputSha(directory, "d1.draft")).not.toBe(draft)
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

// Spine entries for the disclosure tests: one the matter supports and
// substantive ones that rest only on a recorded limitation.
const SUPPORTED = { kind: "requirement", detail: "State that payment is due on delivery.", evidence: [{ source_id: SOURCE, quote: "Payment is due on delivery." }], limitations: ["Currency of the terms is not confirmed."] }
const UNSUPPORTED = { kind: "decision", detail: "Fix delivery to a calendar date rather than a period after the order.", evidence: [], limitations: ["The supply terms give a period, not a date, and no instruction chooses a date."] }

test("only a substantive spine entry without matter support becomes an open item, worded for the lawyer", () => {
  const directory = matter()
  const entry = (kind: string, detail: string) => ({ kind, detail, evidence: [], limitations: [`No supplied source settles: ${detail}`] })
  const substantive = [entry("decision", "Use a fixed delivery date."), entry("requirement", "Include a late-delivery remedy."),
    entry("recommendation", "Offer a shorter payment period."), entry("open-item", "Whether title passes on delivery.")]
  const accepted = acceptEvidenceReply(directory, "spine", JSON.stringify({ entries: [SUPPORTED, entry("fact", "The goods are standard stock."), ...substantive], limitations: ["Authority on remedies was not researched."] }))
  expect(accepted.issues).toEqual([])
  expect(accepted.openItems).toHaveLength(substantive.length)
  for (const [index, item] of accepted.openItems.entries()) {
    expect(item).toMatchObject({ stage: "spine", kind: "unsupported_choice", status: "open", ref: { repair_target: "memo" } })
    expect(item.detail).toContain(substantive[index]!.detail)
    expect(item.detail).toContain(substantive[index]!.limitations[0]!)
    expect(item.detail).not.toMatch(/spine|stage|evidence|entr(y|ies)|limitations/i)
  }
  // Supported entries and evidence-less discrepancy findings give none.
  expect(acceptEvidenceReply(directory, "spine", JSON.stringify({ entries: [SUPPORTED], limitations: [] })).openItems).toEqual([])
  expect(acceptEvidenceReply(directory, "discrepancy", JSON.stringify({ entries: [{ detail: "Two delivery periods", evidence: [], limitations: ["Original unavailable"], material: true, status: "unresolved", resolution: "" }], limitations: [] })).openItems).toEqual([])
})

test("an evidence-less spine decision is disclosed in the notes, held by the final gate until they name it, and repaired only in the notes", async () => {
  const directory = matter()
  const model = scriptedModel({ ...happyPath, [AGENT.spine]: () => JSON.stringify({ entries: [UNSUPPORTED, SUPPORTED], limitations: [] }) })
  expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
  const items = readOpenItems(directory).filter(item => item.stage === "spine")
  expect(items.map(row => [row.kind, row.ref])).toEqual([["unsupported_choice", { repair_target: "memo" }]])
  const item = items[0]!
  expect(item.detail).toContain(UNSUPPORTED.detail)
  expect(item.detail).toContain(UNSUPPORTED.limitations[0]!)
  // The notes are told of it, and the final gate settles it against them.
  expect(model.prompts.find(row => row.agent === AGENT.finalize)?.text).toContain(`- ${item.id}: ${item.detail}`)
  const notes = readLeanText(directory, "drafting-notes.md")!
  expect(notes).toContain(`[ref: ${item.id}]`)
  expect(finalGate(directory)).toMatchObject({ status: "pass", delivery: "review_draft", disclosed: expect.arrayContaining([expect.objectContaining({ id: item.id, stage: "spine", status: "disclosed" })]) })
  writeLeanText(directory, "drafting-notes.md", notes.replace(`[ref: ${item.id}]`, ""))
  const held = finalGate(directory)
  expect(held.status).toBe("blocked")
  expect(held.blocking.map(row => [row.id, row.stage, row.status])).toEqual([[item.id, "spine", "open"]])
  // A repair round for it reaches the notes only: no document body is repaired.
  startRepairRound(directory, held.blocking)
  expect(documentRepairs(directory, "d1")).toBeUndefined()
  expect(/<required_repairs>\n([\s\S]*?)\n<\/required_repairs>/.exec(stagePrompt(directory, "finalize"))?.[1]).toBe(`- ${item.detail}`)
  const resumed = scriptedModel(happyPath)
  expect((await runLean(directory, { sessions: resumed.port })).status).toBe("completed")
  expect([AGENT.draft, AGENT.review, AGENT.changes].map(resumed.count)).toEqual([0, 0, 0])
  expect(resumed.count(AGENT.finalize)).toBe(1)
  expect(readLeanText(directory, "drafting-notes.md")).toContain(`[ref: ${item.id}]`)
  expect(finalGate(directory)).toMatchObject({ status: "pass", disclosed: expect.arrayContaining([expect.objectContaining({ id: item.id, status: "disclosed" })]) })
}, 120_000)
