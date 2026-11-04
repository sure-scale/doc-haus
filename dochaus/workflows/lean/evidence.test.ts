import { afterEach, expect, test } from "bun:test"
import { Database } from "bun:sqlite"
import { writeFileSync } from "node:fs"
import path from "node:path"
import { researchBasis } from "./stages"
import { researchPrompt } from "./research"
import { acceptEvidenceReply, checkEvidence, evidencePrompt } from "./evidence"
import { enqueueLeanRequest, drainLeanRequests } from "./requests"
import { leanSteps } from "./plan"
import { stageInputSha } from "./runner"
import { answerQuestion, writeQuestions, writeLeanText } from "./state"
import { writeInstructionsFile } from "./instructions"
import { cleanupMatters, matter, SOURCE, SOURCE_TEXT } from "./test-support"

afterEach(cleanupMatters)

test("discrepancy evidence must match canonical text and uncertainty stays explicit", () => {
  const directory = matter()
  const entry = { detail: "Delivery term appears in the source", evidence: [{ source_id: SOURCE, quote: SOURCE_TEXT }], limitations: [], material: false, status: "benign", resolution: "Delivery and payment are distinct events." }
  expect(checkEvidence(directory, "discrepancy", { entries: [entry], gaps: [], limitations: [] })).toEqual([])
  expect(checkEvidence(directory, "discrepancy", { entries: [{ ...entry, evidence: [{ source_id: SOURCE, quote: "Invented delivery period" }] }], gaps: [], limitations: [] }).length).toBeGreaterThan(0)
  expect(checkEvidence(directory, "discrepancy", { entries: [{ ...entry, evidence: [], limitations: [] }], gaps: [], limitations: [] }).length).toBeGreaterThan(0)
  expect(checkEvidence(directory, "discrepancy", { entries: [{ ...entry, evidence: [], limitations: ["Original unavailable"], status: "unresolved", resolution: "" }], gaps: [], limitations: [] })).toEqual([])
  expect(checkEvidence(directory, "discrepancy", { entries: [entry, { ...entry, resolution: " " }], gaps: [], limitations: [] })).toEqual([
    "entries[1] has status benign with an empty resolution. Explain the evidence that settles it or the intentional distinction; leave resolution empty only for an unresolved or verify-source entry. Unresolved findings must never be silently corrected.",
  ])
})

test("a quotation of the run's own record of a settled question points to the lawyer's words behind it", () => {
  const directory = matter()
  const decided_by = { text: "Keep it in one document.", source_id: SOURCE, quote: SOURCE_TEXT }
  writeQuestions(directory, [{ id: "q-one", stage: "intake", question: "One document or two?", context: "", status: "answered", answer: "Keep one combined document", answer_kind: "instruction", decided_by }])
  const entry = { detail: "One document is wanted", evidence: [{ source_id: "request", quote: "Keep one combined document" }], limitations: [], material: false, status: "resolved", resolution: "The instruction settles it." }
  const [issue] = checkEvidence(directory, "discrepancy", { entries: [entry], gaps: [], limitations: [] })
  expect(issue).toContain(`Those words are the run's own record of question q-one, which an instruction settled; they are not the lawyer's words. The lawyer's words behind it are "${SOURCE_TEXT}" in "${SOURCE}".`)
  expect(evidencePrompt(directory, "discrepancy")).toContain(`from their words "${SOURCE_TEXT}" in ${SOURCE}`)
})

// A gap is a point the request or instructions need that no source settles.
// Its requirement quotes the words that need it, verified like evidence.
const GAP = { detail: "The supply terms state no governing law; the note needs one.", requirement: { source_id: "request", quote: "Draft a short supply note" },
  evidence: [{ source_id: SOURCE, quote: "Payment is due on delivery." }], limitations: [], material: true, lawyer: true }

test("the discrepancy reply lists gaps, each resting on verified words that need it", () => {
  const directory = matter()
  const reply = (...gaps: object[]) => ({ entries: [], gaps, limitations: [] })
  expect(checkEvidence(directory, "discrepancy", reply(GAP))).toEqual([])
  expect(checkEvidence(directory, "discrepancy", { entries: [], limitations: [] })).toEqual(["Return JSON with entries, gaps and limitations arrays."])
  expect(checkEvidence(directory, "spine", { entries: [], limitations: [] })).toEqual([])
  expect(checkEvidence(directory, "discrepancy", reply({ ...GAP, requirement: undefined }))).toEqual(["gaps[0] needs detail, requirement {source_id,quote}, evidence [{source_id,quote}] and limitations arrays."])
  const [invented] = checkEvidence(directory, "discrepancy", reply({ ...GAP, requirement: { source_id: "request", quote: "Draft a long supply agreement" } }))
  expect(invented).toStartWith('gaps[0].requirement: The supporting quotation "Draft a long supply agreement" in "request" was not found')
  const [unsupported] = checkEvidence(directory, "discrepancy", reply(GAP, { ...GAP, evidence: [{ source_id: SOURCE, quote: "Invented delivery period" }] }))
  expect(unsupported).toStartWith('gaps[1].evidence: The supporting quotation "Invented delivery period"')
  expect(checkEvidence(directory, "discrepancy", reply({ ...GAP, material: "yes", lawyer: undefined }))).toEqual([
    "gaps[0] needs material: true or false.",
    "gaps[0] needs lawyer: true when only the lawyer or client can supply it, else false.",
  ])
})

test("a gap's requirement comes from the lawyer's words or a document the lawyer said to follow", () => {
  const directory = matter()
  const fromSource = { ...GAP, requirement: { source_id: SOURCE, quote: "within thirty days of the order date" } }
  expect(checkEvidence(directory, "discrepancy", { entries: [], gaps: [fromSource], limitations: [] })).toEqual([
    `gaps[0].requirement quotes "${SOURCE}", which gives no instructions for this work. Quote the words that need the point from one of ["request"]; a point only another source raises is an entry or evidence, not a gap.`,
  ])
  writeLeanText(directory, "instructions.json", JSON.stringify({ items: [{ id: "i1", text: "Deliver within thirty days.", source_id: SOURCE, quote: "within thirty days of the order date" }],
    documents: [{ source_id: SOURCE, from: "the client", quote: "within thirty days of the order date", designated_by: "from the supply terms" }] }))
  writeInstructionsFile(directory)
  expect(checkEvidence(directory, "discrepancy", { entries: [], gaps: [fromSource], limitations: [] })).toEqual([])
})

test("the discrepancy job reads the instruction list and keys it; intake and the spine route its gaps", () => {
  const directory = matter()
  const before = stageInputSha(directory, "discrepancy")
  writeLeanText(directory, "instructions.json", JSON.stringify({ items: [{ id: "i1", text: "Draft a short supply note.", source_id: "request", quote: "Draft a short supply note" }], documents: [] }))
  writeInstructionsFile(directory)
  expect(stageInputSha(directory, "discrepancy")).not.toBe(before)
  const prompt = evidencePrompt(directory, "discrepancy")
  expect(prompt).toContain("The lawyer's instructions for this work, as a concise list:\n# The lawyer's instructions")
  expect(prompt).toContain("Then list gaps: each point the request or these instructions need settled that no source, answer or instruction settles.")
  expect(prompt).not.toContain("record the difference as a limitation")
  expect(evidencePrompt(directory, "spine")).toContain("record the difference as a limitation")
  expect(acceptEvidenceReply(directory, "discrepancy", JSON.stringify({ entries: [], gaps: [GAP], limitations: [] })).issues).toEqual([])
  const intake = researchPrompt(directory, "intake", researchBasis(directory))
  expect(intake).toContain(GAP.detail)
  expect(intake).toContain("ask a material gap marked lawyer unless an answer or instruction now settles it; the others stay disclosed placeholders or open items.")
  const spine = evidencePrompt(directory, "spine")
  expect(spine).toContain(GAP.detail)
  expect(spine).toContain("Record each discrepancy gap that no answer or approved research settles as an open-item entry.")
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
  answerQuestion(directory, "q-choice", { text: "Use the later expressly agreed term." })
  expect(stageInputSha(directory, "d1.review")).not.toBe(review)
  expect(evidencePrompt(directory, "discrepancy")).toContain("Use the later expressly agreed term.")
  const spine = stageInputSha(directory, "spine")
  writeLeanText(directory, "updated-notes.md", "Use the earlier term instead.")
  expect(stageInputSha(directory, "spine")).not.toBe(spine)
})

test("spine is source-linked and followed by four independent reviews", () => {
  const directory = matter()
  expect(acceptEvidenceReply(directory, "spine", JSON.stringify({ entries: [{ kind: "requirement", detail: "State payment timing", evidence: [{ source_id: SOURCE, quote: "Payment is due on delivery." }], limitations: [] }], limitations: [] })).issues).toEqual([])
  expect(leanSteps(directory)).toEqual(["instructions", "discrepancy", "intake", "research", "research-review", "acts", "spine", "d1.draft", "d1.review", "d1.review-independent", "d1.review-wording", "d1.review-client-disadvantage", "d1.change-plan", "d1.change-prepare", "d1.changes", "d1.change-review", "d1.apply", "finalize", "checklist-check", "package-review", "gate"])
  expect(evidencePrompt(directory, "spine")).toContain("future stages must retrieve underlying text")
})


test("user instruction evidence uses the canonical request identity and rejects handoff aliases", () => {
  const directory = matter()
  const reply = (source_id: string, quote: string) => ({ entries: [{ kind: "decision", detail: "Follow the supplied instruction", evidence: [{ source_id, quote }], limitations: [] }], limitations: [] })
  enqueueLeanRequest(directory, { kind: "update", text: "Use the revised delivery provision." })
  drainLeanRequests(directory)
  writeQuestions(directory, [{ id: "q-evidence", stage: "intake", question: "Which provision?", context: "Material choice", status: "open" }])
  answerQuestion(directory, "q-evidence", { text: "Use the expressly agreed later provision." })
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
  answerQuestion(directory, "q-feedback", { text: answer })
  const reply = (quote: string) => ({ entries: [{ kind: "decision", detail: "Follow the answer", evidence: [{ source_id: "request", quote }], limitations: [] }], limitations: [] })
  expect(checkEvidence(directory, "spine", reply(answer))).toEqual([])
  const rejected = `Answer: ${answer}`
  const issues = checkEvidence(directory, "spine", reply(rejected))
  expect(issues).toHaveLength(1)
  expect(issues[0]).toContain(JSON.stringify(rejected))
  expect(issues[0]).toContain('in "request"')
  expect(issues[0]).toContain("not found in that canonical source")
})

test("a child source's provenance is a separate labelled line; plain matters' prompts are unchanged", () => {
  const directory = matter()
  const plain = { evidence: evidencePrompt(directory, "spine"), research: researchPrompt(directory, "intake", researchBasis(directory)) }
  expect(plain.evidence).not.toContain("Source provenance")
  expect(plain.research).not.toContain("provenance:")
  writeFileSync(path.join(directory, "source-index.json"), JSON.stringify({ sources: [{ document: "Thread.eml" }, { document: SOURCE, attachment_of: "Thread.eml" }] }))
  const evidence = evidencePrompt(directory, "spine")
  // The allowed source_id array is unchanged; provenance is its own line.
  expect(evidence).toContain(`Evidence source_id values: ${JSON.stringify(["request", SOURCE])}`)
  expect(evidence).toContain(`Source provenance (where each attachment or archive entry came from; display only, not part of the source names or citations): ${JSON.stringify({ [SOURCE]: "attachment of Thread.eml" })}`)
  expect(researchPrompt(directory, "intake", researchBasis(directory))).toContain(`- ${SOURCE} (provenance: attachment of Thread.eml)`)
})
