import { afterEach, expect, test } from "bun:test"
import { Database } from "bun:sqlite"
import { writeFileSync } from "node:fs"
import path from "node:path"
import { LEAN_MAX_RETRIES, researchBasis, stageOutputValid } from "./stages"
import { researchPrompt } from "./research"
import { acceptEvidenceReply, checkEvidence, checkEvidenceReply, evidencePrompt, figures, numbers } from "./evidence"
import { readLeanEvents } from "./events"
import { enqueueLeanRequest, drainLeanRequests } from "./requests"
import { leanSteps } from "./plan"
import { runLean, stageInputSha } from "./runner"
import { answerQuestion, readLeanJson, writeQuestions, writeLeanText } from "./state"
import { writeInstructionsFile } from "./instructions"
import { AGENT, cleanupMatters, happyPath, matter, scriptedModel, SOURCE, SOURCE_TEXT } from "./test-support"

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
  expect(invented).toStartWith('gaps[0].requirement quotes "Draft a long supply agreement" as from "request", but those words are not in the request')
  const [unsupported] = checkEvidence(directory, "discrepancy", reply(GAP, { ...GAP, evidence: [{ source_id: SOURCE, quote: "Invented delivery period" }] }))
  expect(unsupported).toStartWith('gaps[1].evidence[0] quotes "Invented delivery period" as from')
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


// Each figure a spine entry states is stated by one of its own quotes,
// compared as quotations are and without currency signs or separators.
const PRICED = `${SOURCE_TEXT} The price is $4,800,000.`
const priced = (detail: string, quotes: string[], extra: object = {}) =>
  ({ entries: [{ kind: "fact", detail, evidence: quotes.map(quote => ({ source_id: SOURCE, quote })), limitations: [], ...extra }], limitations: [] })
const UNQUOTED = 'entries[0] states "4,800,000", which none of its quotes states; quote the source words that state it.'

test("a spine entry's figures are stated by its own quotes, however the currency sign and separators are written", () => {
  const directory = matter(PRICED)
  expect(checkEvidence(directory, "spine", priced("The price is 4,800,000.", ["Payment is due on delivery.", "The price is $4,800,000."]))).toEqual([])
  expect(checkEvidence(directory, "spine", priced("The purchase price is $4,800,000.", ["4,800,000"]))).toEqual([])
  expect(checkEvidence(directory, "spine", priced("The purchase price is 4800000.", ["The price is $4,800,000."]))).toEqual([])
  expect(checkEvidence(directory, "spine", priced("State payment timing.", ["Payment is due on delivery."]))).toEqual([])
  expect(checkEvidence(directory, "spine", priced("The purchase price is $4,800,000.", ["Payment is due on delivery.", "within thirty days of the order date"]))).toEqual([UNQUOTED])
  // A part of a stated figure is not that figure.
  expect(checkEvidence(directory, "spine", priced("The price is 4,800.", ["The price is $4,800,000."]))).toEqual(['entries[0] states "4,800", which none of its quotes states; quote the source words that state it.'])
  // Only the spine's entries: a discrepancy entry is not held to it.
  expect(checkEvidence(directory, "discrepancy", { ...priced("The price is $4,800,000.", ["Payment is due on delivery."], { material: false, status: "benign", resolution: "Distinct terms." }), gaps: [] })).toEqual([])
})

test("a figure carries a currency sign or %, a decimal point or thousands separators, or has five digits or more", () => {
  for (const [text, figure] of [["$1,500,000", "1500000"], ["12.5", "12.5"], ["1.00%", "1.00"], ["45000", "45000"], ["15 %", "15"], ["€ 15", "15"], ["15€", "15"]])
    expect({ text, figures: figures(text) }).toEqual({ text, figures: [figure] })
  // A bare number of up to four digits (a list or part number, a count, a year), a date or a dotted reference is not.
  for (const text of ["1", "15", "2025", "§ 3", "(2)", "15/01/2025", "2025-01-15", "15.01.2025", "3.2.1", "USD 15"])
    expect({ text, figures: figures(text) }).toEqual({ text, figures: [] })
  // Every number can state a figure.
  expect(numbers("§ 3 of the 2025 terms: $1,500")).toEqual(["3", "2025", "1500"])
})

test("an id's digits are no number, and a dotted number after a reference label is a reference, not a figure", () => {
  // Digits attached to a letter, "-" or "_" are part of an id.
  for (const text of ["q-3", "src_12", "a1", "src_12345", "item A12345", "v1.2"]) {
    expect({ text, figures: figures(text) }).toEqual({ text, figures: [] })
    expect({ text, numbers: numbers(text) }).toEqual({ text, numbers: [] })
  }
  for (const text of ["Section 7.12(c)", "sections 7.12", "clause 4.2", "Art. 3.1", "para. 12.5", "§ 4.2", "§§4.2", "Schedule 2.1"])
    expect({ text, figures: figures(text) }).toEqual({ text, figures: [] })
  expect(numbers("Section 7.12(c)")).toEqual(["7.12"])
  // Money, percentages and decimals stay figures, in a text that cites a part too.
  for (const [text, stated] of [["$1,500,000", ["1500000"]], ["12.5%", ["12.5"]], ["5-10%", ["10"]], ["a rate of 2.5", ["2.5"]], ["the section's 12.5% fee", ["12.5"]],
    ["Section 3 sets a fee of $1,500.50", ["1500.50"]], ["Clause 4: 45000 units", ["45000"]]] as const)
    expect({ text, figures: figures(text) }).toEqual({ text, figures: [...stated] })
})

test("a spine entry's bare small numbers, years and dates are not asked of its quotes; a figure stated in any form is quoted", () => {
  const directory = matter(PRICED)
  expect(checkEvidence(directory, "spine", priced("Under clause 3 of the 2025 terms, delivery is due 15 days after 1 March 2026 (item 2).", ["Payment is due on delivery."]))).toEqual([])
  expect(checkEvidence(directory, "spine", priced("A 4.5% charge and 12000 units are stated.", ["Payment is due on delivery."]))).toEqual([
    'entries[0] states "4.5", which none of its quotes states; quote the source words that state it.',
    'entries[0] states "12000", which none of its quotes states; quote the source words that state it.',
  ])
  // A small number with a currency sign is a figure; a quote stating it as a bare number states it.
  const fee = matter(`${PRICED} A fee of 30 is due on signing.`)
  expect(checkEvidence(fee, "spine", priced("The fee is $30.", ["The price is $4,800,000."]))).toEqual(['entries[0] states "30", which none of its quotes states; quote the source words that state it.'])
  expect(checkEvidence(fee, "spine", priced("The fee is $30.", ["A fee of 30 is due on signing."]))).toEqual([])
})

test("on the final attempt the spine stands with its unquoted figure recorded, and a saved spine stays valid", () => {
  const directory = matter(PRICED)
  const reply = JSON.stringify(priced("The purchase price is $4,800,000.", ["Payment is due on delivery."]))
  expect(checkEvidenceReply(directory, "spine", reply)).toEqual({ issues: [UNQUOTED], recorded: [] })
  const final = checkEvidenceReply(directory, "spine", reply, { final: true })
  expect(final).toEqual({ issues: [], recorded: [UNQUOTED], value: JSON.parse(reply) })
  // Any other issue still fails the final attempt.
  expect(checkEvidenceReply(directory, "spine", JSON.stringify(priced("The price is $4,800,000.", ["Invented price"])), { final: true }).issues[0]).toContain("but those words are not in that document")
  writeLeanText(directory, "spine.json", reply)
  // A spine saved without its atoms re-runs.
  expect(stageOutputValid(directory, "spine")).toBe(false)
  writeLeanText(directory, "spine-atoms.json", JSON.stringify({ entries: JSON.parse(reply).entries }))
  expect(stageOutputValid(directory, "spine")).toBe(true)
})

test("a spine that keeps an unquoted figure is retried with the figure named, then accepted on the final attempt", async () => {
  const directory = matter(PRICED)
  const model = scriptedModel({ ...happyPath, [AGENT.spine]: () => JSON.stringify(priced("The purchase price is $4,800,000.", ["Payment is due on delivery."])) })
  expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
  const spine = model.prompts.filter(prompt => prompt.agent === AGENT.spine)
  expect(spine).toHaveLength(LEAN_MAX_RETRIES + 1)
  for (const retry of spine.slice(1)) expect(retry.text).toContain(`- ${UNQUOTED}`)
  const checks = readLeanEvents(directory).filter(event => event.stage === "spine" && event.event === "checker" && event.step === undefined)
  expect(checks.map(event => [event.attempt, event.issues])).toEqual([0, 1, 2, 3].map(attempt => [attempt, [UNQUOTED]]))
  expect((readLeanJson(directory, "spine.json") as { entries: { detail: string }[] }).entries.map(entry => entry.detail)).toEqual(["The purchase price is $4,800,000."])
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
  expect(checkEvidence(directory, "spine", reply(SOURCE, "Unsupported words"))[0]).toContain(`entries[0].evidence[0] quotes "Unsupported words" as from "${SOURCE}", but those words are not in that document.`)
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
  expect(issues[0]).toContain(`entries[0].evidence[0] quotes ${JSON.stringify(rejected)} as from "request", but those words are not in the request`)
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

test("each missed quotation is named by its entry and gets the reason it missed", () => {
  const filler = "Each party shall bear its own costs of preparing this agreement. ".repeat(6)
  const directory = matter(`The supplier shall deliver the goods within thirty days of the order date. ${filler}Payment is due on delivery of the goods.`)
  const ok = { source_id: SOURCE, quote: "Payment is due on delivery of the goods." }
  const entry = (evidence: unknown[]) => ({ kind: "decision", detail: "Follow the supplied terms", evidence, limitations: [] })
  const check = (...evidence: unknown[]) => checkEvidence(directory, "spine", { entries: [entry([ok]), entry([ok, ...evidence])], limitations: [] })
  // An elision across a long omitted stretch is two passages, not one quotation.
  const [apart] = check({ source_id: SOURCE, quote: "The supplier shall deliver the goods ... Payment is due on delivery" })
  expect(apart).toStartWith(`entries[1].evidence[1] quotes "The supplier shall deliver the goods ... Payment is due on delivery" as from "${SOURCE}"`)
  expect(apart).toContain("in separate passages. Quote each passage as its own quotation, without an ellipsis.")
  // A change of case alone shows the source's own words to copy.
  const [cased] = check({ source_id: SOURCE, quote: "the supplier shall deliver the goods within thirty days" })
  expect(cased).toContain('The document reads: "The supplier shall deliver the goods within thirty days". Copy those exact characters, or write [x] for a change of case.')
})
