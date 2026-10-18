import { afterEach, expect, test } from "bun:test"
import { writeFileSync } from "node:fs"
import path from "node:path"
import { createHash } from "node:crypto"
import { preserveResearchSourceSnapshot } from "../../lib/research-source-snapshot"
import { citationRowsForLeanText, citationRowsForText } from "../../lib/citation-extraction"
import { acceptResearchReply, approvedResearch, checkResearch, researchHandoff, researchPermissions, researchPrompt, researchContextLocators } from "./research"
import { checkDraftingNotes } from "./checks"
import { finalGate } from "./gate"
import { quoteFound } from "./quotes"
import { enqueueLeanRequest } from "./requests"
import { runLean } from "./runner"
import { stageInputs } from "./stages"
import { leanFile, readLeanState, readLeanText, readQuestions, replaceStageOpenItems, writeLeanState, writeLeanText } from "./state"
import { AGENT, cleanupMatters, happyPath, matter, scriptedModel } from "./test-support"

afterEach(cleanupMatters)

function retrieved(directory: string, text = "A party must give written notice before termination.") {
  const hash = (value: string) => createHash("sha256").update(value).digest("hex")
  const url = "https://example.gov/legal-notice"
  const snapshot = preserveResearchSourceSnapshot(directory, { status: "ok", text, extraction: "text", sourceBytes: Buffer.from(text),
    descriptor: { originalUrl: url, finalUrl: url, host: "example.gov", fetchedAt: "2026-10-12T00:00:00Z", contentType: "text/plain", byteLength: Buffer.byteLength(text), bodyHash: hash(text), textHash: hash(text), textChars: text.length } }, [{ start: 0, end: text.length, text, term: "notice" }])
  const finding = { id: "r1", claim: "The provision requires notice", title: "Notice provision", citation: "", url, snapshot: snapshot.sha256, quote: text,
    jurisdiction: "Jurisdiction supplied in intake", currentness: "Retrieved version; later changes not independently checked", limitations: "Currency remains to be confirmed" }
  return { status: "limited", findings: [finding], limitations: ["Currency remains to be confirmed"] }
}

test("unresolved jurisdiction pauses intake before research tools or drafting sessions exist", async () => {
  const directory = matter()
  const model = scriptedModel({ ...happyPath, [AGENT.intake]: () => JSON.stringify({ questions: [{ question: "Which jurisdiction governs?", context: "The legal approach depends on the governing law." }] }) })
  const result = await runLean(directory, { sessions: model.port })
  expect(result.status).toBe("blocked")
  expect(model.sessions.map(row => row.agent)).toEqual([AGENT.discrepancy, AGENT.intake])
  expect(model.sessions[0]?.permissions?.some(row => row.action === "official-source")).toBe(false)
  expect(model.count(AGENT.research)).toBe(0)
  expect(readQuestions(directory)[0]?.stage).toBe("intake")
})

test("supplied intake answer resumes research with the answer and invalidates research on later instruction", async () => {
  const directory = matter()
  const initial = scriptedModel({ [AGENT.intake]: () => JSON.stringify({ questions: [{ question: "Which jurisdiction governs?", context: "Needed to research law" }] }) })
  await runLean(directory, { sessions: initial.port })
  enqueueLeanRequest(directory, { kind: "answer", question_id: readQuestions(directory)[0]!.id, answer: "Jurisdiction supplied by lawyer" })
  const model = scriptedModel(happyPath)
  expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
  expect(model.prompts.find(row => row.agent === AGENT.research)?.text).toContain("Jurisdiction supplied by lawyer")
  expect(model.sessions.find(row => row.agent === AGENT.research)?.permissions).toEqual(researchPermissions)
  const prior = JSON.stringify(stageInputs(directory, "research"))
  writeLeanText(directory, "updated-notes.md", "Use the newly instructed governing jurisdiction.")
  expect(JSON.stringify(stageInputs(directory, "research"))).not.toBe(prior)
}, 120_000)

test("retrieved text verifies exact quotations, while URLs, fabricated quotes and mismatched sources fail", () => {
  const directory = matter()
  const research = retrieved(directory)
  expect(checkResearch(directory, research)).toEqual([])
  expect(checkResearch(directory, { ...research, findings: [{ ...research.findings[0], quote: "Invented text" }] })).not.toEqual([])
  expect(checkResearch(directory, { ...research, findings: [{ ...research.findings[0], url: "https://example.gov/other" }] })).not.toEqual([])
  expect(checkResearch(directory, { ...research, findings: [{ ...research.findings[0], snapshot: "0".repeat(64) }] })).not.toEqual([])
  expect(acceptResearchReply(directory, "research", JSON.stringify(research)).issues).toEqual([])
  expect(approvedResearch(directory)).toEqual([])
  expect(acceptResearchReply(directory, "research-review", JSON.stringify({ approved: ["r1"], rejected: [], limitations: ["Currency unconfirmed"] })).issues).toEqual([])
  expect(approvedResearch(directory)).toHaveLength(1)
  expect(quoteFound(directory, { source_id: research.findings[0]!.url, quote: research.findings[0]!.quote })).toBe(true)
  expect(quoteFound(directory, { source_id: research.findings[0]!.url, quote: "Invented text" })).toBe(false)
  expect(citationRowsForLeanText(directory, "draft", research.findings[0]!.url)[0]?.status).toBe("verified")
  expect(researchHandoff(directory)).toContain("Currency")
})

test("needed research unavailable is disclosed and cannot become execution-ready by disappearing from body", async () => {
  const directory = matter()
  const model = scriptedModel({ ...happyPath, [AGENT.research]: () => JSON.stringify({ status: "limited", findings: [], limitations: ["Research provider unavailable; authority not verified."] }) })
  expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
  expect(finalGate(directory)).toMatchObject({ delivery: "review_draft", disclosed: [expect.objectContaining({ kind: "research_gap" })] })
  writeLeanText(directory, "drafting-notes.md", "# Notes\n\nNo limitations.")
  expect(finalGate(directory)).toMatchObject({ status: "blocked", blocking: [expect.objectContaining({ kind: "research_gap" })] })
  expect(checkResearch(directory, { status: "completed", findings: [], limitations: [] }).length).toBeGreaterThan(0)
}, 120_000)

test("research verifier rejects unsupported proposition without making it drafting authority", () => {
  const directory = matter()
  const research = retrieved(directory)
  acceptResearchReply(directory, "research", JSON.stringify(research))
  const accepted = acceptResearchReply(directory, "research-review", JSON.stringify({ approved: [], rejected: [{ id: "r1", reason: "Not applicable to this jurisdiction" }], limitations: [] }))
  expect(accepted.issues).toEqual([])
  replaceStageOpenItems(directory, "research-review", null, accepted.openItems)
  expect(approvedResearch(directory)).toEqual([])
  writeLeanText(directory, "research-review.json", JSON.stringify({ approved: ["r1"], rejected: [] }))
  expect(approvedResearch(directory)).toEqual([])
  acceptResearchReply(directory, "research-review", JSON.stringify({ approved: [], rejected: [{ id: "r1", reason: "Not applicable" }], limitations: [] }))
  expect(researchHandoff(directory)).toContain("Not applicable")
  writeFileSync(path.join(directory, ".dochaus/context-manifest.jsonl"), JSON.stringify({ entry_id: "earlier-fetch", event: "source_fetch", url: research.findings[0]!.url }) + "\n")
  expect(citationRowsForText(directory, "draft", research.findings[0]!.url)[0]?.status).toBe("verified")
  expect(citationRowsForLeanText(directory, "draft", research.findings[0]!.url)[0]?.status).toBe("unsupported")
})

test("completed legacy runs remain unchanged and do not claim research", async () => {
  const directory = matter()
  const current = readLeanState(directory)!
  const { intake, research, "research-review": review, ...stages } = current.stages
  writeLeanState(directory, { ...current, status: "completed", stages })
  const before = Bun.file(leanFile(directory, "state.json"))
  const content = await before.text()
  const model = scriptedModel({})
  expect((await runLean(directory, { sessions: model.port })).stages.research).toBeUndefined()
  expect(await before.text()).toBe(content)
  expect(model.sessions).toEqual([])
  enqueueLeanRequest(directory, { kind: "continue" })
  expect((await runLean(directory, { sessions: model.port })).stages.research).toBeUndefined()
  expect(await before.text()).toBe(content)
  expect(model.sessions).toEqual([])
})


test("long canonical source text is readable beyond JSON clipping with focused source-context locators", () => {
  const directory = matter()
  const passage = "The notice requirement does not apply where the preceding exception is satisfied."
  const text = "Source context before the provision. ".repeat(200) + passage + " Context after the provision. ".repeat(100)
  const research = retrieved(directory, text)
  research.findings[0]!.quote = passage
  expect(checkResearch(directory, research)).toEqual([])
  const accepted = acceptResearchReply(directory, "research", JSON.stringify(research))
  expect(accepted.issues).toEqual([])
  expect(accepted.outputs).toContain("research-context.txt")
  const projected = readLeanText(directory, "research-context.txt")!
  const lines = projected.split("\n")
  expect(lines.every(line => line.length <= 500)).toBe(true)
  const locator = researchContextLocators(directory)[0]!
  expect(locator).toMatchObject({ id: "r1", path: ".dochaus/lean/research-context.txt" })
  if (!("read" in locator) || !locator.read) throw new Error("Missing focused read locator")
  // Native read clips each long line and the whole response; this focused
  // range stays below both caps while exposing a late quotation and context.
  const readable = lines.slice(locator.read.offset - 1, locator.read.offset - 1 + locator.read.limit).map(line => line.slice(0, 2000)).join("\n")
  expect(locator.read.offset).toBeGreaterThan(10)
  expect(readable.length).toBeLessThan(6000)
  expect(readable).toContain(passage)
  expect(readable).toContain("Source context before the provision")
  expect(readable).toContain("Context after the provision")
  const extracted = projected.split("Canonical extracted text:\n")[1]!.split("End of source text.")[0]!.replaceAll("\n", "")
  expect(extracted).toBe(text)
  expect(researchPrompt(directory, "research-review")).toContain("research-context.txt")
  expect(researchPrompt(directory, "research-review")).toContain("reject that finding with the access limitation")
  expect(checkResearch(directory, { ...research, findings: [{ ...research.findings[0], quote: "Invented wording" }] })).not.toEqual([])
})

test("drafting notes may cite an approved authority by name after a signal word and quote it with an ellipsis", () => {
  const directory = matter()
  const text = "The tenant must keep the premises insured against the listed risks (see the schedule; Ashdown v. Merrow, 12 Synth. 34 [2020]). A subtenant is bound by the same insurance covenant."
  const research = retrieved(directory, text)
  research.findings[0]!.citation = "Ashdown v. Merrow, 12 Synth. 34 (2020)"
  expect(acceptResearchReply(directory, "research", JSON.stringify(research)).issues).toEqual([])
  expect(acceptResearchReply(directory, "research-review", JSON.stringify({ approved: ["r1"], rejected: [], limitations: [] })).issues).toEqual([])
  const url = research.findings[0]!.url
  const notes = (body: string) => checkDraftingNotes(directory, `# Drafting notes\n\n${body}`, []).issues
  const elided = "The tenant must keep the premises insured against the listed risks \u2026 A subtenant is bound by the same insurance covenant"
  expect(notes(`In Ashdown v. Merrow, the court held that "${elided}" [source: ${url}].`)).toEqual([])
  expect(notes(`See Ashdown v. Merrow, 12 Synth. 34 (2020), for the insurance covenant.`)).toEqual([])
  expect(citationRowsForLeanText(directory, "drafting-notes.md", "In Ashdown v. Merrow, the covenant binds.").map(row => [row.citation_text, row.status]))
    .toEqual([["Ashdown v. Merrow", "verified"]])
  // An authority research did not approve is still refused under its own name.
  expect(notes("In Ashdown v. Kestrel, the covenant binds.")).toEqual([
    `The drafting notes cite "Ashdown v. Kestrel", which no matter source, research row or verified authority supports. Remove the citation or describe the point as needing research.`,
  ])
  // Dropping a signal never lets a citation key match another case: "In re"
  // is not part of a case name, "The" can begin one and stays, and the name
  // left after a signal must start a word of the support text.
  const unsupported = (citation: string) => `The drafting notes cite "${citation}", which no matter source, research row or verified authority supports. Remove the citation or describe the point as needing research.`
  expect(notes("In re Ashdown v. Kestrel, the covenant binds.")).toEqual([unsupported("Ashdown v. Kestrel")])
  expect(notes("In Re Ashdown v. Merrow, the covenant binds.")).toEqual([unsupported("In Re Ashdown v. Merrow")])
  expect(notes("The Ashdown v. Merrow decision binds the subtenant.")).toEqual([unsupported("The Ashdown v. Merrow")])
  expect(notes("See Down v. Merrow, on the covenant.")).toEqual([unsupported("Down v. Merrow")])
  expect(notes("Under Ashdown v. Merrow, the covenant binds.")).toEqual([])
  const reordered = "A subtenant is bound by the same insurance covenant \u2026 The tenant must keep the premises insured"
  const refused = notes(`The covenant reads "${reordered}" [source: ${url}].`)
  expect(refused).toHaveLength(1)
  expect(refused[0]).toContain(`quotes "${reordered}" as from "${url}", but those words are not in that document.`)
})
