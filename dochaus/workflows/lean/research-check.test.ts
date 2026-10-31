import { afterEach, expect, test } from "bun:test"
import { createHash } from "node:crypto"
import { preserveResearchSourceSnapshot } from "../../lib/research-source-snapshot"
import { evidencePrompt } from "./evidence"
import { approvedResearch } from "./research"
import { checkAdoptedAgainstResearch } from "./research-check"
import { LEAN_MAX_RETRIES, type StageSessionPort } from "./stages"
import { writeLeanText, writeQuestions, type LeanState, type Question } from "./state"
import { cleanupMatters, matter } from "./test-support"

// A recommendation the lawyer delegated is adopted before legal research. Once
// the research is approved, a focused job checks it against that research
// alone; when the research shows it would not take effect as stated, the
// spine is told to draft to the offered option the research supports instead.
// Synthetic matter; only the model is scripted.

afterEach(cleanupMatters)

const LAW = "Only the managing committee may appoint the treasurer."
const URL = "https://example.gov/club-rules"
const ADOPTED = "The founding member appoints the treasurer in the formation record."
const OTHER = "The managing committee appoints the treasurer in its first resolution."
const QUESTION: Question = { id: "q-treasurer", stage: "intake", question: "Who should appoint the treasurer?", context: "The founder wants one record.", status: "answered", kind: "judgment",
  advice: { options: [{ id: "o1", text: ADOPTED }, { id: "o2", text: OTHER }], recommendation: { option: "o1", rationale: "One record is simpler.", basis: [] } },
  answer: ADOPTED, answer_kind: "adopted", option_id: "o1", answer_class: "delegated", lawyer_words: "Use your judgment." }

function approved(directory: string) {
  const hash = (value: string) => createHash("sha256").update(value).digest("hex")
  const snapshot = preserveResearchSourceSnapshot(directory, { status: "ok", text: LAW, extraction: "text", sourceBytes: Buffer.from(LAW),
    descriptor: { originalUrl: URL, finalUrl: URL, host: "example.gov", fetchedAt: "2026-10-12T00:00:00Z", contentType: "text/plain", byteLength: Buffer.byteLength(LAW), bodyHash: hash(LAW), textHash: hash(LAW), textChars: LAW.length } }, [{ start: 0, end: LAW.length, text: LAW, term: "treasurer" }])
  writeLeanText(directory, "research.json", JSON.stringify({ status: "completed", limitations: [], findings: [{ id: "r1", claim: "The committee appoints the treasurer", title: "Club rules", citation: "", url: URL, snapshot: snapshot.sha256, quote: LAW,
    jurisdiction: "Supplied", currentness: "Retrieved version", limitations: "" }] }))
  writeLeanText(directory, "research-review.json", JSON.stringify({ approved: ["r1"], rejected: [], limitations: [], gaps: [] }))
  expect(approvedResearch(directory)).toHaveLength(1)
}

function scripted(replies: string[]) {
  const prompts: string[] = []
  const port: StageSessionPort = {
    createSession: async () => "ses-check",
    prompt: async ({ text }) => { prompts.push(text); return replies[Math.min(prompts.length - 1, replies.length - 1)]! },
  } as StageSessionPort
  return { port, prompts }
}

const state = { model: null } as unknown as LeanState

test("an adopted recommendation the approved research shows would not take effect is replaced in the spine prompt by the option the research supports", async () => {
  const directory = matter()
  writeQuestions(directory, [QUESTION])
  approved(directory)
  const model = scripted([JSON.stringify({ holds: false, option: "o2", reason: "The research gives the appointment to the committee.", evidence: [{ source_id: URL, quote: LAW }] })])
  await checkAdoptedAgainstResearch(directory, state, { sessions: model.port })
  expect(model.prompts[0]).toContain(LAW)
  const prompt = evidencePrompt(directory, "spine")
  expect(prompt).toContain(`"${ADOPTED}", would not take effect as stated`)
  expect(prompt).toContain(`Draft to "${OTHER}" instead`)
  // Checked once against this research: the next spine run asks nothing.
  await checkAdoptedAgainstResearch(directory, state, { sessions: model.port })
  expect(model.prompts).toHaveLength(1)
})

test("an adopted recommendation stands when the job finds it holds or never quotes the approved research", async () => {
  for (const reply of [JSON.stringify({ holds: true }), JSON.stringify({ holds: false, option: "o2", reason: "Committees usually appoint.", evidence: [{ source_id: URL, quote: "Committees usually appoint officers." }] })]) {
    const directory = matter()
    writeQuestions(directory, [QUESTION])
    approved(directory)
    const model = scripted([reply])
    await checkAdoptedAgainstResearch(directory, state, { sessions: model.port })
    expect(model.prompts).toHaveLength(reply.includes("usually") ? LEAN_MAX_RETRIES + 1 : 1)
    expect(evidencePrompt(directory, "spine")).not.toContain("Draft to")
  }
})
