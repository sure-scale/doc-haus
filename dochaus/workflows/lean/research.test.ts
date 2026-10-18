import { afterEach, expect, test } from "bun:test"
import { existsSync, writeFileSync } from "node:fs"
import path from "node:path"
import { createHash } from "node:crypto"
import { preserveResearchSourceSnapshot } from "../../lib/research-source-snapshot"
import { citationRowsForLeanText, citationRowsForText } from "../../lib/citation-extraction"
import { acceptResearchReply, approvedResearch, checkResearch, researchHandoff, researchPermissions, researchPrompt, researchContextLocators } from "./research"
import { finalGate } from "./gate"
import { quoteFound } from "./quotes"
import { enqueueLeanRequest } from "./requests"
import { runLean } from "./runner"
import type { StepID } from "./plan"
import { readRepairRound } from "./repairs"
import { LEAN_PROMPT_VERSION, researchBasis, stageInputs } from "./stages"
import { hashStageInputs, leanFile, readLeanJson, readLeanState, readLeanText, readOpenItems, readQuestions, replaceStageOpenItems, writeLeanState, writeLeanText } from "./state"
import { AGENT, cleanupMatters, happyPath, matter, scriptedModel, type Reply } from "./test-support"

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
  expect(acceptResearchReply(directory, "research-review", JSON.stringify({ approved: ["r1"], rejected: [], limitations: ["Currency unconfirmed"], gaps: [] })).issues).toEqual([])
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
  const accepted = acceptResearchReply(directory, "research-review", JSON.stringify({ approved: [], rejected: [{ id: "r1", reason: "Not applicable to this jurisdiction" }], limitations: [], gaps: [] }))
  expect(accepted.issues).toEqual([])
  replaceStageOpenItems(directory, "research-review", null, accepted.openItems)
  expect(approvedResearch(directory)).toEqual([])
  writeLeanText(directory, "research-review.json", JSON.stringify({ approved: ["r1"], rejected: [] }))
  expect(approvedResearch(directory)).toEqual([])
  acceptResearchReply(directory, "research-review", JSON.stringify({ approved: [], rejected: [{ id: "r1", reason: "Not applicable" }], limitations: [], gaps: [] }))
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
  expect(researchPrompt(directory, "research-review", researchBasis(directory))).toContain("research-context.txt")
  expect(researchPrompt(directory, "research-review", researchBasis(directory))).toContain("reject that finding with the access limitation")
  expect(checkResearch(directory, { ...research, findings: [{ ...research.findings[0], quote: "Invented wording" }] })).not.toEqual([])
})

const GAP = { issue: "Authority on remedies for late delivery under these supply terms has not been retrieved.", affects: "Remedy wording in the delivery clause." }
const LATER = { issue: "Authority on the timing of payment default under these supply terms has not been retrieved.", affects: "Payment clause default wording." }
const GAP_TEXT = `${GAP.issue} Affects: ${GAP.affects}`
const targeted = (issues: unknown[]) => `New authority issues requiring targeted research: ${JSON.stringify(issues)}\n`
const reviewed = (gaps: { issue: string; affects: string }[]) => JSON.stringify({ approved: [], rejected: [], limitations: [], gaps })
// Each research pass reports in its own words, so its review sees new research.
const researched: Reply = (_, index) => JSON.stringify({ status: "not-needed", findings: [], limitations: [`Research pass ${index + 1} relied on the supplied terms only.`] })
const researchAgents = (agents: string[]) => agents.filter(agent => agent === AGENT.research || agent === AGENT["research-review"])

test("research review reports researchable gaps separately and the checker validates their shape", () => {
  const directory = matter()
  acceptResearchReply(directory, "research", JSON.stringify(retrieved(directory)))
  expect(acceptResearchReply(directory, "research-review", JSON.stringify({ approved: ["r1"], rejected: [], limitations: [] })).issues).not.toEqual([])
  expect(acceptResearchReply(directory, "research-review", JSON.stringify({ approved: ["r1"], rejected: [], limitations: [], gaps: [{ issue: "Unretrieved authority", affects: "" }] })).issues).not.toEqual([])
  expect(acceptResearchReply(directory, "research-review", JSON.stringify({ approved: ["r1"], rejected: [], limitations: [], gaps: ["Unretrieved authority"] })).issues).not.toEqual([])
  const accepted = acceptResearchReply(directory, "research-review", JSON.stringify({ approved: ["r1"], rejected: [], limitations: [], gaps: [GAP] }))
  expect(accepted.issues).toEqual([])
  expect(accepted.openItems).toEqual([expect.objectContaining({ stage: "research-review", kind: "research_gap", detail: GAP_TEXT })])
  expect(JSON.parse(researchHandoff(directory)).limitations).toEqual(["Currency remains to be confirmed", GAP_TEXT])
  const prompt = researchPrompt(directory, "research-review", researchBasis(directory))
  expect(prompt).toContain("A failed or unavailable provider, an inaccessible source, incomplete extraction and currency limits are limitations; a gap is material authority not yet retrieved that a targeted retrieval could find and on which a drafting choice depends.")
  expect(prompt).toContain('"limitations":["other remaining legal research limits the lawyer must see"],"gaps":[{"issue":"material authority not yet retrieved","affects":"the drafting choice that depends on it"}]}.')
  // A review saved before gaps were reported remains valid.
  writeLeanText(directory, "research-review.json", JSON.stringify({ approved: ["r1"], rejected: [], limitations: [] }))
  expect(approvedResearch(directory)).toHaveLength(1)
})

test("first review gaps reach research once as targeted issues before the plan", async () => {
  const directory = matter()
  const model = scriptedModel({ ...happyPath, [AGENT.research]: researched, [AGENT["research-review"]]: (_, index) => reviewed(index ? [] : [GAP]) })
  expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
  expect(researchAgents(model.sessions.map(row => row.agent))).toEqual([AGENT.research, AGENT["research-review"], AGENT.research, AGENT["research-review"]])
  const agents = model.sessions.map(row => row.agent)
  expect(agents.lastIndexOf(AGENT["research-review"])).toBeLessThan(agents.indexOf(AGENT.spine))
  const prompts = model.prompts.filter(row => row.agent === AGENT.research).map(row => row.text)
  expect(prompts[0]).toContain(`${targeted([])}Previously verified research when repairing a new authority issue: No targeted research repair.\n`)
  // The round's research also receives the first pass's research with the gap as a limit.
  const handoff = JSON.stringify({ findings: [], limitations: ["Research pass 1 relied on the supplied terms only.", GAP_TEXT] }, null, 2)
  expect(prompts[1]).toContain(`${targeted([GAP])}Previously verified research when repairing a new authority issue: ${handoff}\n`)
  expect(readLeanJson(directory, "research-gaps.json")).toEqual({ basis: researchBasis(directory), gaps: [GAP] })
  expect(stageInputs(directory, "research").filter(part => part.name === "research-gaps").map(part => part.bytes)).toEqual([JSON.stringify([GAP])])
  expect(readOpenItems(directory).filter(item => item.kind === "research_gap")).toEqual([])
}, 120_000)

test("second review gaps are disclosed without a further research round", async () => {
  const directory = matter()
  const model = scriptedModel({ ...happyPath, [AGENT.research]: researched, [AGENT["research-review"]]: (_, index) => reviewed([index ? LATER : GAP]) })
  expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
  expect(model.count(AGENT.research)).toBe(2)
  expect(model.count(AGENT["research-review"])).toBe(2)
  const gaps = readOpenItems(directory).filter(item => item.kind === "research_gap")
  expect(gaps.map(item => item.detail)).toEqual([expect.stringContaining(LATER.issue)])
  expect(finalGate(directory)).toMatchObject({ delivery: "review_draft", disclosed: expect.arrayContaining([expect.objectContaining({ kind: "research_gap", detail: expect.stringContaining(LATER.issue) })]) })
}, 120_000)

test("a review without gaps keeps the research steps and inputs unchanged", async () => {
  const directory = matter()
  const model = scriptedModel(happyPath)
  expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
  expect(model.sessions.map(row => row.agent).slice(0, 5)).toEqual([AGENT.discrepancy, AGENT.intake, AGENT.research, AGENT["research-review"], AGENT.spine])
  expect(model.count(AGENT.research)).toBe(1)
  expect(model.count(AGENT["research-review"])).toBe(1)
  expect(stageInputs(directory, "research").map(part => part.name)).not.toContain("research-gaps")
  expect(existsSync(leanFile(directory, "research-gaps.json"))).toBe(false)
}, 120_000)

test("a resumed or repeated run with unchanged inputs repeats no research round", async () => {
  const directory = matter()
  const failing = scriptedModel({ ...happyPath, [AGENT["research-review"]]: () => reviewed([GAP]),
    [AGENT.research]: (_, index) => {
      if (index) throw new Error("Research provider connection lost.")
      return JSON.stringify({ status: "not-needed", findings: [], limitations: ["This fixture requests a note from supplied terms only."] })
    } })
  expect((await runLean(directory, { sessions: failing.port })).status).toBe("failed")
  expect(failing.count(AGENT["research-review"])).toBe(1)
  enqueueLeanRequest(directory, { kind: "continue" })
  const resumed = scriptedModel({ ...happyPath, [AGENT.research]: researched, [AGENT["research-review"]]: () => reviewed([LATER]) })
  expect((await runLean(directory, { sessions: resumed.port })).status).toBe("completed")
  expect(researchAgents(resumed.sessions.map(row => row.agent))).toEqual([AGENT.research, AGENT["research-review"]])
  expect(resumed.prompts.find(row => row.agent === AGENT.research)?.text).toContain(GAP.issue)
  enqueueLeanRequest(directory, { kind: "continue" })
  const repeated = scriptedModel({})
  expect((await runLean(directory, { sessions: repeated.port })).status).toBe("completed")
  expect(repeated.sessions).toEqual([])
}, 120_000)

test("a later package-review research repair still reaches research after the gap round", async () => {
  const directory = matter()
  const model = scriptedModel({ ...happyPath, [AGENT.research]: researched,
    [AGENT["research-review"]]: (_, index) => reviewed(index ? [] : [GAP]),
    [AGENT["package-review"]]: (_, index) => JSON.stringify(index ? { status: "pass", delivery: [], findings: [], limitations: [] } : {
      status: "changes-required", delivery: [], findings: [{ target: "research", detail: "Resolve the newly identified authority applicability issue from actual sources." }], limitations: [],
    }),
  })
  expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
  expect(model.count(AGENT.research)).toBe(3)
  expect(model.count(AGENT["research-review"])).toBe(3)
  expect(model.prompts.filter(row => row.agent === AGENT.research).at(-1)?.text).toContain("newly identified authority applicability")
}, 120_000)

test("research unchanged by the gap round keeps its review and that review's disclosed gaps", async () => {
  const directory = matter()
  const model = scriptedModel({ ...happyPath, [AGENT["research-review"]]: () => reviewed([GAP]) })
  expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
  expect(researchAgents(model.sessions.map(row => row.agent))).toEqual([AGENT.research, AGENT["research-review"], AGENT.research])
  expect(readOpenItems(directory).filter(item => item.kind === "research_gap").map(item => item.detail)).toEqual([expect.stringContaining(GAP.issue)])
}, 120_000)

test("a review accepted with gaps before a crash starts its one round when the review is skipped on resume", async () => {
  const directory = matter()
  expect((await runLean(directory, { sessions: scriptedModel(happyPath).port })).status).toBe("completed")
  // The review was accepted with gaps; the run stopped before the round was written.
  writeLeanText(directory, "research-review.json", `${reviewed([GAP])}\n`)
  expect(existsSync(leanFile(directory, "research-gaps.json"))).toBe(false)
  enqueueLeanRequest(directory, { kind: "continue" })
  const resumed = scriptedModel({ ...happyPath, [AGENT.research]: researched })
  expect((await runLean(directory, { sessions: resumed.port })).status).toBe("completed")
  expect(researchAgents(resumed.sessions.map(row => row.agent))).toEqual([AGENT.research, AGENT["research-review"]])
  expect(resumed.prompts.find(row => row.agent === AGENT.research)?.text).toContain(targeted([GAP]))
  expect(readLeanJson(directory, "research-gaps.json")).toEqual({ basis: researchBasis(directory), gaps: [GAP] })
  enqueueLeanRequest(directory, { kind: "continue" })
  const repeated = scriptedModel({})
  expect((await runLean(directory, { sessions: repeated.port })).status).toBe("completed")
  expect(repeated.sessions).toEqual([])
}, 120_000)

test("an existing run whose saved review predates gaps keeps its research keys and re-runs nothing", async () => {
  const directory = matter()
  expect((await runLean(directory, { sessions: scriptedModel(happyPath).port })).status).toBe("completed")
  // The run as saved before gaps were reported: its review has no gaps field,
  // and research's key had no gap input.
  const { gaps, ...saved } = readLeanJson(directory, "research-review.json") as Record<string, unknown>
  expect(gaps).toEqual([])
  writeLeanText(directory, "research-review.json", `${JSON.stringify(saved, null, 2)}\n`)
  const legacy = readLeanState(directory)!
  writeLeanState(directory, { ...legacy, stages: Object.fromEntries(Object.entries(legacy.stages).map(([step, row]) => [step, row?.status === "done"
    ? { ...row, input_sha: hashStageInputs(step as StepID, LEAN_PROMPT_VERSION, stageInputs(directory, step as StepID).filter(part => part.name !== "research-gaps")) } : row])) })
  const keys = (({ research, "research-review": review }) => [research?.input_sha, review?.input_sha])(readLeanState(directory)!.stages)
  enqueueLeanRequest(directory, { kind: "continue" })
  const resumed = scriptedModel({})
  expect((await runLean(directory, { sessions: resumed.port })).status).toBe("completed")
  expect(resumed.sessions).toEqual([])
  expect((({ research, "research-review": review }) => [research?.input_sha, review?.input_sha])(readLeanState(directory)!.stages)).toEqual(keys)
  expect(existsSync(leanFile(directory, "research-gaps.json"))).toBe(false)
}, 120_000)

test("a package-review research repair after the gap round gives the new basis one round, then discloses its gaps", async () => {
  const directory = matter()
  const REMAINING = { issue: "Authority on the effect of a late termination notice under these supply terms has not been retrieved.", affects: "Termination notice wording." }
  const model = scriptedModel({ ...happyPath, [AGENT.research]: researched,
    // The first review of each basis names gaps, and so does the review of the repair's round.
    [AGENT["research-review"]]: (_, index) => reviewed([[GAP], [], [LATER], [REMAINING]][index]!),
    [AGENT["package-review"]]: (_, index) => JSON.stringify(index ? { status: "pass", delivery: [], findings: [], limitations: [] } : {
      status: "changes-required", delivery: [], findings: [{ target: "research", detail: "Resolve the newly identified authority applicability issue from actual sources." }], limitations: [],
    }),
  })
  expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
  expect(model.count(AGENT.research)).toBe(4)
  expect(model.count(AGENT["research-review"])).toBe(4)
  const repairs = readRepairRound(directory)!.items.filter(item => item.ref.repair_target === "research")
  expect(repairs.map(item => item.detail)).toEqual(["Resolve the newly identified authority applicability issue from actual sources."])
  const prompts = model.prompts.filter(row => row.agent === AGENT.research).map(row => row.text)
  expect(prompts.map(text => text.split("\n").find(line => line.startsWith("New authority issues requiring targeted research: ")))).toEqual(
    [[], [GAP], repairs, [...repairs, LATER]].map(issues => targeted(issues).trimEnd()))
  expect(readLeanJson(directory, "research-gaps.json")).toEqual({ basis: researchBasis(directory), gaps: [LATER] })
  expect(readOpenItems(directory).filter(item => item.kind === "research_gap").map(item => item.detail)).toEqual([`${REMAINING.issue} Affects: ${REMAINING.affects}`])
  enqueueLeanRequest(directory, { kind: "continue" })
  const repeated = scriptedModel({})
  expect((await runLean(directory, { sessions: repeated.port })).status).toBe("completed")
  expect(repeated.sessions).toEqual([])
}, 120_000)
