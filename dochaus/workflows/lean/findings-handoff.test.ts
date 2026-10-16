import { afterEach, expect, test } from "bun:test"
import { acceptChangePlan, checkChangePlan, deferredFindingItems, readChangePlan, reviewDispositionText, type ChangePlan, type ReviewDisposition } from "./change-pipeline"
import { readParagraphs } from "./checks"
import { finalGate } from "./gate"
import { packageReviewPrompt } from "./package-review"
import { writeLeanPlan } from "./plan"
import { completedDocumentsText, stagePrompt } from "./prompts"
import { runLean, stageInputSha } from "./runner"
import { disclosableItems } from "./stages"
import { readLeanJson, readOpenItems, writeLeanText } from "./state"
import { AGENT, APPLIED, DRAFT, LETTER, SUPPORT, changesFirst, cleanupMatters, drafter, happyPath, matter, pidOf, reviewFirst, scriptedModel } from "./test-support"

afterEach(cleanupMatters)

const findings: ReviewDisposition[] = [
  { disposition: "deferred", kind: "research", detail: "Check whether an external enforceability rule qualifies the supplied delivery obligation.", reason: "No verified authority resolves that proposition; obtain targeted research before relying on it.", documents: ["d1"], sources: [], limitations: ["Enforceability has not been verified."] },
  { disposition: "deferred", kind: "information", detail: "Confirm the recipient for delivery notices.", reason: "No recipient is supplied; complete this before issue.", documents: ["d1"], sources: [], limitations: ["Recipient not supplied."] },
  { disposition: "recommendation", kind: "suggestion", detail: "Consider a reminder before the supplied delivery deadline.", reason: "This is an optional practical proposal; the agreed period stays unchanged.", documents: ["d1"], sources: SUPPORT, limitations: [] },
  { disposition: "rejected", kind: "suggestion", detail: "Replace the supplied delivery period with immediate delivery.", reason: "The source specifies a period; no instruction supports shortening it.", documents: ["d1"], sources: SUPPORT, limitations: [] },
]
const plan = (rows = findings): ChangePlan => ({ shared_decisions: [], groups: [], dismissed: [], questions: [], findings: structuredClone(rows) })

function setup() {
  const directory = matter()
  writeLeanPlan(directory, [
    { title: "Supply note", purpose: "Record supplied terms", parts: [{ heading: "Delivery", scope: "Supplied obligation" }] },
    { title: "Client advice", purpose: "Explain the supply note and remaining choices", parts: [{ heading: "Advice", scope: "Consequences and remaining action" }] },
  ])
  writeLeanText(directory, "d1/paragraphs.json", JSON.stringify([{ pid: "p111111111111", sha256: "initial", text: "Delivery is due under the supplied terms." }]))
  writeLeanText(directory, "d1/paragraphs-applied.json", JSON.stringify([{ pid: "p222222222222", sha256: "final", text: "Delivery remains due under the supplied terms." }]))
  expect(acceptChangePlan(directory, "d1", JSON.stringify(plan())).issues).toEqual([])
  return directory
}

test("dispositions survive fresh later writers, notes and review without making optional or rejected suggestions open decisions", () => {
  const directory = setup()
  expect(readChangePlan(directory, "d1")?.findings).toEqual(findings)
  for (const prompt of [completedDocumentsText(directory, "d2"), stagePrompt(directory, "d2.draft"), stagePrompt(directory, "finalize"), packageReviewPrompt(directory)]) {
    findings.forEach(row => expect(prompt).toContain(row.detail))
    expect(prompt).toContain('"origin_document": "d1"')
    expect(prompt).toContain('"disposition": "rejected"')
    expect(prompt).toContain('"sources": []')
    expect(prompt).toContain(SUPPORT[0].quote)
  }
  // Scope names the operative body; its later advice still gets the findings.
  expect(stagePrompt(directory, "d2.draft")).not.toContain("p111111111111")
  expect(deferredFindingItems(directory)).toHaveLength(2)
  expect(disclosableItems(directory)).toEqual(deferredFindingItems(directory))
  expect(readOpenItems(directory)).toEqual([]) // no duplicate persisted ledger
  writeLeanText(directory, "drafting-notes.md", "# Notes\n\nNo remaining work.")
  expect(finalGate(directory).status).toBe("blocked")
  writeLeanText(directory, "drafting-notes.md", `# Notes\n\n${deferredFindingItems(directory).map(row => `${row.id}: ${row.detail}`).join("\n")}`)
  expect(finalGate(directory)).toMatchObject({ status: "pass", delivery: "review_draft" })
})

test("admission is causal for evidence, scope, uncertainty and duplicate dispositions", () => {
  const directory = setup()
  expect(checkChangePlan(directory, "d1", plan())).toEqual([])
  const unsupported = plan()
  unsupported.findings![3].sources[0].quote = "Unsupported words absent from canonical text."
  expect(checkChangePlan(directory, "d1", unsupported).join(" ")).toContain("not in that document")
  const wrongScope = plan()
  wrongScope.findings![0].documents = ["d9"]
  expect(checkChangePlan(directory, "d1", wrongScope).join(" ")).toContain("planned document ids")
  const noLimitation = plan()
  noLimitation.findings![0].limitations = []
  expect(checkChangePlan(directory, "d1", noLimitation).join(" ")).toContain("evidence limitation")
  const duplicated = plan([...findings, { ...findings[0], detail: `  ${findings[0].detail.toUpperCase()} `, documents: ["d2"] }])
  expect(checkChangePlan(directory, "d1", duplicated).join(" ")).toContain("consolidate duplicate findings")
  expect(checkChangePlan(directory, "d1", plan())).toEqual([])
})

test("changed dispositions invalidate dependent consumers while leaving earlier drafting stable", () => {
  const directory = setup()
  const steps = ["d1.draft", "d2.draft", "d2.review", "d2.change-plan", "finalize", "package-review", "gate"] as const
  const before = steps.map(step => stageInputSha(directory, step))
  const amended = plan()
  amended.findings![2].reason = "Explain this as a reminder option without altering the operative period."
  expect(acceptChangePlan(directory, "d1", JSON.stringify(amended)).issues).toEqual([])
  expect(steps.map((step, index) => stageInputSha(directory, step) !== before[index])).toEqual([false, true, true, true, true, true, true])
})

test("legacy dismissal retains review navigation without being treated as evidence-backed rejection", () => {
  const directory = setup()
  writeLeanText(directory, "d1/review-independent.md", "[p111111111111] Further authority is needed before relying on enforceability.")
  expect(acceptChangePlan(directory, "d1", JSON.stringify({ shared_decisions: [], groups: [], dismissed: ["Defer the research proposition to advice."], questions: [] })).issues).toEqual([])
  const context = reviewDispositionText(directory, "d1")
  expect(context).toContain("legacy_decisions")
  expect(context).toContain("unadjudicated_review_history")
  expect(context).toContain("Further authority is needed")
  expect(context).not.toContain('"disposition": "rejected"')
  expect(completedDocumentsText(directory, "d2")).not.toContain("p111111111111")
})

test("prior-document free-text pointers are navigation while saved records and canonical quotes stay intact", () => {
  const directory = setup()
  const referenced = plan()
  referenced.findings![0].detail += " Review [p111111111111]."
  referenced.findings![0].reason += " Final location p222222222222."
  referenced.findings![0].limitations.push("Unverified for [p111111111111].")
  referenced.dismissed = ["Earlier idea at [p111111111111] was out of scope."]
  expect(acceptChangePlan(directory, "d1", JSON.stringify(referenced)).issues).toEqual([])
  const earlier = completedDocumentsText(directory, "d2")
  expect(earlier).not.toContain("p111111111111")
  expect(earlier).not.toContain("p222222222222")
  expect(earlier).toContain("earlier paragraph")
  expect(earlier).toContain(SUPPORT[0].quote)
  expect(readChangePlan(directory, "d1")).toEqual(referenced)
  expect(packageReviewPrompt(directory)).toContain("p111111111111")
})

test("coordinator applies accepted amendments while deferred findings survive apply and durable resume", async () => {
  const directory = matter()
  const model = scriptedModel({ ...happyPath, [AGENT.draft]: drafter(DRAFT, LETTER), [AGENT.changes]: changesFirst, [AGENT.review]: reviewFirst,
    [AGENT["change-plan"]]: prompt => {
      const own = /<draft_paragraphs>\n([\s\S]*?)\n<\/draft_paragraphs>/.exec(prompt)?.[1] ?? ""
      if (!own.includes("The supplier shall deliver the goods.")) return JSON.stringify(plan([]))
      return JSON.stringify({ ...plan(), groups: [{ id: "delivery", pids: [pidOf(own, "The supplier shall deliver the goods.")], goal: "Retain the supplied delivery period.", sources: SUPPORT, limitations: [] }] })
    },
  })
  const result = await runLean(directory, { sessions: model.port })
  expect(result.stop_reason).toBeNull()
  expect(result.status).toBe("completed")
  expect(readParagraphs(directory, "d1", "paragraphs-applied.json").map(row => row.text)).toContain(APPLIED)
  expect(readLeanJson(directory, "d1/apply.json")).toMatchObject({ applied: [{ mode: "code" }] })
  expect(finalGate(directory)).toMatchObject({ status: "pass", delivery: "review_draft" })
  expect(readOpenItems(directory).some(row => row.detail === findings[0].detail)).toBe(false)
  expect(deferredFindingItems(directory)).toHaveLength(2)
  const resumed = scriptedModel({})
  expect((await runLean(directory, { sessions: resumed.port })).status).toBe("completed")
  expect(resumed.prompts).toEqual([])
  findings.forEach(row => expect(stagePrompt(directory, "d2.draft")).toContain(row.detail))
  expect(stagePrompt(directory, "finalize")).toContain(deferredFindingItems(directory)[0].id)
}, 120_000)
