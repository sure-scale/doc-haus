import { afterEach, expect, test } from "bun:test"
import { readFileSync } from "node:fs"
import { CHECKLIST_JOB_AGENTS } from "./checklist-job"
import type { ChecklistCheckerInput, ChecklistProducerInput } from "./checklist-contract"
import { quoteFound } from "./quotes"
import { finalGate } from "./gate"
import { drainLeanRequests, enqueueLeanRequest, leanClarifications, pendingLeanRequests } from "./requests"
import { runLean, stageInputSha } from "./runner"
import { leanFile, readLeanJson, readLeanState, readLeanText, writeLeanState, writeLeanText } from "./state"
import { AGENT, cleanupMatters, happyPath, matter, scriptedModel } from "./test-support"

afterEach(cleanupMatters)

function inspector(status: "satisfied" | "unsatisfied" | "indeterminate", options: { second?: boolean; failSecond?: boolean } = {}) {
  return scriptedModel({ ...happyPath,
    [CHECKLIST_JOB_AGENTS.extract]: text => {
      const input = JSON.parse(text.split("SOURCE_INPUT:\n\n")[1]!) as ChecklistProducerInput
      return JSON.stringify({ items: input.source.document === "request" ? [{
        check: "The supplied delivery timing is included.", basis: "directive", source_refs: input.passages,
        qualifications: [], target: { ...input.targets[0], treatment: "operative" },
      }].flatMap(item => options.second ? [item, { ...item, check: "Payment remains due on delivery." }] : [item]) : [] })
    },
    [CHECKLIST_JOB_AGENTS.check]: text => {
      const input = JSON.parse(text.split("CHECK_INPUT:\n\n")[1]!) as ChecklistCheckerInput
      if (options.failSecond && input.item.check.startsWith("Payment")) throw new Error("Scripted read-only inspection failure")
      const findingStatus = input.item.check.startsWith("Payment") ? "satisfied" : status
      return JSON.stringify({ item_id: input.item_id, status: findingStatus, reason: `Focused result ${status}.`, evidence: input.current_evidence,
        source_refs: input.item.source_refs, inspection_scope: "Current body paragraphs.", limitation: "Approval is not inferred." })
    },
  })
}

const clarification = "The delivery condition refers to the order date in the supplied terms; inspect the existing final wording."
const preserved = ["request.md", "plan.json", "checklist/extraction.json", "d1/body.docx", "d1/body-applied.docx", "d1/redline.docx", "d1/changes.json", "d1/change-plan.json", "drafting-notes.md", "drafting-notes.docx"]

// A clarification is admitted only at a blocked checklist: here the second
// requirement's failed check job is the integrity failure that blocks it.
for (const result of ["satisfied", "unsatisfied"] as const) test(`blocked clarification reinspects existing evidence: ${result}`, async () => {
  const directory = matter()
  writeLeanText(directory, "repairs.json", JSON.stringify({ round: 5, substantive_attempts: 5, requested_at: "saved", items: [] }))
  expect((await runLean(directory, { sessions: inspector("indeterminate", { second: true, failSecond: true }).port })).status).toBe("blocked")
  const originals = preserved.map(file => readFileSync(leanFile(directory, file)))
  const early = ["discrepancy", "intake", "research", "spine", "d1.draft", "d1.review", "d1.change-plan", "d1.changes", "d1.apply", "finalize"] as const
  const keys = early.map(step => stageInputSha(directory, step))
  const before = stageInputSha(directory, "checklist-check")
  const repair = readLeanText(directory, "repairs.json")
  const request = enqueueLeanRequest(directory, { kind: "clarification", text: clarification })
  expect(leanClarifications(directory)).toEqual([])
  expect(stageInputSha(directory, "checklist-check")).toBe(before)
  drainLeanRequests(directory)
  expect(pendingLeanRequests(directory)).toEqual([`${request.id}.json`])
  // The reinspection of the adverse finding plus the re-run of the failed check.
  const review = inspector(result, { second: true })
  expect((await runLean(directory, { sessions: review.port })).status).toBe("completed")
  expect(review.count(CHECKLIST_JOB_AGENTS.check)).toBe(2)
  expect(review.count(CHECKLIST_JOB_AGENTS.extract)).toBe(0)
  for (const agent of [AGENT.draft, AGENT.changes, AGENT.finalize]) expect(review.count(agent)).toBe(0)
  expect(review.count(AGENT["package-review"])).toBe(1)
  const reinspected = review.prompts.find(row => row.agent === CHECKLIST_JOB_AGENTS.check && row.text.includes("The supplied delivery timing is included."))?.text
  expect(reinspected).toContain(clarification)
  expect(reinspected).toContain("cannot implement it")
  expect(readLeanJson(directory, "checklist/checks.json")).toMatchObject({ checks: [{ finding: { status: result }, reinspection: { original_finding: { status: "indeterminate" }, status: "succeeded" } }, { status: "succeeded", finding: { status: "satisfied" } }] })
  expect(readLeanText(directory, "updated-notes.md")).toBeUndefined()
  expect(quoteFound(directory, { source_id: "request", quote: clarification })).toBe(true)
  expect(readLeanText(directory, "repairs.json")).toBe(repair)
  expect(preserved.map(file => readFileSync(leanFile(directory, file)))).toEqual(originals)
  expect(early.map(step => stageInputSha(directory, step))).toEqual(keys)
  // An unsatisfied result is disclosed with the review draft: the substantive
  // allowance is used, so no repair round starts and repairs.json is unchanged.
  expect(finalGate(directory)).toMatchObject({ status: "pass", delivery: result === "satisfied" ? "execution_ready" : "review_draft_with_findings" })
  if (result === "unsatisfied") {
    enqueueLeanRequest(directory, { kind: "clarification", text: clarification })
    const repeated = inspector("satisfied", { second: true })
    expect((await runLean(directory, { sessions: repeated.port })).status).toBe("completed")
    expect(repeated.prompts).toHaveLength(0)
  }
}, 120_000)

test("a queued clarification rejects a changed stop point without becoming drafting instructions", () => {
  const directory = matter()
  const request = enqueueLeanRequest(directory, { kind: "clarification", text: clarification })
  drainLeanRequests(directory, { ...readLeanState(directory)!, status: "completed" })
  expect(readLeanJson(directory, `requests/done/${request.id}.json`)).toMatchObject({ outcome: { status: "rejected" } })
  expect(leanClarifications(directory)).toEqual([])
  expect(readLeanText(directory, "updated-notes.md")).toBeUndefined()
})


test("a failed clarification inspection cannot keep a stale satisfied finding as a release", async () => {
  const directory = matter()
  expect((await runLean(directory, { sessions: inspector("indeterminate", { second: true }).port })).status).toBe("completed")
  // Settled blocked at the checklist, as before adverse findings were disclosed.
  const state = readLeanState(directory)!
  writeLeanState(directory, { ...state, status: "blocked", stop_reason: "Checklist handoff blocked. See checklist/report.md.",
    stages: { ...state.stages, "checklist-check": { ...state.stages["checklist-check"]!, status: "blocked" } } })
  enqueueLeanRequest(directory, { kind: "clarification", text: clarification })
  const review = inspector("satisfied", { second: true, failSecond: true })
  expect((await runLean(directory, { sessions: review.port })).status).toBe("blocked")
  expect(review.count(CHECKLIST_JOB_AGENTS.check)).toBe(2)
  expect(finalGate(directory).status).toBe("blocked")
  const repeated = inspector("satisfied", { second: true })
  expect((await runLean(directory, { sessions: repeated.port })).status).toBe("blocked")
  expect(repeated.prompts).toHaveLength(0)
}, 120_000)
