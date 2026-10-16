import { afterEach, expect, test } from "bun:test"
import { checkPackageReview, packageReviewPrompt } from "./package-review"
import { writeLeanPlan } from "./plan"
import { stagePrompt } from "./prompts"
import { writeLeanText } from "./state"
import { cleanupMatters, matter } from "./test-support"

afterEach(cleanupMatters)

test("a material semantic finding cannot coexist with a package pass", () => {
  const directory = matter()
  expect(checkPackageReview(directory, { status: "pass", delivery: [], findings: [], limitations: ["Awaiting a disclosed decision."] })).toEqual([])
  expect(checkPackageReview(directory, { status: "pass", findings: [], limitations: [] }).join(" ")).toContain("delivery as an array")
  expect(checkPackageReview(directory, { status: "pass", delivery: [], findings: [{ detail: "The memo contradicts the operative body." }], limitations: [] }).join(" ")).toContain("pass cannot")
  expect(checkPackageReview(directory, { status: "changes-required", delivery: [], findings: [{ target: "memo", detail: "The memo contradicts the operative body." }], limitations: [] })).toEqual([])
  expect(checkPackageReview(directory, { status: "changes-required", delivery: [], findings: [], limitations: [] }).join(" ")).toContain("material correction")
  expect(checkPackageReview(directory, { status: "changes-required", delivery: [], findings: [{ target: "body", detail: "Requires an affected body" }], limitations: [] }).join(" ")).toContain("affected document id")
})

test("changes always retain request and legal research without a later update", () => {
  const directory = matter()
  const prompt = stagePrompt(directory, "d1.changes")
  expect(prompt).toContain("<request>")
  expect(prompt).toContain("<legal_research>")
  writeLeanText(directory, "d1/review.md", "First independent review")
  writeLeanText(directory, "d1/review-independent.md", "Second independent review")
  expect(stagePrompt(directory, "d1.changes")).toContain("review:\nFirst independent review\n\nreview-independent:\nSecond independent review")
})

test("review sessions receive different substantive focus without another reviewer’s findings", () => {
  const directory = matter()
  writeLeanText(directory, "d1/review.md", "Prior review content must stay independent")
  expect(stagePrompt(directory, "d1.review")).toContain("Function, anatomy and package completeness")
  const prompt = stagePrompt(directory, "d1.review-independent")
  expect(prompt).toContain("Source and legal-commercial fidelity")
  expect(prompt).not.toContain("Prior review content must stay independent")
})

test("final semantic review reads memo with live source access", () => {
  const directory = matter()
  writeLeanText(directory, "drafting-notes.md", "# Notes\nA material description of the final wording.")
  const prompt = packageReviewPrompt(directory)
  expect(prompt).toContain("A material description of the final wording.")
  expect(prompt).toContain("search-document")
  expect(prompt).toContain("after revisions")
})

test("memo and semantic review inspect the final body rather than initial or proposed wording", () => {
  const directory = matter()
  writeLeanPlan(directory, [{ title: "Supply note", purpose: "Record agreed supply terms", parts: [{ heading: "Delivery", scope: "Delivery terms" }] }])
  writeLeanText(directory, "d1/paragraphs.json", JSON.stringify([{ pid: "p111111111111", sha256: "initial", text: "Initial wording" }]))
  writeLeanText(directory, "d1/paragraphs-applied.json", JSON.stringify([{ pid: "p222222222222", sha256: "final", text: "Actual final operative wording" }]))
  expect(stagePrompt(directory, "finalize")).toContain("Actual final operative wording")
  expect(stagePrompt(directory, "finalize")).not.toContain("Initial wording")
  expect(packageReviewPrompt(directory)).toContain("Actual final operative wording")
  expect(checkPackageReview(directory, { status: "changes-required", delivery: [], findings: [{ document: "d1", pid: "p111111111111", detail: "Wrong baseline" }], limitations: [] }).join(" ")).toContain("named final document")
})
