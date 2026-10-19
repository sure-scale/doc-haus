import { afterEach, expect, test } from "bun:test"
import { checkPackageReview, packageReviewPrompt, readPackageReview } from "./package-review"
import { writeLeanPlan } from "./plan"
import { stagePrompt } from "./prompts"
import { acceptStageReply } from "./stages"
import { readLeanText, writeLeanText } from "./state"
import { cleanupMatters, matter } from "./test-support"

afterEach(cleanupMatters)

// The lawyer's words name no file and give no title: a row that gives a
// planned title plus .docx names nothing the lawyer asked for.
const REQUEST = "Prepare a supply note from the supply terms, together with a short covering letter on the points to resolve before signing."
const UNGROUNDED = [
  { document: "d1", role: "body", file: "Supply note.docx", instruction_quote: "Prepare a supply note from the supply terms" },
  { document: "d2", role: "body", file: "Covering letter.docx", instruction_quote: "together with a short covering letter on the points to resolve before signing" },
]
const withheld = (file: string) => `${file}: no naming instruction in the lawyer's words; the planned name stands`
function twoDocuments(request = REQUEST) {
  const directory = matter()
  writeLeanText(directory, "request.md", request)
  writeLeanPlan(directory, [{ title: "Supply note", purpose: "Operative terms", parts: [{ heading: "Terms", scope: "Preserve" }] },
    { title: "Covering letter", purpose: "Points to resolve", parts: [{ heading: "Points", scope: "List" }] }])
  return directory
}
const review = (delivery: unknown[], extra: Record<string, unknown> = {}) =>
  JSON.stringify({ status: "pass", findings: [], limitations: ["A disclosed gap."], delivery, ...extra })

test("an ungrounded delivery rename is retried before the last attempt and withheld on it, the planned names standing", async () => {
  const directory = twoDocuments()
  const retried = await acceptStageReply(directory, "package-review", review(UNGROUNDED))
  expect(retried.outputs).toEqual([])
  expect(retried.issues).toHaveLength(2)
  for (const issue of retried.issues) expect(issue).toEndWith("When the lawyer's words give that output no file name or explicit title, return no delivery row for it.")
  expect(readLeanText(directory, "package-review.json")).toBeUndefined()
  const last = await acceptStageReply(directory, "package-review", review(UNGROUNDED), { final: true })
  expect(last).toEqual({ issues: [], outputs: ["package-review.json"], openItems: [] })
  expect(readPackageReview(directory)).toEqual({ status: "pass", findings: [], delivery: [],
    limitations: ["A disclosed gap.", withheld("Supply note.docx"), withheld("Covering letter.docx")] })
})

test("on the last attempt a grounded delivery row still applies beside a withheld one", async () => {
  const directory = twoDocuments(`${REQUEST} Save the covering letter as Points to resolve.docx.`)
  const grounded = { document: "d2", role: "body", file: "Points to resolve.docx", instruction_quote: "Save the covering letter as Points to resolve.docx" } as const
  expect((await acceptStageReply(directory, "package-review", review([UNGROUNDED[0], grounded]))).issues).toHaveLength(1)
  expect((await acceptStageReply(directory, "package-review", review([UNGROUNDED[0], grounded]), { final: true })).issues).toEqual([])
  expect(readPackageReview(directory)).toEqual({ status: "pass", findings: [], delivery: [grounded], limitations: ["A disclosed gap.", withheld("Supply note.docx")] })
  expect((await acceptStageReply(directory, "package-review", review([grounded]), { final: true })).issues).toEqual([])
  expect(readPackageReview(directory)).toEqual({ status: "pass", findings: [], delivery: [grounded], limitations: ["A disclosed gap."] })
})

test("grounded delivery rows that still conflict on the last attempt leave the planned names, never the review's findings", async () => {
  const directory = twoDocuments(`${REQUEST} Save the supply note as Supply terms.docx.`)
  const grounded = { document: "d1", role: "body", file: "Supply terms.docx", instruction_quote: "Save the supply note as Supply terms.docx" }
  // Two names for one document are a conflict across rows, not an ungrounded row.
  const twice = [grounded, { ...grounded, role: "redline" }]
  expect((await acceptStageReply(directory, "package-review", review(twice))).issues.join(" ")).toContain("one named primary artifact")
  const findings = [{ target: "memo", detail: "The notes misdescribe the delivery period." }]
  expect((await acceptStageReply(directory, "package-review", review(twice, { status: "changes-required", findings }), { final: true })).issues).toEqual([])
  const saved = readPackageReview(directory)
  expect(saved).toMatchObject({ status: "changes-required", findings, delivery: [] })
  expect(saved?.limitations).toHaveLength(3)
  expect(saved?.limitations.slice(1)).toEqual(Array(2).fill("Supply terms.docx: conflicts with another delivery selection; the planned name stands"))
})

test("the last attempt never drops a finding or changes the status to pass the check", async () => {
  const directory = twoDocuments()
  const findings = [{ target: "memo", detail: "The notes misdescribe the delivery period." }]
  const last = await acceptStageReply(directory, "package-review", review(UNGROUNDED, { findings }), { final: true })
  expect(last.issues).toEqual(["A pass cannot contain unresolved correction findings; use changes-required."])
  expect(readLeanText(directory, "package-review.json")).toBeUndefined()
  expect((await acceptStageReply(directory, "package-review", review(UNGROUNDED, { status: "changes-required", findings }), { final: true })).issues).toEqual([])
  expect(readPackageReview(directory)).toMatchObject({ status: "changes-required", findings, delivery: [] })
})

test("a missing delivery list is retried before the last attempt and read as no rename on it", async () => {
  const directory = twoDocuments()
  const missing = (extra: Record<string, unknown> = {}) => JSON.stringify({ status: "pass", findings: [], limitations: ["A disclosed gap."], ...extra })
  const noted = ["A disclosed gap.", "delivery field missing; planned names stand"]
  expect((await acceptStageReply(directory, "package-review", missing())).issues.join(" ")).toContain("delivery as an array")
  expect(readLeanText(directory, "package-review.json")).toBeUndefined()
  expect((await acceptStageReply(directory, "package-review", missing(), { final: true })).issues).toEqual([])
  expect(readPackageReview(directory)).toEqual({ status: "pass", findings: [], delivery: [], limitations: noted })
  // A delivery that is not a list is no rename either; status and findings are kept.
  const findings = [{ target: "memo" as const, detail: "The notes misdescribe the delivery period." }]
  expect((await acceptStageReply(directory, "package-review", missing({ status: "changes-required", findings, delivery: "none" }), { final: true })).issues).toEqual([])
  expect(readPackageReview(directory)).toEqual({ status: "changes-required", findings, delivery: [], limitations: noted })
  // Status and findings that fail the check still fail the last attempt.
  expect((await acceptStageReply(directory, "package-review", missing({ findings }), { final: true })).issues).toEqual(["A pass cannot contain unresolved correction findings; use changes-required."])
  expect((await acceptStageReply(directory, "package-review", missing({ limitations: "none" }), { final: true })).issues).toEqual(["Use status pass or changes-required, findings array and string limitations array."])
})

test("a prose preamble before one package review object does not cost an attempt", async () => {
  const directory = twoDocuments()
  const accepted = await acceptStageReply(directory, "package-review", `Based on my review of the package, here is my assessment.\n\n${review([])}`)
  expect(accepted.issues).toEqual([])
  expect(readPackageReview(directory)).toEqual({ status: "pass", findings: [], limitations: ["A disclosed gap."], delivery: [] })
  expect((await acceptStageReply(directory, "package-review", "No JSON here.")).issues).toEqual(["Reply with the package review JSON only."])
})

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
