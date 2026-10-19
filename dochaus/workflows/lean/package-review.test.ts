import { afterEach, expect, test } from "bun:test"
import { deliveredNotes, finalGate } from "./gate"
import { checkPackageReview, packageReviewPrompt, readPackageReview, withheldDeliveryItems, type WithheldDelivery } from "./package-review"
import { readLeanPlan, writeLeanPlan } from "./plan"
import { LEAN_REQUESTED_FILES_SCHEMA, requestedFilesItems, requestedFilesPlanned, writeRequestedFiles } from "./requested-files"
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
const withheld = (file: string, reason = "no naming instruction in the lawyer's words") => `${file}: ${reason}; the planned name stands`
const NOT_QUOTED = { cause: "not_quoted", reason: "no naming instruction in the lawyer's words", asked: false } as const
function twoDocuments(request = REQUEST, options: { basis?: string; memo?: { name: string; file: string }; file?: string } = {}) {
  const directory = matter()
  writeLeanText(directory, "request.md", request)
  writeLeanPlan(directory, [{ title: "Supply note", purpose: "Operative terms", parts: [{ heading: "Terms", scope: "Preserve" }] },
    { title: "Covering letter", purpose: "Points to resolve", parts: [{ heading: "Points", scope: "List" }], ...(options.file ? { file: options.file } : {}) }],
  options.basis, undefined, options.memo)
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
    limitations: ["A disclosed gap.", withheld("Supply note.docx"), withheld("Covering letter.docx")],
    withheld_delivery: [{ document: "d1", role: "body", file: "Supply note.docx", ...NOT_QUOTED }, { document: "d2", role: "body", file: "Covering letter.docx", ...NOT_QUOTED }] })
  // A plan title with .docx, which the lawyer's words never give as a name, is no item for the lawyer.
  expect(withheldDeliveryItems(directory)).toEqual([])
  expect(finalGate(directory).items).toEqual([])
})

test("on the last attempt a grounded delivery row still applies beside a withheld one", async () => {
  const directory = twoDocuments(`${REQUEST} Save the covering letter as Points to resolve.docx.`)
  const grounded = { document: "d2", role: "body", file: "Points to resolve.docx", instruction_quote: "Save the covering letter as Points to resolve.docx" } as const
  expect((await acceptStageReply(directory, "package-review", review([UNGROUNDED[0], grounded]))).issues).toHaveLength(1)
  expect((await acceptStageReply(directory, "package-review", review([UNGROUNDED[0], grounded]), { final: true })).issues).toEqual([])
  expect(readPackageReview(directory)).toEqual({ status: "pass", findings: [], delivery: [grounded], limitations: ["A disclosed gap.", withheld("Supply note.docx")],
    withheld_delivery: [{ document: "d1", role: "body", file: "Supply note.docx", ...NOT_QUOTED }] })
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
  const conflict = { file: "Supply terms.docx", cause: "conflict", reason: "conflicts with another delivery selection", asked: true } as const
  expect(saved?.withheld_delivery).toEqual([{ document: "d1", role: "body", ...conflict }, { document: "d1", role: "redline", ...conflict }])
  // The lawyer asked for the name, so it is disclosed once, never blocking.
  expect(withheldDeliveryItems(directory).map(item => [item.detail, item.status])).toEqual([[`You asked for "Supply terms.docx"; it could not be applied to "Supply note" because it conflicts with another file name chosen for this package, so that document keeps the file name "Supply note.docx".`, "disclosed"]])
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
  const notList = ["A disclosed gap.", "delivery was not a list; planned names stand"]
  for (const delivery of ["none", null]) {
    expect((await acceptStageReply(directory, "package-review", missing({ status: "changes-required", findings, delivery }), { final: true })).issues).toEqual([])
    expect(readPackageReview(directory)).toEqual({ status: "changes-required", findings, delivery: [], limitations: notList })
  }
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

test("a reply of two objects is retried, on the last attempt too, never read as its last object", async () => {
  const directory = twoDocuments()
  const later = review([], { status: "changes-required", findings: [{ target: "memo", detail: "The notes misdescribe the delivery period." }] })
  for (const final of [false, true]) {
    expect((await acceptStageReply(directory, "package-review", `${review([])}\n${later}`, { final })).issues).toEqual(["Reply with the package review JSON only."])
    expect(readLeanText(directory, "package-review.json")).toBeUndefined()
  }
})

test("a name the lawyer asked for, withheld on the last attempt for a wrong document id, is disclosed without blocking the gate", async () => {
  const directory = twoDocuments(`${REQUEST} Save the supply note as Supply terms.docx.`)
  writeLeanText(directory, "drafting-notes.md", "# Notes\nThe supply note records the agreed terms.")
  const row = { document: "d3", role: "body", file: "Supply terms.docx", instruction_quote: "Save the supply note as Supply terms.docx" }
  expect((await acceptStageReply(directory, "package-review", review([row]))).issues).toEqual(["Delivery selection must name an existing final document id."])
  // A reply's own record of withheld rows is never kept.
  const claimed = { withheld_delivery: [{ file: "Invented.docx", cause: "document", reason: "names no planned document", asked: true }] }
  expect((await acceptStageReply(directory, "package-review", review([row], claimed), { final: true })).issues).toEqual([])
  expect(readPackageReview(directory)).toEqual({ status: "pass", findings: [], delivery: [], limitations: ["A disclosed gap.", withheld("Supply terms.docx", "names no planned document")],
    withheld_delivery: [{ document: "d3", role: "body", file: "Supply terms.docx", cause: "document", reason: "names no planned document", asked: true }] })
  const sentence = `You asked for "Supply terms.docx"; it could not be matched to a document in this package, so every file keeps its planned name.`
  const gate = finalGate(directory)
  expect(gate.blocking).toEqual([])
  expect(gate.status).toBe("pass")
  expect(gate.delivery).toBe("review_draft")
  expect(gate.disclosed.map(item => [item.stage, item.detail])).toEqual([["package-review", sentence]])
  expect(deliveredNotes(directory, gate)).toBe(["# File names not applied", "", "These requested file names could not be applied. Each is disclosed with this review draft.", "",
    `- ${sentence}`, "", "# Notes", "The supply note records the agreed terms."].join("\n"))
})

test("a duplicate grounded delivery row is kept once on the last attempt, with nothing withheld", async () => {
  const directory = twoDocuments(`${REQUEST} Save the supply note as Supply terms.docx.`)
  const grounded = { document: "d1", role: "body", file: "Supply terms.docx", instruction_quote: "Save the supply note as Supply terms.docx" } as const
  expect((await acceptStageReply(directory, "package-review", review([grounded, { ...grounded }]))).issues.join(" ")).toContain("one named primary artifact")
  expect((await acceptStageReply(directory, "package-review", review([grounded, { ...grounded }]), { final: true })).issues).toEqual([])
  expect(readPackageReview(directory)).toEqual({ status: "pass", findings: [], delivery: [grounded], limitations: ["A disclosed gap."] })
  expect(withheldDeliveryItems(directory)).toEqual([])
})

test("on the last attempt a grounded row is kept beside a grounded row that takes the memo's name", async () => {
  const directory = twoDocuments(`${REQUEST} Save the covering letter as Points to resolve.docx. Save the supply note as Points memo.docx.`,
    { memo: { name: "points memo", file: "Points memo.docx" } })
  const kept = { document: "d2", role: "body", file: "Points to resolve.docx", instruction_quote: "Save the covering letter as Points to resolve.docx" } as const
  const clash = { document: "d1", role: "body", file: "Points memo.docx", instruction_quote: "Save the supply note as Points memo.docx" }
  expect((await acceptStageReply(directory, "package-review", review([kept, clash]))).issues.length).toBeGreaterThan(0)
  expect((await acceptStageReply(directory, "package-review", review([kept, clash]), { final: true })).issues).toEqual([])
  expect(readPackageReview(directory)).toEqual({ status: "pass", findings: [], delivery: [kept], limitations: ["A disclosed gap.", withheld("Points memo.docx", "the notes already have that file name")],
    withheld_delivery: [{ document: "d1", role: "body", file: "Points memo.docx", cause: "notes_name", reason: "the notes already have that file name", asked: true }] })
  expect(withheldDeliveryItems(directory).map(item => item.detail)).toEqual([`You asked for "Points memo.docx"; it could not be applied to "Supply note" because the memo already has that name, so that document keeps the file name "Supply note.docx".`])
})

test("a rename the lawyer's words never give stays a limitation, with no item for the lawyer", async () => {
  const directory = twoDocuments()
  const invented = { document: "d1", role: "body", file: "Supply note final.docx", instruction_quote: "Prepare a supply note from the supply terms" }
  expect((await acceptStageReply(directory, "package-review", review([invented]), { final: true })).issues).toEqual([])
  expect(readPackageReview(directory)?.withheld_delivery).toEqual([{ document: "d1", role: "body", file: "Supply note final.docx", ...NOT_QUOTED }])
  expect(withheldDeliveryItems(directory)).toEqual([])
  expect(finalGate(directory).items).toEqual([])
})

test("each withheld delivery row records its own plain reason", async () => {
  const save = "Save the supply note as Supply terms.docx"
  const cases: [string, unknown[], WithheldDelivery[], string[], { file?: string }][] = [
    ["wrong document id", [{ document: "d9", role: "body", file: "Supply terms.docx", instruction_quote: save }],
      [{ document: "d9", role: "body", file: "Supply terms.docx", cause: "document", reason: "names no planned document", asked: true }],
      [`You asked for "Supply terms.docx"; it could not be matched to a document in this package, so every file keeps its planned name.`], {}],
    ["role", [{ document: "d1", role: "clean", file: "Supply terms.docx", instruction_quote: save }],
      [{ document: "d1", role: "clean", file: "Supply terms.docx", cause: "role", reason: "names no role the run delivers", asked: true }],
      [`You asked for "Supply terms.docx"; it could not be applied to "Supply note" because it does not say which file of that document it names, so that document keeps the file name "Supply note.docx".`], {}],
    ["notes row naming a document", [{ document: "d1", role: "drafting_notes", file: "Supply terms.docx", instruction_quote: save }],
      [{ document: "d1", role: "drafting_notes", file: "Supply terms.docx", cause: "notes_document", reason: "a selection of the notes that names a document", asked: true }],
      [`You asked for "Supply terms.docx"; it could not be matched to a document in this package, so every file keeps its planned name.`], {}],
    ["name shape", [{ document: "d1", role: "body", file: "Supply: terms.docx", instruction_quote: "Save the supply note as Supply: terms.docx" }],
      [{ document: "d1", role: "body", file: "Supply: terms.docx", cause: "file", reason: "not a plain Word file name", asked: true }],
      [`You asked for "Supply: terms.docx"; it could not be applied to "Supply note" because it is not a usable Word file name, so that document keeps the file name "Supply note.docx".`], {}],
    ["no naming instruction", [{ document: "d1", role: "body", file: "Supply terms.docx", instruction_quote: "Prepare a supply note from the supply terms" }],
      [{ document: "d1", role: "body", file: "Supply terms.docx", cause: "not_quoted", reason: "no naming instruction in the lawyer's words", asked: true }],
      [`You asked for "Supply terms.docx"; it could not be applied to "Supply note", so that document keeps the file name "Supply note.docx".`], {}],
    ["format", [{ document: "d1", role: "body", file: "Word.docx", instruction_quote: "Deliver both in Word" }],
      [{ document: "d1", role: "body", file: "Word.docx", cause: "format", reason: "names a format, not a file", asked: false }], [], {}],
    ["not a row", ["Supply terms.docx"], [{ cause: "selection", reason: "not a complete delivery selection", asked: false }], [], {}],
    ["conflict between rows", [{ document: "d1", role: "body", file: "Supply terms.docx", instruction_quote: save },
      { document: "d2", role: "body", file: "Supply terms.docx", instruction_quote: "Save the covering letter as Supply terms.docx" }],
    [{ document: "d1", role: "body", file: "Supply terms.docx", cause: "conflict", reason: "conflicts with another delivery selection", asked: true },
      { document: "d2", role: "body", file: "Supply terms.docx", cause: "conflict", reason: "conflicts with another delivery selection", asked: true }],
    [`You asked for "Supply terms.docx"; it could not be applied to "Supply note" because it conflicts with another file name chosen for this package, so that document keeps the file name "Supply note.docx".`,
      `You asked for "Supply terms.docx"; it could not be applied to "Covering letter" because it conflicts with another file name chosen for this package, so that document keeps the file name "Covering letter.docx".`], {}],
    ["another planned file's name", [{ document: "d1", role: "body", file: "Letter terms.docx", instruction_quote: "Save the supply note as Letter terms.docx" }],
      [{ document: "d1", role: "body", file: "Letter terms.docx", cause: "planned_name", reason: "another planned file already has that name", asked: true }],
      [`You asked for "Letter terms.docx"; it could not be applied to "Supply note" because another file in this package already has that name, so that document keeps the file name "Supply note.docx".`], { file: "Letter terms.docx" }],
  ]
  for (const [label, rows, records, details, options] of cases) {
    const directory = twoDocuments(`${REQUEST} ${save}. Save the covering letter as Supply terms.docx. Save the supply note as Supply: terms.docx. Save the supply note as Letter terms.docx. Deliver both in Word.`, options)
    expect((await acceptStageReply(directory, "package-review", review(rows), { final: true })).issues, label).toEqual([])
    const saved = readPackageReview(directory)
    expect(saved?.delivery, label).toEqual([])
    expect(saved?.withheld_delivery, label).toEqual(records)
    expect(saved?.limitations.slice(1), label).toEqual(records.map(row => withheld(row.file ?? "A delivery selection without a file name", row.reason)))
    expect(withheldDeliveryItems(directory).map(item => item.detail), label).toEqual(details)
  }
})

test("a withheld name a requested-files item already raises is not raised again", async () => {
  const directory = twoDocuments(`${REQUEST} Save the supply note as Supply terms.docx.`, { basis: "basis-1" })
  const plan = readLeanPlan(directory)!
  writeRequestedFiles(directory, { schema: LEAN_REQUESTED_FILES_SCHEMA, basis: "basis-1", planned: requestedFilesPlanned(plan.documents, plan.memo ?? undefined), stated: true,
    files: [{ holds: ["d1"], name: "Supply terms.docx", quote: "Save the supply note as Supply terms.docx" }] })
  expect(requestedFilesItems(directory)).toHaveLength(1)
  const row = { document: "d3", role: "body", file: "Supply terms.docx", instruction_quote: "Save the supply note as Supply terms.docx" }
  expect((await acceptStageReply(directory, "package-review", review([row]), { final: true })).issues).toEqual([])
  expect(readPackageReview(directory)?.withheld_delivery).toHaveLength(1)
  expect(withheldDeliveryItems(directory)).toEqual([])
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
