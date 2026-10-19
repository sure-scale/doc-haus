import { afterEach, expect, test } from "bun:test"
import { renderMarkdownToDocx } from "../../lib/render"
import { DELIVERED_NOTES } from "./deliverables"
import { deliveredNotes, finalGate, notesCoverageReport } from "./gate"
import { checkPackageReview, packageReviewPrompt, readPackageReview, withheldDeliveryItems, type WithheldDelivery } from "./package-review"
import { readLeanPlan, writeLeanPlan } from "./plan"
import { LEAN_REQUESTED_FILES_SCHEMA, requestedFilesItems, requestedFilesPlanned, writeRequestedFiles } from "./requested-files"
import { stagePrompt } from "./prompts"
import { acceptStageReply, stageOutputValid } from "./stages"
import { readLeanText, writeLeanBytes, writeLeanText } from "./state"
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
  const conflict = { file: "Supply terms.docx", cause: "conflict", reason: "conflicts with another delivery selection", asked: true, asked_file: "Supply terms.docx", asked_spelling: "Supply terms.docx" } as const
  expect(saved?.withheld_delivery).toEqual([{ document: "d1", role: "body", ...conflict }, { document: "d1", role: "redline", ...conflict }])
  // The lawyer's instructions give the name, so it is disclosed once, never blocking.
  expect(withheldDeliveryItems(directory).map(item => [item.detail, item.status])).toEqual([[`Your instructions mention the name "Supply terms.docx", but it was not applied: "Supply note" keeps its planned file name because the name conflicts with another file name chosen for this package.`, "disclosed"]])
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
    withheld_delivery: [{ document: "d3", role: "body", file: "Supply terms.docx", cause: "document", reason: "names no planned document", asked: true, asked_file: "Supply terms.docx", asked_spelling: "Supply terms.docx" }] })
  const sentence = `Your instructions mention the name "Supply terms.docx", but it was not applied: the name could not be matched to a document in this package.`
  const gate = finalGate(directory)
  expect(gate.blocking).toEqual([])
  expect(gate.status).toBe("pass")
  expect(gate.delivery).toBe("review_draft")
  expect(gate.disclosed.map(item => [item.stage, item.detail])).toEqual([["package-review", sentence]])
  expect(deliveredNotes(directory, gate)).toBe(["# File names not applied", "", "These names in your instructions could not be applied. Each is disclosed with this review draft.", "",
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
    withheld_delivery: [{ document: "d1", role: "body", file: "Points memo.docx", cause: "notes_name", reason: "the notes already have that file name", asked: true, asked_file: "Points memo.docx", asked_spelling: "Points memo.docx" }] })
  expect(withheldDeliveryItems(directory).map(item => item.detail)).toEqual([`Your instructions mention the name "Points memo.docx", but it was not applied: "Supply note" keeps its planned file name because the memo already has the name.`])
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
      [{ document: "d9", role: "body", file: "Supply terms.docx", cause: "document", reason: "names no planned document", asked: true, asked_file: "Supply terms.docx", asked_spelling: "Supply terms.docx" }],
      [`Your instructions mention the name "Supply terms.docx", but it was not applied: the name could not be matched to a document in this package.`], {}],
    ["role", [{ document: "d1", role: "clean", file: "Supply terms.docx", instruction_quote: save }],
      [{ document: "d1", role: "clean", file: "Supply terms.docx", cause: "role", reason: "names no role the run delivers", asked: true, asked_file: "Supply terms.docx", asked_spelling: "Supply terms.docx" }],
      [`Your instructions mention the name "Supply terms.docx", but it was not applied: "Supply note" keeps its planned file name because it was not settled which file takes the name.`], {}],
    ["notes row naming a document", [{ document: "d1", role: "drafting_notes", file: "Supply terms.docx", instruction_quote: save }],
      [{ document: "d1", role: "drafting_notes", file: "Supply terms.docx", cause: "notes_document", reason: "a selection of the notes that names a document", asked: true, asked_file: "Supply terms.docx", asked_spelling: "Supply terms.docx" }],
      [`Your instructions mention the name "Supply terms.docx", but it was not applied: the name could not be matched to a document in this package.`], {}],
    ["name shape", [{ document: "d1", role: "body", file: "Supply: terms.docx", instruction_quote: "Save the supply note as Supply: terms.docx" }],
      [{ document: "d1", role: "body", file: "Supply: terms.docx", cause: "file", reason: "not a plain Word file name", asked: true, asked_file: "Supply: terms.docx", asked_spelling: "Supply: terms.docx" }],
      [`Your instructions mention the name "Supply: terms.docx", but it was not applied: "Supply note" keeps its planned file name because the name is not a usable Word file name.`], {}],
    // The name is in the lawyer's words, but not in the instruction the row quotes.
    ["no naming instruction", [{ document: "d1", role: "body", file: "Supply terms.docx", instruction_quote: "Prepare a supply note from the supply terms" }],
      [{ document: "d1", role: "body", file: "Supply terms.docx", cause: "not_quoted", reason: "no naming instruction in the lawyer's words", asked: false }],
      [], {}],
    ["format", [{ document: "d1", role: "body", file: "Word.docx", instruction_quote: "Deliver both in Word" }],
      [{ document: "d1", role: "body", file: "Word.docx", cause: "format", reason: "names a format, not a file", asked: false }], [], {}],
    ["not a row", ["Supply terms.docx"], [{ cause: "selection", reason: "not a complete delivery selection", asked: false }], [], {}],
    ["conflict between rows", [{ document: "d1", role: "body", file: "Supply terms.docx", instruction_quote: save },
      { document: "d2", role: "body", file: "Supply terms.docx", instruction_quote: "Save the covering letter as Supply terms.docx" }],
    [{ document: "d1", role: "body", file: "Supply terms.docx", cause: "conflict", reason: "conflicts with another delivery selection", asked: true, asked_file: "Supply terms.docx", asked_spelling: "Supply terms.docx" },
      { document: "d2", role: "body", file: "Supply terms.docx", cause: "conflict", reason: "conflicts with another delivery selection", asked: true, asked_file: "Supply terms.docx", asked_spelling: "Supply terms.docx" }],
    // One item for the name lists every file it was not applied to.
    [`Your instructions mention the name "Supply terms.docx", but it was not applied: "Supply note" and "Covering letter" keep their planned file names because the name conflicts with another file name chosen for this package.`], {}],
    ["another planned file's name", [{ document: "d1", role: "body", file: "Letter terms.docx", instruction_quote: "Save the supply note as Letter terms.docx" }],
      [{ document: "d1", role: "body", file: "Letter terms.docx", cause: "planned_name", reason: "another planned file already has that name", asked: true, asked_file: "Letter terms.docx", asked_spelling: "Letter terms.docx" }],
      [`Your instructions mention the name "Letter terms.docx", but it was not applied: "Supply note" keeps its planned file name because another file in this package already has the name.`], { file: "Letter terms.docx" }],
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

const SAVE = "Save the supply note as Supply terms.docx"
const GROUNDED = { document: "d1", role: "body", file: "Supply terms.docx", instruction_quote: SAVE } as const
const asked = (file: string, cause: WithheldDelivery["cause"], reason: string, document: string | null = "d1", role = "body", spelled = file, written = spelled): WithheldDelivery =>
  ({ document, role, file, cause, reason, asked: true, asked_file: spelled, asked_spelling: written })

test("a name the lawyer did not ask for is kept as a record and never shown to the lawyer, whatever refused it", async () => {
  const cases: [string, Record<string, unknown>, WithheldDelivery][] = [
    ["a path", { document: "d1", role: "body", file: ".dochaus/lean/d1/final.docx", instruction_quote: "Prepare a supply note from the supply terms" },
      { document: "d1", role: "body", file: ".dochaus/lean/d1/final.docx", cause: "file", reason: "not a plain Word file name", asked: false }],
    ["an internal id", { document: "d1", role: "clean", file: "d2 letter.docx", instruction_quote: "Prepare a supply note from the supply terms" },
      { document: "d1", role: "clean", file: "d2 letter.docx", cause: "role", reason: "names no role the run delivers", asked: false }],
    ["a stage", { document: "d1", role: "clean", file: "package-review d1.docx", instruction_quote: "Prepare a supply note" },
      { document: "d1", role: "clean", file: "package-review d1.docx", cause: "role", reason: "names no role the run delivers", asked: false }],
    ["a stage on no document", { document: "d9", role: "body", file: "package-review d1.docx", instruction_quote: "Prepare a supply note" },
      { document: "d9", role: "body", file: "package-review d1.docx", cause: "document", reason: "names no planned document", asked: false }],
  ]
  for (const [label, row, record] of cases) {
    const directory = twoDocuments()
    writeLeanText(directory, "drafting-notes.md", "# Notes\nThe supply note records the agreed terms.")
    expect((await acceptStageReply(directory, "package-review", review([row]), { final: true })).issues, label).toEqual([])
    expect(readPackageReview(directory)?.withheld_delivery, label).toEqual([record])
    expect(withheldDeliveryItems(directory), label).toEqual([])
    const gate = finalGate(directory)
    expect(gate.items, label).toEqual([])
    const notes = deliveredNotes(directory, gate) ?? ""
    expect(notes, label).not.toContain("File names not applied")
    expect(notes, label).not.toContain(record.file!)
  }
})

test("a file name the lawyer's words give outside the instruction a row quotes is no item", async () => {
  const directory = twoDocuments(`The attached Supply terms.docx sets out the terms. ${REQUEST}`)
  const row = { document: "d1", role: "body", file: "Supply terms.docx", instruction_quote: "Prepare a supply note from the supply terms" }
  expect((await acceptStageReply(directory, "package-review", review([row]), { final: true })).issues).toEqual([])
  expect(readPackageReview(directory)?.withheld_delivery).toEqual([{ document: "d1", role: "body", file: "Supply terms.docx", ...NOT_QUOTED }])
  expect(withheldDeliveryItems(directory)).toEqual([])
})

test("a name asked for in another case is disclosed in the lawyer's spelling", async () => {
  const directory = twoDocuments(`${REQUEST} Save the supply note as supply terms.docx.`)
  const row = { document: "d1", role: "body", file: "Supply terms.docx", instruction_quote: "Save the supply note as supply terms.docx" }
  expect((await acceptStageReply(directory, "package-review", review([row]), { final: true })).issues).toEqual([])
  expect(readPackageReview(directory)?.withheld_delivery).toEqual([asked("Supply terms.docx", "not_quoted", "no naming instruction in the lawyer's words", "d1", "body", "supply terms.docx")])
  expect(withheldDeliveryItems(directory).map(item => item.detail)).toEqual([`Your instructions mention the name "supply terms.docx", but it was not applied: "Supply note" keeps its planned file name because the name was not selected exactly as your instructions spell it.`])
})

test("variants of a kept row's name add nothing: no record and no item", async () => {
  const directory = twoDocuments(`${REQUEST} ${SAVE}.`)
  const variants = [{ ...GROUNDED, file: " Supply terms.docx" }, { ...GROUNDED, file: "supply terms.DOCX" }, { ...GROUNDED, instruction_quote: "Prepare a supply note" }]
  expect((await acceptStageReply(directory, "package-review", review([GROUNDED, ...variants]), { final: true })).issues).toEqual([])
  expect(readPackageReview(directory)).toEqual({ status: "pass", findings: [], limitations: ["A disclosed gap."], delivery: [GROUNDED] })
  // A variant on another document is withheld, but the name it gives is applied.
  const other = { ...GROUNDED, document: "d2", file: "supply terms.docx" }
  expect((await acceptStageReply(directory, "package-review", review([GROUNDED, other]), { final: true })).issues).toEqual([])
  expect(readPackageReview(directory)?.withheld_delivery).toEqual([asked("supply terms.docx", "not_quoted", "no naming instruction in the lawyer's words", "d2", "body", "Supply terms.docx")])
  expect(withheldDeliveryItems(directory)).toEqual([])
})

test("two names for one document are both withheld and each disclosed once", async () => {
  const directory = twoDocuments(`${REQUEST} ${SAVE}. Save the supply note as Supply final.docx.`)
  const final = { document: "d1", role: "body", file: "Supply final.docx", instruction_quote: "Save the supply note as Supply final.docx" }
  expect((await acceptStageReply(directory, "package-review", review([GROUNDED, final]), { final: true })).issues).toEqual([])
  const conflict = "conflicts with another delivery selection"
  expect(readPackageReview(directory)).toMatchObject({ delivery: [], withheld_delivery: [asked("Supply terms.docx", "conflict", conflict), asked("Supply final.docx", "conflict", conflict)] })
  expect(withheldDeliveryItems(directory).map(item => item.detail)).toEqual(["Supply terms.docx", "Supply final.docx"].map(name =>
    `Your instructions mention the name "${name}", but it was not applied: "Supply note" keeps its planned file name because the name conflicts with another file name chosen for this package.`))
})

test("one name refused twice is one item", async () => {
  const directory = twoDocuments(`${REQUEST} ${SAVE}.`)
  expect((await acceptStageReply(directory, "package-review", review([{ ...GROUNDED, document: "d9" }, { ...GROUNDED, role: "clean" }]), { final: true })).issues).toEqual([])
  expect(readPackageReview(directory)?.withheld_delivery).toEqual([asked("Supply terms.docx", "document", "names no planned document", "d9"),
    asked("Supply terms.docx", "role", "names no role the run delivers", "d1", "clean")])
  // The row matched to no document adds nothing beside the file the name was not applied to.
  expect(withheldDeliveryItems(directory).map(item => item.detail)).toEqual([`Your instructions mention the name "Supply terms.docx", but it was not applied: "Supply note" keeps its planned file name because it was not settled which file takes the name.`])
})

test("a name the row's own document already has is no item; one only another document has is", async () => {
  const own = twoDocuments(`${REQUEST} Save the supply note as Supply note.docx.`)
  const role = { document: "d1", role: "clean", file: "Supply note.docx", instruction_quote: "Save the supply note as Supply note.docx" }
  expect((await acceptStageReply(own, "package-review", review([role]), { final: true })).issues).toEqual([])
  expect(readPackageReview(own)?.withheld_delivery).toEqual([asked("Supply note.docx", "role", "names no role the run delivers", "d1", "clean")])
  expect(withheldDeliveryItems(own)).toEqual([])
  const other = twoDocuments(`${REQUEST} Save the supply note as Covering letter.docx.`)
  const wrong = { document: "d9", role: "body", file: "Covering letter.docx", instruction_quote: "Save the supply note as Covering letter.docx" }
  expect((await acceptStageReply(other, "package-review", review([wrong]), { final: true })).issues).toEqual([])
  expect(withheldDeliveryItems(other).map(item => item.detail)).toEqual([`Your instructions mention the name "Covering letter.docx", but it was not applied: the name could not be matched to a document in this package.`])
})

test("a reply's own withheld_delivery is never kept, on any attempt", async () => {
  const claimed = { withheld_delivery: [asked("Invented.docx", "document", "names no planned document")] }
  for (const final of [false, true]) {
    const directory = twoDocuments()
    expect((await acceptStageReply(directory, "package-review", review([], claimed), { final })).issues).toEqual([])
    expect(readPackageReview(directory)).toEqual({ status: "pass", findings: [], limitations: ["A disclosed gap."], delivery: [] })
    expect(withheldDeliveryItems(directory)).toEqual([])
  }
})

const MENTION = "The attached Supply terms.docx sets out the terms"

/** What the lawyer reads after the last attempt: the gate's items and the delivered notes. */
async function lawyerText(directory: string, rows: unknown[]) {
  expect((await acceptStageReply(directory, "package-review", review(rows), { final: true })).issues).toEqual([])
  const gate = finalGate(directory)
  return [...gate.items.map(item => item.detail), deliveredNotes(directory, gate) ?? ""].join("\n")
}

test("a name the quoted instruction only mentions is disclosed as mentioned, never as asked for", async () => {
  const mentioned = `Your instructions mention the name "Supply terms.docx", but it was not applied:`
  const cases: [string, unknown[], string, { file?: string }][] = [
    ["no such document", [{ ...GROUNDED, document: "d9", instruction_quote: MENTION }], "the name could not be matched to a document in this package", {}],
    ["no role", [{ ...GROUNDED, role: "clean", instruction_quote: MENTION }], `"Supply note" keeps its planned file name because it was not settled which file takes the name`, {}],
    ["two files", [{ ...GROUNDED, instruction_quote: MENTION }, { ...GROUNDED, document: "d2", instruction_quote: MENTION }],
      `"Supply note" and "Covering letter" keep their planned file names because the name conflicts with another file name chosen for this package`, {}],
    ["another planned file's name", [{ ...GROUNDED, instruction_quote: MENTION }],
      `"Supply note" keeps its planned file name because another file in this package already has the name`, { file: "Supply terms.docx" }],
  ]
  for (const [label, rows, rest, options] of cases) {
    const directory = twoDocuments(`${MENTION}. ${REQUEST}`, options)
    writeLeanText(directory, "drafting-notes.md", "# Notes\nThe supply note records the agreed terms.")
    const text = await lawyerText(directory, rows)
    expect(withheldDeliveryItems(directory).map(item => item.detail), label).toEqual([`${mentioned} ${rest}.`])
    // The gate item and the delivered notes say the same, and neither claims a request.
    expect(text.split(`${mentioned} ${rest}.`), label).toHaveLength(3)
    expect(text, label).not.toMatch(/asked|request/i)
  }
})

test("a name that is only part of what the lawyer's words give is no item: a title cut short, or a name inside a longer one", async () => {
  const cases: [string, string, string][] = [
    ["a title cut short", `${REQUEST} The supply note should be titled Supply terms and conditions.`, "The supply note should be titled Supply terms"],
    ["inside a longer name", `${REQUEST} Save the supply note as ACME-Supply terms.docx.`, "Save the supply note as ACME-Supply terms.docx"],
    ["after a full stop inside a longer name", `${REQUEST} Save the supply note as ACME.Supply terms.docx.`, "Save the supply note as ACME.Supply terms.docx"],
  ]
  for (const [label, request, quote] of cases) {
    const directory = twoDocuments(request)
    expect((await acceptStageReply(directory, "package-review", review([{ ...GROUNDED, document: "d9", instruction_quote: quote }]), { final: true })).issues, label).toEqual([])
    expect(readPackageReview(directory)?.withheld_delivery, label).toEqual([{ document: "d9", role: "body", file: "Supply terms.docx", cause: "document", reason: "names no planned document", asked: false }])
    expect(withheldDeliveryItems(directory), label).toEqual([])
  }
})

test("a quote that is not in the lawyer's words gives no item, though it holds the name", async () => {
  const directory = twoDocuments(`${REQUEST} ${SAVE}.`)
  const row = { ...GROUNDED, document: "d9", instruction_quote: "Name the supply note Supply terms.docx" }
  expect((await acceptStageReply(directory, "package-review", review([row]), { final: true })).issues).toEqual([])
  expect(readPackageReview(directory)?.withheld_delivery).toEqual([{ document: "d9", role: "body", file: "Supply terms.docx", cause: "document", reason: "names no planned document", asked: false }])
  expect(withheldDeliveryItems(directory)).toEqual([])
})

test("a name one row's file already has is no item, though another row gives it to another file", async () => {
  const directory = twoDocuments(`${REQUEST} Save the supply note as Covering letter.docx. Save the covering letter as Covering letter.docx.`)
  const rows = [{ document: "d1", role: "clean", file: "Covering letter.docx", instruction_quote: "Save the supply note as Covering letter.docx" },
    { document: "d2", role: "clean", file: "Covering letter.docx", instruction_quote: "Save the covering letter as Covering letter.docx" }]
  expect((await acceptStageReply(directory, "package-review", review(rows), { final: true })).issues).toEqual([])
  expect(readPackageReview(directory)?.withheld_delivery).toEqual([asked("Covering letter.docx", "role", "names no role the run delivers", "d1", "clean"),
    asked("Covering letter.docx", "role", "names no role the run delivers", "d2", "clean")])
  expect(withheldDeliveryItems(directory)).toEqual([])
})

test("a name selected with a folder path is disclosed in the lawyer's spelling, never with the path", async () => {
  for (const file of [".dochaus/lean/d1/Supply terms.docx", "C:\\Matters\\Supply terms.docx"]) {
    const directory = twoDocuments(`${REQUEST} ${SAVE}.`)
    writeLeanText(directory, "drafting-notes.md", "# Notes\nThe supply note records the agreed terms.")
    const text = await lawyerText(directory, [{ ...GROUNDED, file }])
    expect(readPackageReview(directory)?.withheld_delivery, file).toEqual([asked(file, "file", "not a plain Word file name", "d1", "body", "Supply terms.docx")])
    expect(withheldDeliveryItems(directory).map(item => item.detail), file).toEqual([`Your instructions mention the name "Supply terms.docx", but it was not applied: "Supply note" keeps its planned file name because the name was not selected exactly as your instructions spell it.`])
    expect(text, file).not.toMatch(/\.dochaus|lean\/|Matters/)
  }
})

test("one name refused for two documents is one item naming both", async () => {
  const directory = twoDocuments(`${REQUEST} ${SAVE}. Save the covering letter as Supply terms.docx.`)
  const rows = [{ ...GROUNDED, role: "clean" }, { ...GROUNDED, document: "d2", role: "clean", instruction_quote: "Save the covering letter as Supply terms.docx" }]
  expect((await acceptStageReply(directory, "package-review", review(rows), { final: true })).issues).toEqual([])
  expect(withheldDeliveryItems(directory).map(item => item.detail)).toEqual([`Your instructions mention the name "Supply terms.docx", but it was not applied: "Supply note" and "Covering letter" keep their planned file names because it was not settled which file takes the name.`])
})

test("a file name's surrounding spaces alone never refuse its row on the last attempt", async () => {
  const directory = twoDocuments(`${REQUEST} ${SAVE}.`)
  const padded = { ...GROUNDED, file: "Supply terms.docx " }
  // Earlier attempts still ask for the bare name.
  expect((await acceptStageReply(directory, "package-review", review([padded]))).issues).toEqual(["Delivery file must be the requested Word filename, without a folder path."])
  expect((await acceptStageReply(directory, "package-review", review([padded]), { final: true })).issues).toEqual([])
  expect(readPackageReview(directory)).toEqual({ status: "pass", findings: [], limitations: ["A disclosed gap."], delivery: [GROUNDED] })
  expect(withheldDeliveryItems(directory)).toEqual([])
})

test("a title the lawyer gives without an extension is disclosed as written, never with the .docx the run adds", async () => {
  const request = `${REQUEST} The supply note should be titled Supply terms.`
  const quote = "The supply note should be titled Supply terms"
  const directory = twoDocuments(request)
  writeLeanText(directory, "drafting-notes.md", "# Notes\nThe supply note records the agreed terms.")
  const text = await lawyerText(directory, [{ ...GROUNDED, document: "d9", instruction_quote: quote }])
  const detail = `Your instructions mention the name "Supply terms", but it was not applied: the name could not be matched to a document in this package.`
  expect(withheldDeliveryItems(directory).map(item => item.detail)).toEqual([detail])
  expect(text.split(detail)).toHaveLength(3)
  expect(text).not.toContain("Supply terms.docx")
  expect(readPackageReview(directory)?.withheld_delivery).toEqual([asked("Supply terms.docx", "document", "names no planned document", "d9", "body", "Supply terms.docx", "Supply terms")])
  // The file name the title gives still matches a kept row's name: no item.
  const kept = twoDocuments(request)
  const rows = [{ ...GROUNDED, instruction_quote: quote }, { ...GROUNDED, role: "clean", instruction_quote: quote }]
  expect((await acceptStageReply(kept, "package-review", review(rows), { final: true })).issues).toEqual([])
  expect(readPackageReview(kept)).toMatchObject({ delivery: [rows[0]], withheld_delivery: [asked("Supply terms.docx", "role", "names no role the run delivers", "d1", "clean", "Supply terms.docx", "Supply terms")] })
  expect(withheldDeliveryItems(kept)).toEqual([])
})

// A name the lawyer's words open with any mark but a hyphen, a full stop, an
// ampersand or a plus sign (a name they join is cut short, above and below), or
// give in a folder of their own:
// [label, instruction, rows, the name quoted, why].
const ANGLE = "Save the supply note as <Supply terms.docx>"
const UNMATCHED = "the name could not be matched to a document in this package"
const OPENED: [string, string, unknown[], string, string][] = [
  ["an angle bracket", ANGLE, [{ ...GROUNDED, document: "d9", instruction_quote: ANGLE }], "Supply terms.docx", UNMATCHED],
  ["an angle bracket, two files", ANGLE, [{ ...GROUNDED, instruction_quote: ANGLE }, { ...GROUNDED, document: "d2", instruction_quote: ANGLE }], "Supply terms.docx",
    `"Supply note" and "Covering letter" keep their planned file names because the name conflicts with another file name chosen for this package`],
  ["low quotation marks", "Save the supply note as „Supply terms.docx“", [{ ...GROUNDED, document: "d9", instruction_quote: "Save the supply note as „Supply terms.docx" }], "Supply terms.docx", UNMATCHED],
  ["a dash", "Save the supply note as—Supply terms.docx", [{ ...GROUNDED, document: "d9", instruction_quote: "Save the supply note as—Supply terms.docx" }], "Supply terms.docx", UNMATCHED],
  ["their own folder", "Save the supply note as drafts/Supply terms.docx",
    [{ ...GROUNDED, file: "drafts/Supply terms.docx", instruction_quote: "Save the supply note as drafts/Supply terms.docx" }], "drafts/Supply terms.docx",
    `"Supply note" keeps its planned file name because a file name cannot include a folder`],
]
for (const [label, instruction, rows, name, rest] of OPENED)
  test(`a name the lawyer's words give after ${label} is disclosed as they write it`, async () => {
    const request = `${REQUEST} ${instruction}.`
    const directory = twoDocuments(request)
    writeLeanText(directory, "drafting-notes.md", "# Notes\nThe supply note records the agreed terms.")
    const text = await lawyerText(directory, rows)
    const detail = `Your instructions mention the name "${name}", but it was not applied: ${rest}.`
    expect(withheldDeliveryItems(directory).map(item => item.detail)).toEqual([detail])
    expect(text.split(detail)).toHaveLength(3)
    expect(request).toContain(name)
  })

// A name the lawyer's words join to a longer one with an ampersand or a plus
// sign, as with a hyphen, is cut short: no item.
for (const [label, joined] of [["an ampersand", "Smith&Supply terms.docx"], ["a plus sign", "Q3+Supply terms.docx"]] as const)
  test(`a name joined to a longer one by ${label} is no item`, async () => {
    const quote = `Save the supply note as ${joined}`
    const directory = twoDocuments(`${REQUEST} ${quote}.`)
    expect((await acceptStageReply(directory, "package-review", review([{ ...GROUNDED, document: "d9", instruction_quote: quote }]), { final: true })).issues).toEqual([])
    expect(readPackageReview(directory)?.withheld_delivery).toEqual([{ document: "d9", role: "body", file: "Supply terms.docx", cause: "document", reason: "names no planned document", asked: false }])
    expect(withheldDeliveryItems(directory)).toEqual([])
  })

test("a title selected in a folder only the model gives is disclosed as the lawyer writes it, for its spelling, the title being a usable name", async () => {
  const quote = "The supply note should be titled Supply terms"
  const directory = twoDocuments(`${REQUEST} ${quote}.`)
  writeLeanText(directory, "drafting-notes.md", "# Notes\nThe supply note records the agreed terms.")
  const text = await lawyerText(directory, [{ ...GROUNDED, file: "drafts/Supply terms.docx", instruction_quote: quote }])
  expect(readPackageReview(directory)?.withheld_delivery).toEqual([asked("drafts/Supply terms.docx", "file", "not a plain Word file name", "d1", "body", "Supply terms.docx", "Supply terms")])
  const detail = `Your instructions mention the name "Supply terms", but it was not applied: "Supply note" keeps its planned file name because the name was not selected exactly as your instructions spell it.`
  expect(withheldDeliveryItems(directory).map(item => item.detail)).toEqual([detail])
  expect(text.split(detail)).toHaveLength(3)
  expect(text).not.toContain("drafts")
})

test("a quote the lawyer's words hold only when its letters are lowered one by one gives no item", async () => {
  // Lowered in context, the quote's last sigma is final and the lawyer's is not.
  const directory = twoDocuments(`${REQUEST} Save the supply note as Report.docx ΤΑΣΟΣ.`)
  const row = { document: "d1", role: "body", file: "Report.docx", instruction_quote: "Save the supply note as Report.docx ΤΑΣ" }
  expect((await acceptStageReply(directory, "package-review", review([row]), { final: true })).issues).toEqual([])
  expect(readPackageReview(directory)?.withheld_delivery).toEqual([{ document: "d1", role: "body", file: "Report.docx", ...NOT_QUOTED }])
  expect(withheldDeliveryItems(directory)).toEqual([])
})

test("a name a kept row applies with a run of spaces is no item", async () => {
  const directory = twoDocuments(`${REQUEST} ${SAVE}.`)
  const spaced = { ...GROUNDED, file: "Supply  terms.docx" }
  expect((await acceptStageReply(directory, "package-review", review([spaced, { ...GROUNDED, role: "clean" }]), { final: true })).issues).toEqual([])
  expect(readPackageReview(directory)).toMatchObject({ delivery: [spaced], withheld_delivery: [asked("Supply terms.docx", "role", "names no role the run delivers", "d1", "clean")] })
  expect(withheldDeliveryItems(directory)).toEqual([])
})

test("a gate that discloses a file name not applied needs its delivered notes, even without drafting notes", async () => {
  const directory = twoDocuments(`${REQUEST} ${SAVE}.`)
  expect((await acceptStageReply(directory, "package-review", review([{ ...GROUNDED, document: "d9" }]), { final: true })).issues).toEqual([])
  const gate = finalGate(directory)
  expect(gate.disclosed).toHaveLength(1)
  const delivered = deliveredNotes(directory, gate)!
  expect(delivered).toContain("# File names not applied")
  writeLeanText(directory, "gate.json", JSON.stringify(gate))
  writeLeanText(directory, "drafting-notes-coverage.md", notesCoverageReport(gate))
  expect(stageOutputValid(directory, "gate")).toBe(false)
  writeLeanText(directory, DELIVERED_NOTES.markdown, delivered)
  writeLeanBytes(directory, DELIVERED_NOTES.docx, (await renderMarkdownToDocx(delivered)).bytes)
  expect(stageOutputValid(directory, "gate")).toBe(true)
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
