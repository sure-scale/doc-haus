import { afterEach, expect, test } from "bun:test"
import { packageReviewPrompt } from "./package-review"
import { parsePlan, writeLeanPlan } from "./plan"
import { stagePrompt } from "./prompts"
import { appliedChangeEvidence, exportAppliedChangeComments } from "./review-evidence"
import { applyAmendments } from "../../lib/drafting-amendment-apply"
import { ensureParagraphPointers, paragraphSha256, type Amendment } from "../../lib/drafting-amendment"
import { readFileSync } from "node:fs"
import { recordLeanDocuments } from "./deliverables"
import { stageInputSha } from "./runner"
import { renderMarkdownToDocx } from "../../lib/render"
import { docxodus } from "../../lib/docxodus"
import { leanFile, writeLeanText, writeLeanBytes } from "./state"
import { cleanupMatters, happyPath, matter, registering, scriptedModel, SOURCE, SUPPORT } from "./test-support"

afterEach(cleanupMatters)

test("a superseded deletion explanation stays unresolved when later reconciliation retains wording", async () => {
  const directory = matter()
  const original = "The archive record is paper."
  const updated = "The archive record is digital."
  writeLeanBytes(directory, "d1/body.docx", (await renderMarkdownToDocx(original)).bytes)
  const initialPath = leanFile(directory, "d1/body.docx")
  const pid = Object.keys(await ensureParagraphPointers(initialPath))[0]!
  const initial = new Uint8Array(readFileSync(initialPath))
  const amendments: Amendment[] = [
    { id: "delete", pid, existing_sha256: paragraphSha256(original), operation: "delete", updated: "", reason: "Remove the archive record paragraph.", sources: SUPPORT },
    { id: "restore", pid, existing_sha256: paragraphSha256(original), updated, reason: "Keep the revised archive record instruction.", sources: SUPPORT },
  ]
  const outPath = leanFile(directory, "d1/body-applied.docx")
  const result = await applyAmendments({ docxPath: initialPath, outPath, amendments, merge: async () => updated })
  expect(result.unapplied).toEqual([])
  expect(result.applied[0]!.result_pids).toEqual([pid])
  const final = new Uint8Array(readFileSync(outPath))
  const redline = await (await docxodus()).compareDocuments(initial, final, { authorName: "Doc.Haus" })
  const exported = exportAppliedChangeComments(redline, final, "d1", amendments, result.applied, [], [{ pid, text: original, sha256: paragraphSha256(original) }])
  expect(exported.comments.attached.map(row => row.id)).toEqual(["restore"])
  expect(exported.comments.unresolved[0]).toMatchObject({ id: "delete", reason: expect.stringContaining("retained wording") })
  expect(exported.openItems).toHaveLength(1)
}, 120_000)

test("delivery context projects the files registration uses without treating the plan as user intent", async () => {
  const directory = matter()
  const request = "Revise input-terms.docx; supply an annotated comparison named delivery-review.docx and a separate explanatory note."
  writeLeanText(directory, "request.md", request)
  const document = { title: "Delivery terms", purpose: "Revise delivery", parts: [{ heading: "Delivery", scope: "Approved period" }], mode: "revise" as const, deliverable: "redline" as const, file: "delivery-review.docx" }
  writeLeanPlan(directory, [document])
  const bytes = (await renderMarkdownToDocx("# Delivery terms\n\n## Delivery\n\nDelivery is due within twelve days.")).bytes
  writeLeanBytes(directory, "d1/body-applied.docx", bytes)
  writeLeanBytes(directory, "d1/redline.docx", bytes)
  writeLeanBytes(directory, "drafting-notes.docx", bytes)
  const projected = JSON.parse(/<delivery_artifacts>\n([\s\S]*?)\n<\/delivery_artifacts>/.exec(packageReviewPrompt(directory))?.[1] ?? "[]") as { kind: string; planned_filename: string; file_status: string }[]
  expect(projected.map(row => [row.kind, row.planned_filename, row.file_status])).toEqual([
    ["body", "delivery-review - clean.docx", "present"],
    ["redline", "delivery-review.docx", "present"],
    ["drafting_notes", "delivery-review - drafting notes.docx", "present"],
  ])
  const registered = await recordLeanDocuments(directory, { ...registering(), sessions: scriptedModel(happyPath).port }, true)
  expect(registered.documents.flatMap(row => row.deliverable ?? [])).toEqual(projected.map(row => row.planned_filename))
  // Change only the model's selection: review must see the resulting mismatch,
  // without relabelling that selection as the lawyer's requested output.
  const admitted = parsePlan("Acting for: not stated\n# Delivery terms\nMode: revise\nDeliverable: body\nFile: input-terms.docx\n## Delivery\nApproved period", request)
  expect(admitted.issues).toEqual([])
  expect(admitted.value?.documents[0]?.file).toBe("input-terms.docx")
  writeLeanPlan(directory, admitted.value!.documents)
  for (const prompt of [stagePrompt(directory, "finalize"), packageReviewPrompt(directory)]) {
    expect(prompt).toContain("<delivery_artifacts>")
    expect(prompt).toContain('"planned_filename": "input-terms - redline.docx"')
    expect(prompt).toContain('"planned_file": "input-terms.docx"')
    expect(prompt).not.toContain('"requested_file": "input-terms.docx"')
    expect(prompt).toContain("delivery-review.docx")
    expect(prompt).toContain('"kind": "drafting_notes"')
  }
  expect(parsePlan("Acting for: not stated\n# Delivery terms\nFile: input-terms.docx\n## Delivery", "Revise input-terms.docx and deliver the revised body as input-terms.docx.").issues).toEqual([])
}, 120_000)

test("memo and final reviewer retain source arrays independently of application mode", () => {
  const directory = matter()
  writeLeanPlan(directory, [{ title: "Supply note", purpose: "Record terms", parts: [{ heading: "Delivery", scope: "Agreed delivery" }] }])
  const amendment = { id: "a-source", pid: "p111111111111", existing_sha256: "initial", updated: "Supported revised wording.", reason: "Carry the supplied delivery period.", sources: SUPPORT }
  const changes = (sources: typeof SUPPORT) => writeLeanText(directory, "d1/changes.json", JSON.stringify({ amendments: [{ ...amendment, sources }], questions: [], shared_paragraphs: [] }))
  changes(SUPPORT)
  const apply = (mode: "code" | "merged") => writeLeanText(directory, "d1/apply.json", JSON.stringify({ applied: [{ id: amendment.id, pid: amendment.pid, mode }], unapplied: [] }))
  apply("code")
  const supported = appliedChangeEvidence(directory, "d1")[0]!
  expect(supported.evidence_basis).toBe("source-passages-provided")
  expect(supported.sources).toEqual(SUPPORT)
  for (const prompt of [stagePrompt(directory, "finalize"), packageReviewPrompt(directory)]) {
    expect(prompt).toContain(SOURCE)
    expect(prompt).toContain(SUPPORT[0]!.quote)
    expect(prompt).toContain('"application_mode": "code"')
    expect(prompt).toContain('"evidence_basis": "source-passages-provided"')
  }
  apply("merged")
  expect(appliedChangeEvidence(directory, "d1")[0]!.evidence_basis).toBe(supported.evidence_basis)
  expect(appliedChangeEvidence(directory, "d1")[0]!.sources).toEqual(SUPPORT)
  const key = stageInputSha(directory, "package-review")
  changes([])
  expect(appliedChangeEvidence(directory, "d1")[0]!.evidence_basis).toBe("no-source-passage")
  expect(appliedChangeEvidence(directory, "d1")[0]!.sources).toEqual([])
  expect(stageInputSha(directory, "package-review")).not.toBe(key)
  expect(packageReviewPrompt(directory)).toContain('"evidence_basis": "no-source-passage"')
})


test("review reads the native comparison separately from the clean body and keys its actual bytes", async () => {
  const directory = matter()
  writeLeanPlan(directory, [{ title: "Equipment terms", purpose: "Revise rent", mode: "revise", deliverable: "redline", parts: [{ heading: "Rent", scope: "Approved amount" }] }])
  const before = (await renderMarkdownToDocx("# Equipment terms\n\n## Rent\n\nThe monthly rent is 2400 units." )).bytes
  const after = (await renderMarkdownToDocx("# Equipment terms\n\n## Rent\n\nThe monthly rent is 2700 units." )).bytes
  const redline = await (await docxodus()).compareDocuments(before, after, { authorName: "doc.haus" })
  writeLeanText(directory, "d1/paragraphs-applied.json", JSON.stringify([{ pid: "p111111111111", sha256: "final", text: "The monthly rent is 2700 units." }]))
  writeLeanBytes(directory, "d1/redline.docx", redline)
  writeLeanText(directory, "d1/apply.json", JSON.stringify({ applied: [], unapplied: [], comments: { author: "Doc.Haus", attached: [], unresolved: [] } }))
  const prompt = packageReviewPrompt(directory)
  expect(prompt).toContain("<native_export")
  expect(prompt).toContain('"delivery_registration": "pending"')
  expect(prompt).toContain('"deleted_text"')
  expect(prompt).toContain("2400")
  expect(prompt).toContain("2700")
  expect(prompt).toContain('"unapplied_amendment_count": 0')
  expect(stagePrompt(directory, "finalize")).toContain("<native_export")
  const packageKey = stageInputSha(directory, "package-review")
  const memoKey = stageInputSha(directory, "finalize")
  writeLeanBytes(directory, "d1/redline.docx", after)
  expect(stageInputSha(directory, "package-review")).not.toBe(packageKey)
  expect(stageInputSha(directory, "finalize")).not.toBe(memoKey)
  expect(packageReviewPrompt(directory)).toContain('"deleted_text": []')
}, 120_000)


test("native outcome counts keep withheld proposals separate from amendment application and initial explanations", () => {
  const directory = matter()
  writeLeanPlan(directory, [{ title: "Supply note", purpose: "Record terms", parts: [{ heading: "Delivery", scope: "Delivery" }] }])
  writeLeanText(directory, "d1/paragraphs.json", JSON.stringify([{ pid: "p111111111111", sha256: "initial", text: "Delivery term" }]))
  writeLeanText(directory, "d1/changes.json", JSON.stringify({ amendments: [], explanations: [{ id: "i1", pid: "p111111111111", reason: "Initial explanation", sources: SUPPORT }], questions: [], shared_paragraphs: [], withheld: [{ amendment: { id: "a1", pid: "p111111111111", existing_sha256: "initial", updated: "Unverified proposal", reason: "Not supported", sources: [] }, issues: ["Missing confirmation marker"] }] }))
  writeLeanText(directory, "d1/apply.json", JSON.stringify({ applied: [], unapplied: [] }))
  const prompt = packageReviewPrompt(directory)
  expect(prompt).toContain('"applied_amendment_count": 0')
  expect(prompt).toContain('"unapplied_amendment_count": 0')
  expect(prompt).toContain('"withheld_proposal_count": 1')
  expect(prompt).toContain('"change_origin": "initial-draft-revision"')
})
