import { quoteFound, instructionFound } from "../workflows/lean/quotes"
import { LEAN_PROMPT_VERSION, stageInputs } from "../workflows/lean/stages"
import { LEAN_STATE_SCHEMA, hashStageInputs, readLeanState, writeLeanText } from "../workflows/lean/state"
import { createHash } from "node:crypto"
import { afterEach, expect, test } from "bun:test"
import { mkdtempSync, rmSync, readFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { Database } from "bun:sqlite"
import { z } from "zod"
import { renderMarkdownToDocx } from "../lib/render"
import { ensureParagraphRows, paragraphSha256 } from "../lib/drafting-amendment"
import { exportDocxComments, readDocxCommentParagraphs } from "../lib/docx-comments"
import { docxodus } from "../lib/docxodus"
import { currentDocumentView } from "../lib/current-document-retrieval"
import { writeLeanPlan } from "../workflows/lean/plan"
import { leanFile, writeLeanBytes } from "../workflows/lean/state"
import searchTool from "../tool/search-document"
import readTool from "../tool/read-document"
const roots: string[] = []
afterEach(() => roots.splice(0).forEach(root => rmSync(root, { recursive: true, force: true })))
function fixture() {
  const directory = mkdtempSync(path.join(tmpdir(), "current-evidence-"));roots.push(directory)
  writeLeanPlan(directory, [{ title: "Record", purpose: "Record", file: "record.docx", parts: [{ heading: "Record", scope: "Record" }] }, { title: "Other", purpose: "Related record", file: "other.docx", parts: [{ heading: "Other", scope: "Other" }] }]);return directory
}
function acceptApplyForTest(directory: string, document = "d1") {
 const step = `${document}.apply` as `d${number}.apply`
 const prior = readLeanState(directory)
 writeLeanText(directory, "state.json", JSON.stringify({ schema: LEAN_STATE_SCHEMA, profile: "lean", stages: { ...prior?.stages,
 [step]: { status: "done", input_sha: hashStageInputs(step, LEAN_PROMPT_VERSION, stageInputs(directory, step)) } } }))
}
async function save(directory: string, file: string, text: string) { const bytes = (await renderMarkdownToDocx(text)).bytes;writeLeanBytes(directory, file, bytes);if (file.endsWith("/body-applied.docx")) acceptApplyForTest(directory, file.split("/")[0]);return bytes }
async function search(directory: string, args: Record<string, unknown>) {
  const result = await searchTool.execute(z.object(searchTool.args).parse(args), { directory } as Parameters<typeof searchTool.execute>[1]);return typeof result === "string" ? result : JSON.parse(result.output)
}
async function read(directory: string, args: Record<string, unknown>) {
  const result = await readTool.execute(z.object(readTool.args).parse(args), { directory } as Parameters<typeof readTool.execute>[1]);if (typeof result === "string") throw Error(result)
  expect(result.output.length).toBeLessThanOrEqual(6000);return JSON.parse(result.output)
}
test("real current paragraph texts and PIDs are accessible; unindexed source, missing current doc and current no-hits are distinct", async () => {
  const dir = fixture(), bytes = await save(dir, "d1/body-applied.docx", "# Custody\n\nBefore receipt.\n\nViolet receipt recorded.\n\nAfter receipt.")
  const db = new Database(path.join(dir, ".dochaus/legal.db"));db.exec("CREATE TABLE chunks (doc_name TEXT, text TEXT)");db.close();acceptApplyForTest(dir)
  expect(await search(dir, { document: "record.docx", query: "violet", view: "source" })).toContain("Source view unavailable")
  const hit = await search(dir, { document: "d1", query: "violet", view: "current_body" });expect(hit.status).toBe("searched");expect(hit.hits[0].text).toBe("Violet receipt recorded.");expect(hit.hits[0].pids.length).toBeGreaterThan(0)
  const nohit = await search(dir, { document: "d1", query: "zirconium", view: "current_body" });expect(nohit.status).toBe("searched");expect(nohit.hits).toEqual([])
  expect((await search(dir, { document: "missing.docx", query: "violet", view: "current_body" })).status).toBe("unavailable")
  const page = await read(dir, { document: "record.docx", view: "current_body", offset: 1, count: 3 })
  expect(page.entries.map(row => row.text)).toEqual(readDocxCommentParagraphs(bytes).slice(1, 4).map(row => row.paragraph_text));expect(page.entries.every(row => row.pids.length && row.artifact.endsWith("/body-applied.docx"))).toBe(true);expect(page.next_offset).toBeNull()
  expect(readFileSync(leanFile(dir, "d1/body-applied.docx"))).toEqual(Buffer.from(bytes))
})
test("long Unicode paragraphs page without lost characters or gaps and include continuation coordinates", async () => {
  const dir = fixture();await save(dir, "d1/body.docx", "α😀é\t  ".repeat(1500) + "\n\nLast paragraph.")
  const original = currentDocumentView(dir, "d1", "current_body");if (original.status !== "available") throw Error(original.reason)
  const pieces = new Map<number, string[]>();let offset = 0, text_offset = 0, calls = 0
  for (;;) {
    const page = await read(dir, { document: "d1", view: "current_body", offset, text_offset, limit: 1000, count: 3 });calls++
    for (const entry of page.entries) { const prior = pieces.get(entry.index) ?? [];prior.push(entry.text);pieces.set(entry.index, prior) }
    if (page.next_offset === null) break
    expect(page.next_offset > offset || page.next_text_offset > text_offset).toBe(true);offset = page.next_offset;text_offset = page.next_text_offset;if (calls > 100) throw Error("No EOF")
  }
  expect([...pieces.values()].map(parts => parts.join(""))).toEqual(original.entries.map(row => row.text));expect(calls).toBeGreaterThan(2)
})
test("native comments and tracked deletions never become current body; related documents keep separate identity", async () => {
  const dir = fixture(), prior = (await renderMarkdownToDocx("Prior violet receipt.")).bytes, next = (await renderMarkdownToDocx("Revised violet receipt.")).bytes
  const redline = await (await docxodus()).compareDocuments(prior, next, { authorName: "Reviewer" }), paragraph = readDocxCommentParagraphs(redline)[0]!
  const annotated = exportDocxComments(redline, [{ id: "note", author: "Reviewer", text: "Separate annotation only.", anchor: { paragraph_text: paragraph.paragraph_text, paragraph_index: paragraph.paragraph_index } }]);expect(annotated.unresolved).toEqual([])
  writeLeanBytes(dir, "d1/redline.docx", annotated.bytes);await save(dir, "d1/body-applied.docx", "Latest violet receipt.");await save(dir, "d2/body.docx", "Related custody definition.")
  const comments = await search(dir, { document: "d1", view: "comments", query: "annotation" });expect(comments.hits[0].text).toBe("Separate annotation only.");expect(comments.hits[0].artifact).toEndWith("/redline.docx");expect(comments.hits[0].markers.length).toBeGreaterThan(0)
  expect((await search(dir, { document: "d1", view: "current_body", query: "annotation" })).hits).toEqual([])
  const deleted = await read(dir, { document: "d1", view: "deleted_text" });expect(deleted.entries.map(row => row.text).join(" ")).toContain("Prior");expect(deleted.entries.every(row => row.artifact.endsWith("/redline.docx"))).toBe(true)
  expect((await read(dir, { document: "d1", view: "current_body" })).entries.map(row => row.text)).toEqual(["Latest violet receipt."])
  expect((await search(dir, { document: "other.docx", view: "current_body", query: "related" })).hits[0].text).toBe("Related custody definition.")
  writeLeanBytes(dir, "d1/body-applied.docx", new TextEncoder().encode("corrupt"));expect((await search(dir, { document: "d1", view: "current_body", query: "latest" })).status).toBe("unavailable")
})

test("current reads observe changed bytes and invalid cursors explicitly, without falling back from a broken applied artifact", async () => {
  const dir = fixture();await save(dir, "d1/body.docx", "Original custody wording.")
  const first = await read(dir, { document: "d1", view: "current_body" });expect(first.entries[0].text).toBe("Original custody wording.")
  await save(dir, "d1/body-applied.docx", "Updated custody wording.")
  const second = await read(dir, { document: "d1", view: "current_body" });expect(second.entries[0].text).toBe("Updated custody wording.");expect(second.entries[0].sha256).not.toBe(first.entries[0].sha256)
  expect((await read(dir, { document: "d1", view: "current_body", offset: 100 })).status).toBe("unavailable")
  expect((await read(dir, { document: "d1", view: "current_body", text_offset: 10000 })).status).toBe("unavailable")
  expect((await read(dir, { document: "d1", view: "current_body", offset: 1 })).next_offset).toBeNull()
  expect(await search(dir, { view: "current_body", query: "custody" })).toContain("no search was performed")
  expect(await search(dir, { view: "current_body", document: "d1", query: "custody", structural_path: "Custody" })).toContain("no search was performed")
})


test("pending projection is searchable separately from actual Word, preserving split/deletion coordinates and input provenance", async () => {
  const dir = fixture()
  const initial = ["No additional explanation.", "Separate scope.", "Obsolete provision."]
    .map((text, index) => ({ pid: `p${index + 1}`, text, sha256: createHash("sha256").update(text).digest("hex") }))
  await save(dir, "d1/body.docx", initial.map(row => row.text).join("\n\n"))
  writeLeanText(dir, "d1/paragraphs.json", JSON.stringify(initial))
  const changes = { amendments: [
    { pid: "p1", existing_sha256: initial[0]!.sha256, updated: "The proposed explanation identifies the ambiguous basket and preserves the cap qualification." },
    { pid: "p2", existing_sha256: initial[1]!.sha256, operation: "split", updated: ["First projected part.", "Second projected part."] },
    { pid: "p3", existing_sha256: initial[2]!.sha256, operation: "delete", updated: "" },
  ], questions: [] }
  writeLeanText(dir, "d1/changes.json", JSON.stringify(changes))
  const actual = await read(dir, { document: "d1", view: "current_body" })
  expect(actual.entries.map(row => row.text)).toEqual(initial.map(row => row.text))
  const found = await search(dir, { document: "record.docx", view: "proposed_body", query: "ambiguous basket" })
  expect(found.hits).toHaveLength(1);expect(found.hits[0].initial_pid).toBe("p1")
  expect(found.representation).toContain("not actual Word")
  const proposed = await read(dir, { document: "d1", view: "proposed_body", count: 10 })
  expect(proposed.entries.map(row => row.text)).toEqual([changes.amendments[0]!.updated, "First projected part.", "Second projected part."])
  expect(proposed.entries.map(row => [row.initial_pid, row.proposed_part_index])).toEqual([["p1", 0], ["p2", 0], ["p2", 1]])
  expect(proposed.entries.every(row => row.representation === "generated_proposal" && row.basis.includes(".dochaus/lean/d1/paragraphs.json") && row.artifact.endsWith("changes.json"))).toBe(true)
  expect((await search(dir, { document: "d1", view: "current_body", query: "ambiguous basket" })).hits).toEqual([])
  expect(quoteFound(dir, { source_id: ".dochaus/lean/d1/changes.json", quote: String(changes.amendments[0]!.updated) })).toBe(false)
  expect(instructionFound(dir, String(changes.amendments[0]!.updated))).toBe(false)
  expect((await read(dir, { document: "notes", view: "proposed_body" })).status).toBe("unavailable")
  expect(readFileSync(leanFile(dir, "d1/changes.json"), "utf8")).toBe(JSON.stringify(changes))
})

test("pending projection rejects missing, stale and duplicate operations instead of showing old Word or claiming absence", async () => {
  const dir = fixture(), text = "Original condition.", sha256 = createHash("sha256").update(text).digest("hex")
  await save(dir, "d1/body.docx", text)
  writeLeanText(dir, "d1/paragraphs.json", JSON.stringify([{ pid: "p1", text, sha256 }]))
  expect((await search(dir, { document: "d1", view: "proposed_body", query: "condition" })).status).toBe("unavailable")
  const amendment = { pid: "p1", existing_sha256: "stale", updated: "Proposed condition." }
  writeLeanText(dir, "d1/changes.json", JSON.stringify({ amendments: [amendment], questions: [] }))
  expect((await read(dir, { document: "d1", view: "proposed_body" })).status).toBe("unavailable")
  amendment.existing_sha256 = sha256
  writeLeanText(dir, "d1/changes.json", JSON.stringify({ amendments: [amendment, amendment], questions: [] }))
  expect((await search(dir, { document: "d1", view: "proposed_body", query: "proposed" })).status).toBe("unavailable")
  amendment.pid = "unknown"
  writeLeanText(dir, "d1/changes.json", JSON.stringify({ amendments: [amendment], questions: [] }))
  expect((await read(dir, { document: "d1", view: "proposed_body" })).status).toBe("unavailable")
  expect((await read(dir, { document: "d1", view: "current_body" })).entries[0].text).toBe(text)
})

test("a body Word numbers shows each paragraph's list number beside its text, and a search for the number finds the clause", async () => {
  const dir = fixture()
  await save(dir, "d1/body.docx", "# Supply Agreement\n\n## 1. Delivery\n\n1.1 The Supplier shall deliver the goods.\n\n## 2. Payment\n\n2.1 The Customer shall pay within thirty days.")
  const rows = (await ensureParagraphRows(leanFile(dir, "d1/body.docx"))).map(row => ({ ...row, sha256: paragraphSha256(row.text) }))
  writeLeanText(dir, "d1/paragraphs.json", JSON.stringify(rows))
  writeLeanText(dir, "d1/changes.json", JSON.stringify({ amendments: [], questions: [] }))
  for (const view of ["current_body", "proposed_body"] as const) {
    // The text stays the paragraph's own text, which quotes are checked against.
    const hit = await search(dir, { document: "d1", query: "2.1", view })
    expect(hit.hits[0]).toMatchObject({ label: "2.1", text: "The Customer shall pay within thirty days." })
    expect((await read(dir, { document: "d1", view })).entries.map(row => [row.label, row.text])).toContainEqual(["2.", "Payment"])
  }
})
