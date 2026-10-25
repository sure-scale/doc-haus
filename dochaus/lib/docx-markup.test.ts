import { afterEach, expect, test } from "bun:test"
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { Database } from "bun:sqlite"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { readDocxMarkup, readMarkupComments } from "./docx-markup"
import { currentDocumentView } from "./current-document-retrieval"
import { exportDocxComments, readDocxCommentParagraphs } from "./docx-comments"
import { renderMarkdownToDocx } from "./render"
import { detectInjection } from "./untrusted-text"
import { writeLeanPlan } from "../workflows/lean/plan"
import { writeLeanBytes } from "../workflows/lean/state"
import { COUNTERPARTY_MARKUP, counterpartyMarkupDocx } from "../test/support/counterparty-markup-fixture"
import { revisionTablesOriginalDocx } from "../test/support/revision-tables-fixture"

const [AVERY, MORGAN] = COUNTERPARTY_MARKUP.reviewers as [string, string]
const { dates, plantedDeletion, plantedComment } = COUNTERPARTY_MARKUP
const NOTICES = "3. Notices. Notices must be in writing."
const ZWSP = "\u200B"

const roots: string[] = []
afterEach(() => roots.splice(0).forEach(root => rmSync(root, { recursive: true, force: true })))

// The fixture with one package part edited as text.
function edited(part: string, edit: (xml: string) => string) {
  const parts = unzipSync(counterpartyMarkupDocx())
  parts[part] = strToU8(edit(strFromU8(parts[part]!)))
  return zipSync(parts)
}

test("every tracked change and comment is listed with its author, kind, date and place", async () => {
  const markup = await readDocxMarkup(counterpartyMarkupDocx())
  expect(markup.revisions.map(row => [row.kind, row.family, row.author, row.text, row.locations.map(at => at.paragraph)])).toEqual([
    ["delete", "content_delete", AVERY, "thirty", [1]],
    ["insert", "content_insert", AVERY, "forty-five", [1]],
    ["insert", "content_insert", MORGAN, "the first ", [2]],
    // One move: the paragraph it leaves and the one it lands in, below clause 4.
    ["move", "move", AVERY, `${NOTICES}¶`, [3, 5]],
    ["format", "properties_change", MORGAN, "indirect", [4]],
    ["delete", "content_delete", AVERY, ` ${plantedDeletion}`, [6]],
    // Avery's adjacent insertions are one unit; Morgan's deletion inside them is another.
    ["insert", "content_insert", AVERY, "on ten Business Days' notice at any time ", [7]],
    ["delete", "content_delete", MORGAN, "at any time ", [7]],
    ["delete", "paragraph_mark", MORGAN, "¶", [8]],
    ["insert", "content_insert", MORGAN, "8. Waiver. No waiver is effective unless in writing.¶", [10]],
  ])
  expect(markup.revisions.every(row => row.date === dates[row.author] && row.scope === "body" && row.status === "supported" && row.diagnostic === null)).toBe(true)
  expect(new Set(markup.revisions.map(row => row.id)).size).toBe(markup.revisions.length)
  expect(markup.revisions[3]!.locations.map(at => at.context)).toEqual([NOTICES, NOTICES])
  // Comment paragraphs count as revision locations do: Avery's planted comment sits on clause 5, paragraph 6.
  expect(markup.comments.map(({ findings: _, ...comment }) => comment)).toEqual([
    { id: 0, author: AVERY, initials: "AL", date: dates[AVERY], text: "Our production lead time needs forty-five days.", parent: null, resolved: false, quote: "forty-five", paragraph: 1 },
    { id: 1, author: MORGAN, initials: "MO", date: dates[MORGAN], text: "Agreed if the Supplier confirms stock levels monthly.", parent: 0, resolved: false, quote: "", paragraph: 1 },
    { id: 2, author: MORGAN, initials: "MO", date: dates[MORGAN], text: "Fixed price period confirmed.", parent: null, resolved: true, quote: "twelve months", paragraph: 2 },
    { id: 3, author: AVERY, initials: "AL", date: dates[AVERY], text: plantedComment, parent: null, resolved: false, quote: "5. Term. This Agreement continues for two years.", paragraph: 6 },
  ])
  expect(markup.authors).toEqual([{ name: AVERY, revisions: 5, comments: 2 }, { name: MORGAN, revisions: 5, comments: 2 }])
})

test("both planted instructions are flagged on the item that carries them, their visible text unchanged", async () => {
  const markup = await readDocxMarkup(counterpartyMarkupDocx())
  const flagged = [...markup.revisions, ...markup.comments].filter(item => item.findings.length)
  expect(flagged.map(item => [item.text, item.findings.map(finding => [finding.field, finding.rule])])).toEqual([
    [` ${plantedDeletion}`, [["text", "instruction-override"]]],
    [plantedComment, [["text", "ai-directive"], ["text", "concealment"]]],
  ])
})

test("a zero-width character in a comment is stripped and reported, and cannot hide an instruction from the scan", async () => {
  const hidden = plantedComment.replace("AI assistant", `AI assis${ZWSP}tant`)
  // Unnormalized, the split word escapes the AI-directive rule.
  expect(detectInjection(hidden).map(finding => finding.rule)).not.toContain("ai-directive")
  const markup = await readDocxMarkup(edited("word/comments.xml", xml => xml.replace("Fixed price period", `Fixed price${ZWSP} period`).replace(plantedComment, hidden)))
  expect(markup.comments[2]!.text).toBe("Fixed price period confirmed.")
  expect(markup.comments[2]!.findings.map(finding => [finding.field, finding.rule])).toEqual([["text", "zero-width"]])
  expect(markup.comments[3]!.text).toBe(plantedComment)
  expect(markup.comments[3]!.findings.map(finding => finding.rule)).toEqual(["zero-width", "ai-directive", "concealment"])
})

test("ranges that never close, or name no comment, keep each quote to its cap", async () => {
  // Every run was once appended to every open range: 2,000 unclosed ranges over
  // 20,000 paragraphs took seconds and gigabytes.
  const starts = Array.from({ length: 2_000 }, (_, at) => `<w:commentRangeStart w:id="${100 + at}"/>`).join("")
  const words = `<w:p><w:r><w:t xml:space="preserve">word </w:t></w:r></w:p>`.repeat(20_000)
  const bytes = edited("word/document.xml", xml => xml.replace("<w:p>", `<w:p>${starts}`)
    .replace(`<w:commentRangeEnd w:id="3"/>`, "").replace("<w:sectPr/>", `${words}<w:sectPr/>`))
  const started = performance.now()
  const comments = await readMarkupComments(bytes)
  expect(performance.now() - started).toBeLessThan(2_000)
  expect(comments.map(comment => [comment.id, comment.quote.slice(0, 20)])).toEqual([[0, "forty-five"], [1, ""], [2, "twelve months"], [3, "5. Term. This Agreem"]])
  // Comment 3's range now runs to the end of the document.
  expect(Array.from(comments[3]!.quote)).toHaveLength(201)
})

test("a comment with no thread records has an unknown resolved state, and a clean file gives empty lists", async () => {
  expect(await readDocxMarkup(revisionTablesOriginalDocx())).toEqual({ revisions: [], comments: [], authors: [] })
  const plain = (await renderMarkdownToDocx("The Buyer shall pay within thirty days.")).bytes
  const paragraph = readDocxCommentParagraphs(plain)[0]!
  const flat = exportDocxComments(plain, [{ id: "note", author: "Reviewer", text: "Check the period.", anchor: { paragraph_text: paragraph.paragraph_text, paragraph_index: paragraph.paragraph_index } }])
  const markup = await readDocxMarkup(flat.bytes)
  expect(markup.revisions).toEqual([])
  expect(markup.comments).toEqual([{ id: 0, author: "Reviewer", initials: null, date: null, text: "Check the period.", parent: null, resolved: null, quote: paragraph.paragraph_text, paragraph: 1, findings: [] }])
})

test("agent views of an uploaded source are normalized and attributed; a planned body stays exact", async () => {
  const directory = mkdtempSync(path.join(tmpdir(), "markup-views-"));roots.push(directory)
  const name = COUNTERPARTY_MARKUP.name
  mkdirSync(path.join(directory, ".dochaus"))
  const db = new Database(path.join(directory, ".dochaus/legal.db"));db.exec("CREATE TABLE documents (name TEXT, role TEXT)");db.query("INSERT INTO documents VALUES (?, 'source')").run(name);db.close()
  const source = edited("word/comments.xml", xml => xml.replace("Fixed price period", `Fixed price${ZWSP} period`))
  const parts = unzipSync(source)
  parts["word/document.xml"] = strToU8(strFromU8(parts["word/document.xml"]!).replace("9. Governing law", `9. Governing${ZWSP} law`))
  writeFileSync(path.join(directory, name), zipSync(parts))
  writeLeanPlan(directory, [{ title: "Record", purpose: "Record", file: "record.docx", parts: [{ heading: "Record", scope: "Record" }] }])
  writeLeanBytes(directory, "d1/body.docx", (await renderMarkdownToDocx(`Governing${ZWSP} law.`)).bytes)
  const entries = async (document: string, view: Parameters<typeof currentDocumentView>[2]) => {
    const result = await currentDocumentView(directory, document, view)
    if (result.status !== "available") throw new Error(result.reason)
    return result.entries
  }

  const comments = await entries(name, "comments")
  expect(comments.map(({ comment_id, author, date, parent, resolved, text, flagged }) => ({ comment_id, author, date, parent, resolved, text, flagged }))).toEqual([
    { comment_id: "0", author: AVERY, date: dates[AVERY], parent: null, resolved: false, text: "Our production lead time needs forty-five days.", flagged: undefined },
    { comment_id: "1", author: MORGAN, date: dates[MORGAN], parent: "0", resolved: false, text: "Agreed if the Supplier confirms stock levels monthly.", flagged: undefined },
    { comment_id: "2", author: MORGAN, date: dates[MORGAN], parent: null, resolved: true, text: "Fixed price period confirmed.", flagged: ["zero-width"] },
    { comment_id: "3", author: AVERY, date: dates[AVERY], parent: null, resolved: false, text: plantedComment, flagged: ["ai-directive", "concealment"] },
  ])
  expect(comments[3]!.markers).toEqual(expect.arrayContaining([expect.objectContaining({ kind: "commentRangeStart", paragraph_index: 4, xml_paragraph_index: 5 })]))
  const deleted = await entries(name, "deleted_text")
  expect(deleted.filter(entry => entry.flagged).map(entry => [entry.text, entry.flagged])).toEqual([[` ${plantedDeletion}`, ["instruction-override"]]])
  const body = await entries(name, "current_body")
  expect(body.find(entry => entry.text.startsWith("9."))).toMatchObject({ text: "9. Governing law. This Agreement is governed by the law of the Province.", flagged: ["zero-width"] })
  // Amendments target a planned document's paragraphs by their exact text.
  expect((await entries("d1", "current_body")).map(entry => [entry.text, entry.flagged])).toEqual([[`Governing${ZWSP} law.`, undefined]])
})
