import { afterEach, expect, test } from "bun:test"
import { createHash } from "node:crypto"
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { Database } from "bun:sqlite"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { counterpartyItems, readDocxMarkup, readMarkupComments } from "./docx-markup"
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

// A package with parts edited as text; a missing part is edited from "".
function repackaged(bytes: Uint8Array, edits: Record<string, (xml: string) => string>) {
  const parts = unzipSync(bytes)
  for (const [part, edit] of Object.entries(edits)) parts[part] = strToU8(edit(parts[part] ? strFromU8(parts[part]) : ""))
  return zipSync(parts)
}
// The fixture with one package part edited as text.
const edited = (part: string, edit: (xml: string) => string) => repackaged(counterpartyMarkupDocx(), { [part]: edit })
const by = (author: string) => `w:author="${author}" w:date="${dates[author]}"`
const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`
const inserted = (id: number, author: string, text: string) => `<w:ins w:id="${id}" ${by(author)}>${run(text)}</w:ins>`

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

test("a move's places say which one it leaves and which one it lands in, whichever way it moves", async () => {
  const sides = async (bytes: Uint8Array) => (await readDocxMarkup(bytes)).revisions.filter(row => row.kind === "move").map(row => row.locations.map(at => [at.paragraph, at.side]))
  // The Notices clause moves down, below clause 4.
  expect(await sides(counterpartyMarkupDocx())).toEqual([[[3, "from"], [5, "to"]]])
  // With its sides swapped it moves up: its places still come in document order, so the first is where it lands.
  const up = edited("word/document.xml", xml => xml.replace(/move(From|To)/g, (_, side) => `move${side === "From" ? "To" : "From"}`))
  expect(await sides(up)).toEqual([[[3, "to"], [5, "from"]]])
})

test("both planted instructions are flagged on the item that carries them, their visible text unchanged", async () => {
  const markup = await readDocxMarkup(counterpartyMarkupDocx())
  const flagged = [...markup.revisions, ...markup.comments].filter(item => item.findings.length)
  expect(flagged.map(item => [item.text, item.findings.map(finding => [finding.field, finding.rule])])).toEqual([
    [` ${plantedDeletion}`, [["text", "instruction-override"]]],
    [plantedComment, [["text", "ai-directive"], ["text", "concealment"]]],
  ])
})

test("counterparty items are each content change and open thread, with author, place, folded replies and flags", async () => {
  const bytes = counterpartyMarkupDocx()
  const source = { name: COUNTERPARTY_MARKUP.name, sha256: createHash("sha256").update(bytes).digest("hex") }
  const markup = await readDocxMarkup(bytes)
  const items = counterpartyItems(source, markup)
  const at = (paragraph: number, side?: "from" | "to") => ({ part: "body", paragraph, ...side ? { side } : {} })
  expect(items.map(({ source: _, ref: __, date: ___, ...item }) => item)).toEqual([
    { author: AVERY, change: "delete", text: "thirty", places: [at(1)] },
    { author: AVERY, change: "insert", text: "forty-five", places: [at(1)] },
    { author: MORGAN, change: "insert", text: "the first ", places: [at(2)] },
    { author: AVERY, change: "move", text: `${NOTICES}¶`, places: [at(3, "from"), at(5, "to")] },
    // Morgan's bold on clause 4 is formatting, not a point to answer.
    { author: AVERY, change: "delete", text: ` ${plantedDeletion}`, places: [at(6)], flagged: ["instruction-override"] },
    { author: AVERY, change: "insert", text: "on ten Business Days' notice at any time ", places: [at(7)] },
    { author: MORGAN, change: "delete", text: "at any time ", places: [at(7)] },
    { author: MORGAN, change: "delete", text: "¶", places: [at(8)] },
    { author: MORGAN, change: "insert", text: "8. Waiver. No waiver is effective unless in writing.¶", places: [at(10)] },
    { author: AVERY, change: "comment", text: "Our production lead time needs forty-five days.", quote: "forty-five", places: [at(1)],
      replies: [{ author: MORGAN, date: dates[MORGAN], text: "Agreed if the Supplier confirms stock levels monthly." }] },
    // Morgan's resolved note on clause 2 is settled and left out.
    { author: AVERY, change: "comment", text: plantedComment, quote: "5. Term. This Agreement continues for two years.", places: [at(6)], flagged: ["ai-directive", "concealment"] },
  ])
  // Each joins its card by version and id.
  expect(items.map(item => item.ref)).toEqual([...markup.revisions.filter(row => row.kind !== "format").map(row => ({ kind: "revision", id: row.id })), { kind: "comment", id: 0 }, { kind: "comment", id: 3 }])
  expect(items.map(item => [item.source, item.date])).toEqual(items.map(item => [source, dates[item.author]]))
  // A change in a numbered clause is placed by the number Word prints there.
  const rendered = (await renderMarkdownToDocx("Terms.\n\n1. Supply.\n2. Price.\n3. Term.")).bytes
  const numbered = repackaged(rendered, { "word/document.xml": xml => xml.replace("Term.</w:t></w:r>", `$&${inserted(90, AVERY, " Revised.")}`) })
  expect(counterpartyItems(source, await readDocxMarkup(numbered)).map(item => item.places)).toEqual([[{ part: "body", paragraph: 4, label: "3." }]])
})

test("our own changes and threads are not theirs to answer; a thread they answered marks our entries", async () => {
  const bytes = counterpartyMarkupDocx()
  const source = { name: COUNTERPARTY_MARKUP.name, sha256: createHash("sha256").update(bytes).digest("hex") }
  const markup = await readDocxMarkup(bytes)
  const all = counterpartyItems(source, markup, [])
  const summary = (items: ReturnType<typeof counterpartyItems>) => items.map(item => [item.author, item.change, item.ours ?? false, (item.replies ?? []).map(reply => [reply.author, reply.ours ?? false])])
  // Morgan is our side (matched trimmed, in any case): Morgan's changes go, and Avery's
  // thread keeps Morgan's reply as ours.
  expect(summary(counterpartyItems(source, markup, ["  morgan OKAFOR "]))).toEqual([
    ...all.filter(item => item.change !== "comment" && item.author === AVERY).map(item => [AVERY, item.change, false, []]),
    [AVERY, "comment", false, [[MORGAN, true]]],
    [AVERY, "comment", false, []],
  ])
  // Avery is: Avery's changes and Avery's lone planted comment go; the thread Morgan answered stays, its root ours.
  expect(summary(counterpartyItems(source, markup, [AVERY.toUpperCase()]))).toEqual([
    ...all.filter(item => item.change !== "comment" && item.author === MORGAN).map(item => [MORGAN, item.change, false, []]),
    [AVERY, "comment", true, [[MORGAN, false]]],
  ])
  // Both: nothing is theirs.
  expect(counterpartyItems(source, markup, [AVERY, MORGAN])).toEqual([])
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

test("a tracked edit inside a comment's own text is a comment change, placed in that comment", async () => {
  const markup = await readDocxMarkup(edited("word/comments.xml", xml => xml.replace("Fixed price period confirmed.</w:t></w:r>", `$&${inserted(50, MORGAN, " Subject to review.")}`)))
  const edit = markup.revisions.find(row => row.text === " Subject to review.")!
  // Not body paragraph 3: a comment's paragraphs are no place in the agreement.
  expect([edit.kind, edit.author, edit.scope, edit.locations]).toEqual(["insert", MORGAN, "comment", [{ paragraph: 1, context: "Fixed price period confirmed. Subject to review.", comment: 2 }]])
  expect(markup.revisions.filter(row => row.scope === "body")).toHaveLength(10)
})

test("a table, row, cell or page-setup change is placed at its first paragraph and names what changed", async () => {
  const changed = (name: string, id: number, set: string) => `<w:${name}>${set}<w:${name}Change w:id="${id}" ${by(MORGAN)}><w:${name}/></w:${name}Change></w:${name}>`
  const cell = (text: string, properties = "") => `<w:tc>${properties}<w:p>${run(text)}</w:p></w:tc>`
  const table = `<w:tbl>${changed("tblPr", 60, `<w:tblW w:w="0" w:type="auto"/>`)}<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>` +
    `<w:tr>${cell("Party")}${cell("Address")}</w:tr><w:tr>${changed("trPr", 61, "<w:cantSplit/>")}${cell("Buyer")}${cell("1 High Street", changed("tcPr", 62, `<w:shd w:val="clear" w:fill="D9D9D9"/>`))}</w:tr></w:tbl><w:p/>`
  const section = `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:sectPrChange w:id="63" ${by(MORGAN)}><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:sectPrChange></w:sectPr>`
  const markup = await readDocxMarkup(edited("word/document.xml", xml => xml.replace("<w:sectPr/>", table + section)))
  // The table follows clause 9, paragraph 11; a cell's paragraph counts as a paragraph.
  expect(markup.revisions.filter(row => row.family === "properties_change" && !row.text).map(row => row.locations)).toEqual([
    [{ paragraph: 12, context: "Party", container: "table" }],
    [{ paragraph: 14, context: "Buyer", container: "row" }],
    [{ paragraph: 15, context: "1 High Street", container: "cell" }],
    // A page setup applies from its section's first paragraph.
    [{ ...markup.revisions[0]!.locations[0]!, container: "section" }],
  ])
})

test("a comment in a footnote or endnote has its quote and its place in that part", async () => {
  const notes = (kind: string, body: string) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:${kind}s xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
    `<w:${kind} w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:${kind}><w:${kind} w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:${kind}>` +
    `<w:${kind} w:id="1">${body}</w:${kind}></w:${kind}s>`
  const range = (id: number, text: string) => `<w:commentRangeStart w:id="${id}"/>${run(text)}<w:commentRangeEnd w:id="${id}"/><w:r><w:commentReference w:id="${id}"/></w:r>`
  const PARTS = ["footnotes", "endnotes"]
  const bytes = edited("[Content_Types].xml", xml => xml.replace("</Types>", `${PARTS.map(part => `<Override PartName="/word/${part}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${part}+xml"/>`).join("")}</Types>`))
  const markup = await readDocxMarkup(repackaged(bytes, {
    "word/_rels/document.xml.rels": xml => xml.replace("</Relationships>", `${PARTS.map(part => `<Relationship Id="rId${part}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${part}" Target="${part}.xml"/>`).join("")}</Relationships>`),
    "word/document.xml": xml => xml.replace("the Province.</w:t></w:r>", `$&<w:r><w:footnoteReference w:id="1"/></w:r><w:r><w:endnoteReference w:id="1"/></w:r>`),
    "word/footnotes.xml": () => notes("footnote", `<w:p><w:r><w:footnoteRef/></w:r>${run(" The Province is named in the Schedule.")}</w:p><w:p>${range(4, "Schedule 1 lists it")}${inserted(40, MORGAN, " and its courts")}${run(".")}</w:p>`),
    "word/endnotes.xml": () => notes("endnote", `<w:p>${range(5, "Endnote text.")}</w:p>`),
    "word/comments.xml": xml => xml.replace("</w:comments>", `<w:comment w:id="4" ${by(MORGAN)} w:initials="MO"><w:p>${run("Which schedule?")}</w:p></w:comment><w:comment w:id="5" ${by(AVERY)} w:initials="AL"><w:p>${run("Drop this note.")}</w:p></w:comment>$&`),
  }))
  // Separator notes hold no paragraph a reader sees, so the comment and Morgan's insertion share footnote paragraph 2.
  expect(markup.comments.slice(4).map(({ findings: _, ...comment }) => comment)).toEqual([
    { id: 4, author: MORGAN, initials: "MO", date: dates[MORGAN], text: "Which schedule?", parent: null, resolved: null, quote: "Schedule 1 lists it", paragraph: 2, scope: "footnote" },
    { id: 5, author: AVERY, initials: "AL", date: dates[AVERY], text: "Drop this note.", parent: null, resolved: null, quote: "Endnote text.", paragraph: 1, scope: "endnote" },
  ])
  const insertion = markup.revisions.find(row => row.text === " and its courts")!
  expect([insertion.scope, insertion.locations]).toEqual(["footnote", [{ paragraph: 2, context: "Schedule 1 lists it and its courts." }]])
  // Body comments are placed as before, with no part named.
  expect(markup.comments.slice(0, 4).map(comment => [comment.paragraph, "scope" in comment])).toEqual([[1, false], [1, false], [2, false], [6, false]])
})

test("a numbered paragraph's location carries the number Word prints, apart from its ordinal", async () => {
  const rendered = (await renderMarkdownToDocx("Terms.\n\n1. Supply.\n2. Price.\n3. Term.")).bytes
  const bytes = repackaged(rendered, { "word/document.xml": xml => xml.replace(/(Terms?)\.<\/w:t><\/w:r>/g, (found, word) => `${found}${inserted(word === "Term" ? 90 : 91, AVERY, " Revised.")}`) })
  expect((await readDocxMarkup(bytes)).revisions.map(row => row.locations)).toEqual([
    [{ paragraph: 1, context: "Terms. Revised." }],
    [{ paragraph: 4, context: "Term. Revised.", label: "3." }],
  ])
})

test("a comment in a numbered clause carries the number Word prints there, as a change there does", async () => {
  const rendered = (await renderMarkdownToDocx("Terms.\n\n1. Supply.\n2. Price.\n3. Term.")).bytes
  const paragraphs = readDocxCommentParagraphs(rendered)
  const on = (text: string) => {
    const { paragraph_text, paragraph_index } = paragraphs.find(row => row.paragraph_text === text)!
    return { paragraph_text, paragraph_index }
  }
  const { bytes } = exportDocxComments(rendered, [
    { id: "term", author: "Reviewer", text: "Shorten the term?", anchor: on("Term.") },
    { id: "terms", author: "Reviewer", text: "Define the terms.", anchor: on("Terms.") },
  ])
  // A file with comments and no tracked changes; the unnumbered paragraph has no label.
  expect((await readDocxMarkup(bytes)).comments.map(comment => [comment.text, comment.paragraph, comment.label])).toEqual([
    ["Shorten the term?", 4, "3."],
    ["Define the terms.", 1, undefined],
  ])
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
