import { expect, test } from "bun:test"
import mammoth from "mammoth"
import { strToU8, zipSync } from "fflate"
import { markTrackedChanges, proposedView } from "./docx-tracked-text"
import { docxodus } from "./docxodus"
import { renderMarkdownToDocx } from "./render"

// Every fixture is invented text: no actual party, transaction or reviewer.

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const TYPES = '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
const RELS = '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'

function docx(body: string, comments?: string) {
  return Buffer.from(zipSync({
    "[Content_Types].xml": strToU8(TYPES),
    "_rels/.rels": strToU8(RELS),
    "word/document.xml": strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${body}</w:body></w:document>`),
    ...(comments ? { "word/comments.xml": strToU8(`<?xml version="1.0" encoding="UTF-8"?><w:comments ${W}>${comments}</w:comments>`) } : {}),
  }))
}
const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`
const deleted = (text: string) => `<w:r><w:delText xml:space="preserve">${text}</w:delText></w:r>`
const ins = (author: string, content: string) => `<w:ins w:id="1" w:author="${author}" w:date="2026-01-01T00:00:00Z">${content}</w:ins>`
const del = (author: string, content: string) => `<w:del w:id="2" w:author="${author}" w:date="2026-01-01T00:00:00Z">${content}</w:del>`
const text = async (buffer: Buffer) => (await mammoth.extractRawText({ buffer: markTrackedChanges(buffer) })).value

test("a file without tracked changes is returned unchanged", async () => {
  const clean = docx(`<w:p>${run("The Seller shall indemnify the Buyer for all Losses.")}</w:p>`)
  expect(markTrackedChanges(clean)).toBe(clean)
  const formatting = docx(`<w:p><w:r><w:rPr><w:b/><w:rPrChange w:id="3" w:author="Reviewer"><w:rPr/></w:rPrChange></w:rPr><w:t>Bold now.</w:t></w:r></w:p>`)
  expect(markTrackedChanges(formatting)).toBe(formatting)
  // An inserted paragraph mark or an insertion with no text changes no wording.
  const markOnly = docx(`<w:p><w:pPr><w:rPr><w:ins w:id="4" w:author="Reviewer"/></w:rPr></w:pPr>${run("Same words.")}</w:p>${ins("Reviewer", "<w:r><w:br/></w:r>")}`)
  expect(markTrackedChanges(markOnly)).toBe(markOnly)
})

test("proposed insertions and deletions are labelled with their author, after a summary line", async () => {
  const markup = docx(`<w:p>${run("The Seller shall indemnify the Buyer for ")}${del("Counterparty", deleted("all"))}${ins("Counterparty", run("direct"))}${run(" Losses")}${ins("Counterparty", run(", capped at the Purchase Price"))}${run(".")}</w:p>`)
  expect(await text(markup)).toBe(
    "[This Word file contains 3 pending tracked changes by Counterparty, marked inline below. Pending tracked changes are not accepted text.]\n\n"
    + "The Seller shall indemnify the Buyer for [Tracked deletion by Counterparty: all][Tracked insertion by Counterparty: direct] Losses[Tracked insertion by Counterparty: , capped at the Purchase Price].\n\n",
  )
})

test("one edit split across adjacent revisions reads as one change; different authors stay apart", async () => {
  const split = docx(`<w:p>${run("Notice ")}${ins("Reviewer A", run("within ten"))}<w:bookmarkStart w:id="9" w:name="_x"/><w:proofErr w:type="spellStart"/>${ins("Reviewer A", run(" Business Days"))}<w:bookmarkEnd w:id="9"/>${ins("Reviewer B", run(" in writing"))}${run(".")}</w:p>`)
  const extracted = await text(split)
  expect(extracted).toContain("Notice [Tracked insertion by Reviewer A: within ten Business Days][Tracked insertion by Reviewer B:  in writing].")
  expect(extracted).toStartWith("[This Word file contains 2 pending tracked changes by Reviewer A, Reviewer B,")
})

test("moved text keeps both ends instead of disappearing", async () => {
  const moved = docx(
    `<w:p><w:moveFromRangeStart w:id="5" w:name="move1" w:author="Reviewer"/><w:moveFrom w:id="6" w:author="Reviewer">${run("Clause moved.")}</w:moveFrom><w:moveFromRangeEnd w:id="5"/></w:p>`
    + `<w:p>${run("Middle clause.")}</w:p>`
    + `<w:p><w:moveToRangeStart w:id="7" w:name="move1" w:author="Reviewer"/><w:moveTo w:id="8" w:author="Reviewer">${run("Clause moved.")}</w:moveTo><w:moveToRangeEnd w:id="7"/></w:p>`,
  )
  const extracted = await text(moved)
  expect(extracted).toContain("[Tracked move from here by Reviewer: Clause moved.]\n\nMiddle clause.\n\n[Tracked move to here by Reviewer: Clause moved.]")
})

test("a deleted paragraph and a deleted table row keep their place and are labelled", async () => {
  const removed = docx(
    `<w:p><w:pPr><w:rPr><w:del w:id="10" w:author="Reviewer"/></w:rPr></w:pPr>${del("Reviewer", deleted("Deleted paragraph."))}</w:p>`
    + `<w:p>${run("Next paragraph.")}</w:p>`
    + `<w:tbl><w:tr><w:tc><w:p>${run("Kept row.")}</w:p></w:tc></w:tr>`
    + `<w:tr><w:trPr><w:del w:id="11" w:author="Reviewer"/></w:trPr><w:tc><w:p>${del("Reviewer", deleted("Deleted row."))}</w:p></w:tc></w:tr></w:tbl>`,
  )
  const extracted = await text(removed)
  expect(extracted).toContain("[Tracked deletion by Reviewer: Deleted paragraph.]\n\nNext paragraph.")
  expect(extracted).toContain("Kept row.")
  expect(extracted).toContain("[Tracked deletion by Reviewer: Deleted row.]")
})

test("a deleted row whose text is not itself tracked is still dropped, never read as agreed text", async () => {
  const row = `<w:tbl><w:tr><w:trPr><w:del w:id="12" w:author="Reviewer"/></w:trPr><w:tc><w:p>${run("Untracked row text.")}</w:p></w:tc></w:tr></w:tbl>`
  const alone = docx(row)
  expect(markTrackedChanges(alone)).toBe(alone)
  const extracted = await text(docx(`<w:p>${ins("Reviewer", run("Added."))}</w:p>${row}`))
  expect(extracted).toContain("[Tracked insertion by Reviewer: Added.]")
  expect(extracted).not.toContain("Untracked row text.")
})

test("changes in parts the reader skips are neither labelled nor counted", async () => {
  const field = `<w:fldSimple w:instr=" DATE "><w:del w:id="13" w:author="Reviewer">${deleted("1 January")}</w:del></w:fldSimple>`
  const textBox = `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="wps"><w:r><w:drawing><w:txbxContent><w:p>${ins("Reviewer", run("Box text."))}</w:p></w:txbxContent></w:drawing></w:r></mc:Choice><mc:Fallback/></mc:AlternateContent>`
  const skipped = docx(`<w:p>${run("Dated ")}${field}${textBox}</w:p>`)
  expect(markTrackedChanges(skipped)).toBe(skipped)
  const extracted = await text(docx(`<w:p>${run("Dated ")}${field}${ins("Reviewer", run("today"))}</w:p>`))
  expect(extracted).toStartWith("[This Word file contains 1 pending tracked change by Reviewer,")
  expect(extracted).toContain("Dated [Tracked insertion by Reviewer: today]")
  // A dropped row's changes are unread too; a bare text box is read and labelled.
  const droppedRow = `<w:tbl><w:tr><w:trPr><w:del w:id="14" w:author="Reviewer"/></w:trPr><w:tc><w:p>${run("Plain cell.")}${ins("Reviewer", run("Row insertion."))}</w:p></w:tc></w:tr></w:tbl>`
  const bareBox = `<w:p><w:r><w:drawing><w:txbxContent><w:p>${ins("Reviewer", run("Box text."))}</w:p></w:txbxContent></w:drawing></w:r></w:p>`
  const mixed = await text(docx(`${droppedRow}${bareBox}`))
  expect(mixed).toStartWith("[This Word file contains 1 pending tracked change by Reviewer,")
  expect(mixed).toContain("[Tracked insertion by Reviewer: Box text.]")
  expect(mixed).not.toContain("Row insertion.")
})

test("text inserted by one author and deleted by another carries both labels", async () => {
  const nested = docx(`<w:p>${run("Term ")}${ins("Reviewer A", `${run("one ")}${del("Reviewer B", deleted("two"))}`)}</w:p>`)
  expect(await text(nested)).toContain("Term [Tracked insertion by Reviewer A: one [Tracked deletion by Reviewer B: two]]")
})

test("author names cannot break a label, a missing author is omitted, and comments are counted", async () => {
  const awkward = docx(
    `<w:p>${ins("Ops]: [Team", run("added"))}${del("", deleted("removed"))}</w:p>`,
    '<w:comment w:id="0" w:author="Reviewer"><w:p><w:r><w:t>Query.</w:t></w:r></w:p></w:comment><w:comment w:id="1" w:author="Reviewer"><w:p><w:r><w:t>Second.</w:t></w:r></w:p></w:comment>',
  )
  const extracted = await text(awkward)
  expect(extracted).toContain("[Tracked insertion by Ops Team: added][Tracked deletion: removed]")
  expect(extracted).toStartWith("[This Word file contains 2 pending tracked changes by Ops Team, marked inline below. Pending tracked changes are not accepted text. It also contains 2 comments, which are not shown.]")
})

test("a removed break or symbol without wording stays unread beside a labelled change", async () => {
  const extracted = await text(docx(`<w:p>${run("One")}${del("Reviewer", "<w:r><w:br/></w:r>")}${run("line")}${ins("Reviewer", run("."))}</w:p>`))
  expect(extracted).toContain("Oneline[Tracked insertion by Reviewer: .]")
})

test("the proposed view blanks labels and removed wording at the same offsets", async () => {
  const labelled = await text(docx(
    `<w:p>${run("Notice ")}${del("Reviewer", deleted("within [five] days"))}${ins("Reviewer", run("within [ten] days"))}${run(".")}</w:p>`
    + `<w:p>${ins("Reviewer", run("7.2 Termination"))}</w:p>`,
  ))
  const view = proposedView(labelled)
  expect(view.length).toBe(labelled.length)
  expect(view.split("\n").length).toBe(labelled.split("\n").length)
  expect(view.split("\n").map(line => line.trim().replace(/\s+/g, " ")).filter(Boolean))
    .toEqual(["Notice within [ten] days .", "7.2 Termination"])
  const unpaired = await text(docx(
    `<w:p>${del("Reviewer", deleted("half [open"))}${run(" kept.")}</w:p><w:p>${run("3. Remedies")}</w:p>`,
  ))
  expect(proposedView(unpaired).split("\n").map(line => line.trim()).filter(Boolean).at(-1)).toBe("3. Remedies")
  // Without the summary line a file has no labels, even if its wording looks like one.
  const clean = "[Tracked insertion: 4. Literal Heading] stays as written."
  expect(proposedView(clean)).toBe(clean)
})

test("a Word redline produced by comparison is labelled", async () => {
  const original = (await renderMarkdownToDocx("# Record\n\nThe prior arrangement remains pending.")).bytes
  const revised = (await renderMarkdownToDocx("# Record\n\nThe revised arrangement remains pending.")).bytes
  const redline = Buffer.from(await (await docxodus()).compareDocuments(original, revised, { authorName: "Reviewer" }))
  const extracted = await text(redline)
  expect(extracted).toMatch(/\[Tracked deletion by Reviewer: [^\]]*prior[^\]]*\]/)
  expect(extracted).toMatch(/\[Tracked insertion by Reviewer: [^\]]*revised[^\]]*\]/)
  expect(extracted).toContain("arrangement remains pending.")
})
