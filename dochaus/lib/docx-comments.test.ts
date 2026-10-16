import { beforeAll, describe, expect, test } from "bun:test"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { exportDocxComments, readDocxCommentParagraphs, type DocxComment } from "./docx-comments"
import { docxodus } from "./docxodus"
import { renderMarkdownToDocx } from "./render"

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const R = "http://schemas.openxmlformats.org/package/2006/relationships"
const CT = "http://schemas.openxmlformats.org/package/2006/content-types"
const FINAL = "The recipient may inspect the material within nineteen days."
const ORIGINAL = "The recipient may inspect the material within twelve days."
const UNCHANGED = "The sender must retain a copy of each notice."
const comment = (id: string, anchor: DocxComment["anchor"]): DocxComment => ({ id, anchor,
  text: "Explains the supplied review period.\nSource: instructions — \"nineteen days\".", author: "doc.haus review" })

describe("native Word comment export", () => {
  let clean: Uint8Array
  let redline: Uint8Array
  beforeAll(async () => {
    clean = (await renderMarkdownToDocx(`# Service terms\n\n${FINAL}\n\n${UNCHANGED}`)).bytes
    const original = (await renderMarkdownToDocx(`# Service terms\n\n${ORIGINAL}\n\n${UNCHANGED}`)).bytes
    redline = await (await docxodus()).compareDocuments(original, clean, { authorName: "Original revision author" })
  })

  test("writes native ranges, references, relationships and content type on a real clean DOCX", async () => {
    const before = unzipSync(clean)
    const output = exportDocxComments(clean, [comment("period", { paragraph_text: FINAL })])
    expect(output.unresolved).toEqual([])
    expect(output.attached).toEqual([{ id: "period", comment_id: "0", paragraph_index: 1, paragraph_text: FINAL }])
    const parts = unzipSync(output.bytes)
    const native = validateNative(parts, ["0"])
    expect(native[0].getAttributeNS(W, "author")).toBe("doc.haus review")
    expect(native[0].textContent).toBe(comment("period", {}).text.replace(/\n/g, ""))
    for (const [name, bytes] of Object.entries(before).filter(([name]) =>
      !["word/document.xml", "word/_rels/document.xml.rels", "[Content_Types].xml"].includes(name))) expect(parts[name]).toEqual(bytes)
    const dx = await docxodus()
    const session = dx.openDocxSession(output.bytes, {})
    try {
      expect(session.project().markdown).toContain(FINAL)
      expect(session.project().markdown).toContain("# Comments")
      expect(session.project().markdown).toContain("**doc.haus review**")
      expect(session.grep("^[\\s\\S]*$", { scope: 1, whitespace: 0 }).map(row => row.text).join("\n")).not.toContain("Explains the supplied")
    } finally { session.close() }
    expect(exportDocxComments(clean, [comment("period", { paragraph_text: FINAL })]).bytes).toEqual(output.bytes)
  })

  test("uses accepted redline text and leaves original insertions/deletions byte-identical", async () => {
    const originalXml = xml(redline, "word/document.xml")
    expect(revisions(originalXml).some(text => text.startsWith("<w:ins"))).toBe(true)
    expect(revisions(originalXml).some(text => text.startsWith("<w:del"))).toBe(true)
    const output = exportDocxComments(redline, [comment("period", { paragraph_text: FINAL, paragraph_index: 1 }),
      comment("notice", { quote: "retain a copy" }), comment("deleted", { quote: "twelve days" })])
    expect(output.attached.map(row => row.id)).toEqual(["period", "notice"])
    expect(output.unresolved).toEqual([{ id: "deleted", reason: expect.stringContaining("No accepted paragraph") }])
    expect(revisions(xml(output.bytes, "word/document.xml"))).toEqual(revisions(originalXml))
    validateNative(unzipSync(output.bytes), ["0", "1"])
    const session = (await docxodus()).openDocxSession(output.bytes, {})
    try {
      expect(session.project().markdown).toContain("# Comments")
      expect(session.grep("^[\\s\\S]*$", { scope: 1, whitespace: 0 }).map(row => row.text).join("\n")).not.toContain("Explains the supplied")
    } finally { session.close() }
  })

  test("joins split formatted runs, decodes XML characters and escapes author and multiline explanation", () => {
    const source = withBody(clean, `<w:p><w:r><w:t>A &amp; </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>B &lt; C</w:t></w:r></w:p>`)
    const input = { ...comment("escape", { quote: "& B <" }), author: `Review & \"verification\"`, text: "A & B < C\nSource: \"quoted\" > wording" }
    const output = exportDocxComments(source, [input])
    expect(output.attached[0].paragraph_text).toBe("A & B < C")
    const native = validateNative(unzipSync(output.bytes), ["0"])
    expect(native[0].getAttributeNS(W, "author")).toBe(input.author)
    expect(Array.from(native[0].getElementsByTagNameNS(W, "p")).map(node => node.textContent)).toEqual(input.text.split("\n"))
    expect(xml(output.bytes, "word/document.xml")).toContain("<w:rPr><w:b/></w:rPr>")
  })

  test("requires exact duplicate disambiguation with a stable ordinal or active paragraph pointer", () => {
    const source = withBody(clean, `<w:p><w:r><w:t>Repeated wording.</w:t></w:r></w:p>`
      + `<w:p><w:bookmarkStart w:id="1" w:name="_dh_p123456789abc"/><w:r><w:t>Repeated wording.</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>`)
    const output = exportDocxComments(source, [comment("ambiguous", { paragraph_text: "Repeated wording." }),
      comment("ordinal", { paragraph_text: "Repeated wording.", paragraph_index: 1 }),
      comment("pointer", { quote: "wording", pid: "p123456789abc" }),
      comment("stale-pointer", { paragraph_text: "Repeated wording.", pid: "p000000000000" }),
      comment("nearby", { paragraph_text: "Repeated wording!" })])
    expect(output.attached.map(row => [row.id, row.paragraph_index])).toEqual([["ordinal", 1], ["pointer", 1]])
    expect(output.unresolved.map(row => row.id)).toEqual(["ambiguous", "stale-pointer", "nearby"])
    validateNative(unzipSync(output.bytes), ["0", "1"])
    const doc = parse(xml(output.bytes, "word/document.xml"))
    expect(doc.getElementsByTagNameNS(W, "p")[0].getElementsByTagNameNS(W, "commentRangeStart").length).toBe(0)
  })

  test("preserves existing native comments and relationships and allocates collision-free IDs", () => {
    const first = exportDocxComments(clean, [comment("first", { paragraph_text: UNCHANGED })])
    const before = unzipSync(first.bytes)
    const originalComment = new XMLSerializer().serializeToString(parse(strFromU8(before["word/comments.xml"])).getElementsByTagNameNS(W, "comment")[0])
    const output = exportDocxComments(first.bytes, [comment("second", { paragraph_text: FINAL })])
    const native = validateNative(unzipSync(output.bytes), ["0", "1"])
    expect(new XMLSerializer().serializeToString(native[0])).toBe(originalComment)
    expect(xml(output.bytes, "word/_rels/document.xml.rels")).toBe(strFromU8(before["word/_rels/document.xml.rels"]))
    expect(xml(output.bytes, "[Content_Types].xml")).toBe(strFromU8(before["[Content_Types].xml"]))
    expect(output.attached[0].comment_id).toBe("1")
  })

  test("honors an existing comments relationship with a non-default part name", () => {
    const first = unzipSync(exportDocxComments(clean, [comment("first", { paragraph_text: UNCHANGED })]).bytes)
    first["word/reviewer-comments.xml"] = first["word/comments.xml"]
    delete first["word/comments.xml"]
    first["word/_rels/document.xml.rels"] = strToU8(strFromU8(first["word/_rels/document.xml.rels"]).replace('Target="comments.xml"', 'Target="reviewer-comments.xml"'))
    first["[Content_Types].xml"] = strToU8(strFromU8(first["[Content_Types].xml"]).replace('/word/comments.xml', '/word/reviewer-comments.xml'))
    const output = exportDocxComments(zipSync(first), [comment("second", { paragraph_text: FINAL })])
    const parts = unzipSync(output.bytes)
    expect(parts["word/comments.xml"]).toBeUndefined()
    expect(parse(strFromU8(parts["word/reviewer-comments.xml"])).getElementsByTagNameNS(W, "comment").length).toBe(2)
    expect(output.attached[0].comment_id).toBe("1")
  })

  test("anchors exact table paragraphs without changing cells or nearby duplicate body text", async () => {
    const source = (await renderMarkdownToDocx("Repeated wording.\n\n| Heading | Direction |\n| --- | --- |\n| Notice | Repeated wording. |\n")).bytes
    const body = parse(xml(source, "word/document.xml"))
    const rows = Array.from(body.getElementsByTagNameNS(W, "p"))
      .filter(node => Array.from(node.getElementsByTagNameNS(W, "t")).map(text => text.textContent).join("").length > 0)
    const tableTarget = rows.findIndex(node => node.parentNode?.nodeType === 1 && (node.parentNode as Element).localName === "tc"
      && node.textContent === "Repeated wording.")
    expect(tableTarget).toBeGreaterThan(0)
    const output = exportDocxComments(source, [comment("table", { paragraph_text: "Repeated wording.", paragraph_index: tableTarget })])
    expect(output.attached[0].paragraph_index).toBe(tableTarget)
    const after = parse(xml(output.bytes, "word/document.xml"))
    expect(after.getElementsByTagNameNS(W, "tc").length).toBe(body.getElementsByTagNameNS(W, "tc").length)
    expect(after.getElementsByTagNameNS(W, "p")[0].getElementsByTagNameNS(W, "commentRangeStart").length).toBe(0)
    validateNative(unzipSync(output.bytes), ["0"])
  })

  test("refuses an existing comments relationship without a target", () => {
    const source = unzipSync(exportDocxComments(clean, [comment("first", { paragraph_text: UNCHANGED })]).bytes)
    const relationships = parse(strFromU8(source["word/_rels/document.xml.rels"]))
    const relationship = Array.from(relationships.getElementsByTagNameNS(R, "Relationship"))
      .find(node => node.getAttribute("Type")?.endsWith("/comments"))!
    relationship.removeAttribute("Target")
    source["word/_rels/document.xml.rels"] = strToU8(new XMLSerializer().serializeToString(relationships))
    expect(() => exportDocxComments(zipSync(source), [comment("second", { paragraph_text: FINAL })]))
      .toThrow("invalid_docx_comments_relationship")
  })

  test("excludes deletion-only paragraphs and deleted pointers, and reports deleted paragraph boundaries", () => {
    const source = withBody(clean, `<w:p><w:del w:id="3" w:author="Earlier reviewer"><w:r><w:delText>Gone.</w:delText></w:r></w:del></w:p>`
      + `<w:p><w:del w:id="4"><w:bookmarkStart w:id="9" w:name="_dh_p123456789abc"/></w:del><w:r><w:t>Kept.</w:t></w:r></w:p>`
      + `<w:p><w:pPr><w:rPr><w:del w:id="5"/></w:rPr></w:pPr><w:r><w:t>Merged boundary.</w:t></w:r></w:p>`)
    const output = exportDocxComments(source, [comment("gone", { quote: "Gone." }), comment("kept", { paragraph_text: "Kept.", paragraph_index: 0 }),
      comment("deleted-pointer", { paragraph_text: "Kept.", pid: "p123456789abc" }), comment("boundary", { paragraph_text: "Merged boundary." })])
    expect(output.attached.map(row => row.id)).toEqual(["kept"])
    expect(output.unresolved.map(row => row.id)).toEqual(["gone", "deleted-pointer", "boundary"])
    expect(output.unresolved[2].reason).toContain("boundary is deleted")
    expect(revisions(xml(output.bytes, "word/document.xml"))).toEqual(revisions(xml(source, "word/document.xml")))
  })

  test("shares exact native paragraph positions, separators and active pointers with attachment", () => {
    const source = withBody(clean, `<w:p/><w:p><w:del w:id="1"><w:r><w:delText>Deleted.</w:delText></w:r></w:del></w:p>`
      + `<w:p><w:bookmarkStart w:id="2" w:name="_dh_p123456789abc"/><w:r><w:t>Before</w:t><w:tab/><w:t>tab.</w:t><w:br/><w:t>Next line.</w:t></w:r><w:bookmarkEnd w:id="2"/></w:p>`
      + `<w:p><w:moveFrom w:id="3"><w:bookmarkStart w:id="4" w:name="_dh_p000000000000"/><w:r><w:t>Moved.</w:t></w:r></w:moveFrom><w:r><w:t>Kept.</w:t></w:r></w:p>`)
    expect(readDocxCommentParagraphs(source)).toEqual([
      { paragraph_index: 0, paragraph_text: "Before\ttab.\nNext line.", pids: ["p123456789abc"] },
      { paragraph_index: 1, paragraph_text: "Kept.", pids: [] },
    ])
    const row = readDocxCommentParagraphs(source)[0]
    expect(exportDocxComments(source, [comment("native", { paragraph_index: row.paragraph_index, paragraph_text: row.paragraph_text, pid: row.pids[0] })]).unresolved).toEqual([])
  })

  test("a comment spanning split results refuses a deleted later paragraph boundary", () => {
    const source = withBody(clean, `<w:p><w:r><w:t>First result.</w:t></w:r></w:p>`
      + `<w:p><w:pPr><w:rPr><w:del w:id="8"/></w:rPr></w:pPr><w:r><w:t>Second result.</w:t></w:r></w:p>`)
    const output = exportDocxComments(source, [comment("split", { paragraphs: [
      { paragraph_index: 0, paragraph_text: "First result." }, { paragraph_index: 1, paragraph_text: "Second result." },
    ] })])
    expect(output.attached).toEqual([])
    expect(output.unresolved[0]!.reason).toContain("boundary is deleted")
    expect(output.bytes).toEqual(source)
  })

  test("leaves bytes intact when all anchors are unresolved and refuses missing attribution", () => {
    const unresolved = exportDocxComments(clean, [comment("missing", { paragraph_text: "Absent paragraph." })])
    expect(unresolved.bytes).toEqual(clean)
    expect(unresolved.attached).toEqual([])
    expect(() => exportDocxComments(clean, [{ ...comment("missing", { paragraph_text: FINAL }), author: "" }])).toThrow("invalid_docx_comment")
    expect(() => exportDocxComments(clean, [comment("same", { paragraph_text: FINAL }), comment("same", { paragraph_text: FINAL })])).toThrow("duplicate_docx_comment_id")
  })
})

function parse(xml: string) { return new DOMParser().parseFromString(xml, "application/xml") }
function xml(bytes: Uint8Array, name: string) { return strFromU8(unzipSync(bytes)[name]) }
function revisions(xml: string) { return [...xml.matchAll(/<w:(ins|del)\b[\s\S]*?<\/w:\1>|<w:(?:ins|del)\b[^>]*\/>/g)].map(match => match[0]) }
function withBody(bytes: Uint8Array, body: string) {
  const parts = unzipSync(bytes)
  parts["word/document.xml"] = strToU8(`<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`)
  return zipSync(parts)
}
function validateNative(parts: ReturnType<typeof unzipSync>, expected: string[]) {
  const doc = parse(strFromU8(parts["word/document.xml"]))
  const native = Array.from(parse(strFromU8(parts["word/comments.xml"])).getElementsByTagNameNS(W, "comment"))
  expect(native.map(node => node.getAttributeNS(W, "id"))).toEqual(expected)
  for (const marker of ["commentRangeStart", "commentRangeEnd", "commentReference"])
    expect(Array.from(doc.getElementsByTagNameNS(W, marker)).map(node => node.getAttributeNS(W, "id")).sort()).toEqual(expected)
  for (const name of ["commentRangeStart", "commentRangeEnd", "commentReference"]) for (const node of Array.from(doc.getElementsByTagNameNS(W, name)))
    expect(node.parentNode?.nodeType === 1 && ["ins", "del"].includes((node.parentNode as Element).localName)).toBe(false)
  const relationships = Array.from(parse(strFromU8(parts["word/_rels/document.xml.rels"])).getElementsByTagNameNS(R, "Relationship"))
  expect(relationships.filter(node => node.getAttribute("Type").endsWith("/comments"))).toHaveLength(1)
  expect(relationships.find(node => node.getAttribute("Type").endsWith("/comments"))?.getAttribute("Target")).toBe("comments.xml")
  const overrides = Array.from(parse(strFromU8(parts["[Content_Types].xml"])).getElementsByTagNameNS(CT, "Override"))
  expect(overrides.find(node => node.getAttribute("PartName") === "/word/comments.xml")?.getAttribute("ContentType"))
    .toBe("application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml")
  return native
}
