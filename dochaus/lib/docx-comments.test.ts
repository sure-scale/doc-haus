import { beforeAll, describe, expect, test } from "bun:test"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { exportDocxComments, preserveDocxComments, readDocxCommentParagraphs, readDocxComments, readDocxRevisionProperties, readDocxRevisionText, type DocxComment } from "./docx-comments"
import { compareDocxWithPointers, withParagraphPointers, withStableRevisionDates } from "./drafting-amendment"
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

  test("a comparison that carries an existing comment exports it once, unchanged, beside a review comment", async () => {
    // Synthetic source: the unchanged paragraph carries a dated reviewer comment.
    const commented = unzipSync(exportDocxComments(withParagraphPointers((await renderMarkdownToDocx(`# Service terms\n\n${ORIGINAL}\n\n${UNCHANGED}`)).bytes),
      [{ id: "existing", anchor: { paragraph_text: UNCHANGED }, author: "Counterparty counsel", text: "Confirm the retention period." }]).bytes)
    commented["word/comments.xml"] = strToU8(strFromU8(commented["word/comments.xml"]).replace("<w:comment ", '<w:comment w:date="2026-02-03T04:05:06Z" '))
    const original = zipSync(commented)
    const existing = parse(strFromU8(commented["word/comments.xml"])).getElementsByTagNameNS(W, "comment")[0]
    expect([existing.getAttributeNS(W, "author"), existing.getAttributeNS(W, "date"), existing.textContent])
      .toEqual(["Counterparty counsel", "2026-02-03T04:05:06Z", "Confirm the retention period."])
    // A replayed proposal edits the other paragraph and leaves the comment in place.
    const edited = unzipSync(original)
    edited["word/document.xml"] = strToU8(strFromU8(edited["word/document.xml"]).replace("twelve days", "nineteen days"))
    const modified = zipSync(edited)
    const compared = withStableRevisionDates(await compareDocxWithPointers(original, modified, { authorName: "doc.haus review" }), "2026-03-04T05:06:07Z")
    // Premise: the comparator itself carries the existing comment into its output.
    expect(readDocxComments(compared)).toHaveLength(1)
    const preserved = preserveDocxComments(original, compared)
    const paragraphs = readDocxCommentParagraphs(modified)
    const output = exportDocxComments(preserved, [comment("period", { paragraph_index: paragraphs[1].paragraph_index, paragraph_text: FINAL })])
    expect(output.unresolved).toEqual([])
    expect(output.attached).toEqual([{ id: "period", comment_id: "1", paragraph_index: 1, paragraph_text: FINAL }])
    const parts = unzipSync(output.bytes)
    const native = validateNative(parts, ["0", "1"])
    expect(new XMLSerializer().serializeToString(native[0])).toBe(new XMLSerializer().serializeToString(existing))
    expect([native[1].getAttributeNS(W, "author"), native[1].textContent]).toEqual(["doc.haus review", comment("period", {}).text.replace(/\n/g, "")])
    expect(readDocxComments(output.bytes).map(row => [row.comment_id, [...new Set(row.markers.map(marker => marker.paragraph_index))]]))
      .toEqual([["0", [2]], ["1", [1]]])
    // Every paragraph keeps the pointer it had before comparison.
    expect(readDocxCommentParagraphs(output.bytes)).toEqual(paragraphs)
    expect(paragraphs.every(row => row.pids.length === 1)).toBe(true)
    const tracked = revisions(xml(output.bytes, "word/document.xml"))
    expect(tracked.some(text => text.startsWith("<w:ins") && text.includes("nineteen"))).toBe(true)
    expect(tracked.some(text => text.startsWith("<w:del") && text.includes("twelve"))).toBe(true)
    expect(new Set(tracked.map(text => text.match(/ w:date="([^"]*)"/)?.[1]))).toEqual(new Set(["2026-03-04T05:06:07Z"]))
  }, 120000)

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

  test("excludes deletion-only paragraphs and deleted pointers, anchors a deleted paragraph boundary as the merged paragraph and refuses an unjoined one", () => {
    const source = withBody(clean, `<w:p><w:del w:id="3" w:author="Earlier reviewer"><w:r><w:delText>Gone.</w:delText></w:r></w:del></w:p>`
      + `<w:p><w:del w:id="4"><w:bookmarkStart w:id="9" w:name="_dh_p123456789abc"/></w:del><w:r><w:t>Kept.</w:t></w:r></w:p>`
      + `<w:p><w:pPr><w:rPr><w:del w:id="5"/></w:rPr></w:pPr><w:r><w:t xml:space="preserve">Merged </w:t></w:r></w:p>`
      + `<w:p><w:r><w:t>boundary.</w:t></w:r></w:p>`
      + `<w:p><w:pPr><w:rPr><w:del w:id="6"/></w:rPr></w:pPr><w:r><w:t>Unjoined boundary.</w:t></w:r></w:p>`)
    expect(readDocxCommentParagraphs(source).map(row => row.paragraph_text)).toEqual(["Kept.", "Merged boundary.", "Unjoined boundary."])
    const output = exportDocxComments(source, [comment("gone", { quote: "Gone." }), comment("kept", { paragraph_text: "Kept.", paragraph_index: 0 }),
      comment("deleted-pointer", { paragraph_text: "Kept.", pid: "p123456789abc" }), comment("fragment", { paragraph_text: "Merged " }),
      comment("boundary", { paragraph_text: "Merged boundary.", paragraph_index: 1 }), comment("unjoined", { paragraph_text: "Unjoined boundary." })])
    expect(output.attached.map(row => [row.id, row.paragraph_index])).toEqual([["kept", 0], ["boundary", 1]])
    expect(output.unresolved.map(row => row.id)).toEqual(["gone", "deleted-pointer", "fragment", "unjoined"])
    expect(output.unresolved[3].reason).toContain("boundary is deleted")
    const body = Array.from(parse(xml(output.bytes, "word/document.xml")).getElementsByTagNameNS(W, "p"))
    expect(body.map(node => ["commentRangeStart", "commentRangeEnd", "commentReference"].map(name => node.getElementsByTagNameNS(W, name).length)))
      .toEqual([[0, 0, 0], [1, 1, 1], [1, 0, 0], [0, 1, 1], [0, 0, 0]])
    expect(readDocxComments(output.bytes).map(row => row.markers.map(marker => marker.paragraph_index))).toEqual([[0, 0, 0], [1, 1, 1]])
    expect(revisions(xml(output.bytes, "word/document.xml"))).toEqual(revisions(xml(source, "word/document.xml")))
  })

  test("numbers a comparator-split paragraph as its accepted paragraph and attaches comments on it and after it", () => {
    // Synthetic shape of a comparison that splits one final paragraph across two
    // w:p: the first carries a deleted paragraph mark, a deleted paragraph follows.
    const source = withBody(clean, `<w:p><w:r><w:t>4.1 Release applies where that condition is:</w:t></w:r></w:p>`
      + `<w:p><w:pPr><w:rPr><w:del w:id="1" w:author="Reviewer"/></w:rPr></w:pPr><w:ins w:id="2"><w:bookmarkStart w:id="20" w:name="_dh_p111111111111"/></w:ins>`
      + `<w:ins w:id="3"><w:r><w:t>(a) the</w:t></w:r></w:ins><w:del w:id="4"><w:r><w:delText>if:</w:delText></w:r></w:del></w:p>`
      + `<w:p><w:pPr><w:rPr><w:ins w:id="5"/></w:rPr></w:pPr><w:bookmarkStart w:id="21" w:name="_dh_p222222222222"/>`
      + `<w:del w:id="6"><w:r><w:delText>(a) the</w:delText></w:r></w:del><w:r><w:t xml:space="preserve"> Tenant has given notice.</w:t></w:r></w:p>`
      + `<w:p><w:pPr><w:rPr><w:del w:id="7"/></w:rPr></w:pPr><w:del w:id="8"><w:r><w:delText>(b) Removed condition.</w:delText></w:r></w:del></w:p>`
      + `<w:p><w:r><w:t>Later.</w:t></w:r></w:p>`)
    const accepted = withBody(clean, ["4.1 Release applies where that condition is:", "(a) the Tenant has given notice.", "Later."]
      .map(text => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`).join(""))
    const rows = readDocxCommentParagraphs(source)
    expect(rows.map(row => [row.paragraph_index, row.paragraph_text]))
      .toEqual(readDocxCommentParagraphs(accepted).map(row => [row.paragraph_index, row.paragraph_text]))
    expect(rows[1].pids).toEqual(["p111111111111", "p222222222222"])
    const later = readDocxCommentParagraphs(accepted).find(row => row.paragraph_text === "Later.")!
    const output = exportDocxComments(source, [comment("later", { paragraph_index: later.paragraph_index, paragraph_text: later.paragraph_text }),
      comment("joined", { paragraph_index: 1, paragraph_text: "(a) the Tenant has given notice.", pid: "p222222222222" }),
      comment("fragment", { paragraph_text: "(a) the" }), comment("removed", { deleted_paragraph_text: "(b) Removed condition." })])
    expect(output.attached).toEqual([
      { id: "later", comment_id: "0", paragraph_index: 2, paragraph_text: "Later." },
      { id: "joined", comment_id: "1", paragraph_index: 1, paragraph_text: "(a) the Tenant has given notice." },
      { id: "removed", comment_id: "2", paragraph_index: 3, paragraph_text: "(b) Removed condition.", anchor_kind: "deleted" },
    ])
    expect(output.unresolved).toEqual([{ id: "fragment", reason: expect.stringContaining("No accepted paragraph") }])
    validateNative(unzipSync(output.bytes), ["0", "1", "2"])
    const body = Array.from(parse(xml(output.bytes, "word/document.xml")).getElementsByTagNameNS(W, "p"))
    const ids = (node: Element, name: string) => Array.from(node.getElementsByTagNameNS(W, name)).map(marker => marker.getAttributeNS(W, "id"))
    expect(body.map(node => [ids(node, "commentRangeStart"), ids(node, "commentRangeEnd"), ids(node, "commentReference")])).toEqual([
      [[], [], []], [["1"], [], []], [[], ["1"], ["1"]], [["2"], ["2"], ["2"]], [["0"], ["0"], ["0"]],
    ])
    // The deleted-paragraph explanation keeps its own exact deleted anchor, outside the accepted paragraph it merges into.
    expect(readDocxComments(output.bytes).map(row => [row.comment_id, row.markers.map(marker => marker.paragraph_index)]))
      .toEqual([["0", [2, 2, 2]], ["1", [1, 1, 1]], ["2", [null, null, null]]])
    expect(readDocxComments(output.bytes)[1].markers.map(marker => [marker.xml_paragraph_index, marker.accepted_paragraph_text, marker.pids]))
      .toEqual([[1, "(a) the Tenant has given notice.", ["p111111111111", "p222222222222"]], [2, "(a) the Tenant has given notice.", ["p111111111111", "p222222222222"]],
        [2, "(a) the Tenant has given notice.", ["p111111111111", "p222222222222"]]])
    expect(readDocxCommentParagraphs(output.bytes)).toEqual(rows)
    expect(readDocxRevisionText(output.bytes)).toEqual(readDocxRevisionText(source))
    expect(readDocxRevisionProperties(output.bytes)).toEqual(readDocxRevisionProperties(source))
  })

  test("restores an existing comment after a comparator-split paragraph at its accepted position", () => {
    const paragraph = (text: string) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`
    const original = exportDocxComments(withBody(clean, ["The first term applies.", "Retained.", "Commented."].map(paragraph).join("")),
      [{ id: "existing", anchor: { paragraph_text: "Commented." }, author: "Counterparty counsel", text: "Confirm this term." }]).bytes
    const compared = withBody(clean, `<w:p><w:pPr><w:rPr><w:del w:id="1"/></w:rPr></w:pPr><w:r><w:t xml:space="preserve">The first </w:t></w:r></w:p>`
      + `<w:p><w:r><w:t>term applies.</w:t></w:r></w:p>` + paragraph("Retained.") + paragraph("Commented."))
    const preserved = preserveDocxComments(original, compared)
    expect(readDocxComments(preserved).map(row => [row.text, row.markers.map(marker => marker.paragraph_index)])).toEqual([["Confirm this term.", [2, 2, 2]]])
    expect(readDocxCommentParagraphs(preserved).map(row => row.paragraph_text)).toEqual(["The first term applies.", "Retained.", "Commented."])
    expect(revisions(xml(preserved, "word/document.xml"))).toEqual(revisions(xml(compared, "word/document.xml")))
  })

  test("refuses to restore an existing comment whose original paragraph has a pending deleted paragraph mark", async () => {
    const original = exportDocxComments(withBody(clean, `<w:p><w:pPr><w:rPr><w:del w:id="90" w:author="Earlier reviewer"/></w:rPr></w:pPr>`
      + `<w:r><w:t xml:space="preserve">Alpha </w:t></w:r></w:p><w:p><w:r><w:t>Beta.</w:t></w:r></w:p><w:p><w:r><w:t>Notice is five days.</w:t></w:r></w:p>`),
      [{ id: "existing", anchor: { paragraph_text: "Alpha Beta." }, author: "Counterparty counsel", text: "Confirm this term." }]).bytes
    const modified = withBody(clean, `<w:p><w:r><w:t>Alpha Beta.</w:t></w:r></w:p><w:p><w:r><w:t>Notice is seven days.</w:t></w:r></w:p>`)
    const compared = await compareDocxWithPointers(original, modified, { authorName: "doc.haus review" })
    // Premise: the comparator accepts the original's pending revision, so both sides number two accepted paragraphs.
    expect(xml(compared, "word/document.xml")).not.toContain("Earlier reviewer")
    expect(readDocxCommentParagraphs(original).map(row => row.paragraph_text)).toEqual(["Alpha Beta.", "Notice is five days."])
    expect(readDocxCommentParagraphs(compared).map(row => row.paragraph_text)).toEqual(["Alpha Beta.", "Notice is seven days."])
    // Restoring would carry the pending deletion into the redline as this comparison's tracked change.
    expect(() => preserveDocxComments(original, compared)).toThrow("original_comment_range_was_edited")
  }, 120000)

  test("joins a deleted paragraph mark across body-level range, proofing and permission markup", () => {
    const markups = [`<w:bookmarkEnd w:id="50"/>`, `<w:commentRangeStart w:id="51"/>`, `<w:moveToRangeEnd w:id="52"/>`, `<w:proofErr w:type="gramEnd"/>`, `<w:permStart w:id="53"/>`]
    const sources = markups.map(markup => withBody(clean, `<w:p><w:pPr><w:rPr><w:del w:id="1"/></w:rPr></w:pPr><w:r><w:t xml:space="preserve">Alpha </w:t></w:r></w:p>`
      + `${markup}<w:p><w:r><w:t>Beta.</w:t></w:r></w:p><w:p><w:r><w:t>Later.</w:t></w:r></w:p>`))
    expect(sources.map(source => readDocxCommentParagraphs(source).map(row => [row.paragraph_index, row.paragraph_text])))
      .toEqual(markups.map(() => [[0, "Alpha Beta."], [1, "Later."]]))
    expect(sources.map(source => {
      const output = exportDocxComments(source, [comment("joined", { paragraph_index: 0, paragraph_text: "Alpha Beta." }), comment("later", { paragraph_index: 1, paragraph_text: "Later." })])
      return [output.attached.map(row => [row.id, row.paragraph_index]), output.unresolved]
    })).toEqual(markups.map(() => [[["joined", 0], ["later", 1]], []]))
  })

  test("keeps deleted-paragraph comments and pointers on members without accepted text outside the joined paragraph", () => {
    const mark = `<w:pPr><w:rPr><w:del w:id="1"/></w:rPr></w:pPr>`
    // A deletion-only middle member, then a deletion-only last member whose paragraph mark ends the joined paragraph.
    const source = withBody(clean, `<w:p>${mark}<w:bookmarkStart w:id="11" w:name="_dh_p111111111111"/><w:r><w:t xml:space="preserve">Alpha </w:t></w:r></w:p>`
      + `<w:p>${mark}<w:bookmarkStart w:id="12" w:name="_dh_p222222222222"/><w:del w:id="2"><w:r><w:delText>Removed middle.</w:delText></w:r></w:del></w:p>`
      + `<w:p><w:bookmarkStart w:id="13" w:name="_dh_p333333333333"/><w:r><w:t>Beta.</w:t></w:r></w:p>`
      + `<w:p>${mark}<w:r><w:t>Gamma.</w:t></w:r></w:p><w:p><w:del w:id="3"><w:r><w:delText>Removed last.</w:delText></w:r></w:del></w:p>`)
    const rows = readDocxCommentParagraphs(source)
    expect(rows).toEqual([{ paragraph_index: 0, paragraph_text: "Alpha Beta.", pids: ["p111111111111", "p333333333333"] },
      { paragraph_index: 1, paragraph_text: "Gamma.", pids: [] }])
    const output = exportDocxComments(source, [comment("middle", { deleted_paragraph_text: "Removed middle." }),
      comment("last", { deleted_paragraph_text: "Removed last." }), comment("joined", { paragraph_index: 0, paragraph_text: "Alpha Beta." }),
      comment("gamma", { paragraph_index: 1, paragraph_text: "Gamma." })])
    expect(output.unresolved).toEqual([])
    expect(output.attached).toEqual([
      { id: "middle", comment_id: "0", paragraph_index: 1, paragraph_text: "Removed middle.", anchor_kind: "deleted" },
      { id: "last", comment_id: "1", paragraph_index: 4, paragraph_text: "Removed last.", anchor_kind: "deleted" },
      { id: "joined", comment_id: "2", paragraph_index: 0, paragraph_text: "Alpha Beta." },
      { id: "gamma", comment_id: "3", paragraph_index: 1, paragraph_text: "Gamma." },
    ])
    validateNative(unzipSync(output.bytes), ["0", "1", "2", "3"])
    const joined = (xmlIndex: number) => [0, xmlIndex, "Alpha Beta.", "", ["p111111111111", "p333333333333"]]
    expect(readDocxComments(output.bytes).map(row => [row.comment_id, row.markers.map(marker =>
      [marker.paragraph_index, marker.xml_paragraph_index, marker.accepted_paragraph_text, marker.deleted_paragraph_text, marker.pids])])).toEqual([
      ["0", Array(3).fill([null, 1, "", "Removed middle.", ["p222222222222"]])],
      ["1", Array(3).fill([null, 4, "", "Removed last.", []])],
      ["2", [joined(0), joined(2), joined(2)]],
      ["3", Array(3).fill([1, 3, "Gamma.", "", []])],
    ])
    expect(readDocxCommentParagraphs(output.bytes)).toEqual(rows)
    expect(readDocxRevisionText(output.bytes)).toEqual(readDocxRevisionText(source))
    expect(readDocxRevisionProperties(output.bytes)).toEqual(readDocxRevisionProperties(source))
  })

  test("covers joined members with live content but no text in the joined paragraph's range and pointers", () => {
    const mark = `<w:pPr><w:rPr><w:del w:id="1"/></w:rPr></w:pPr>`
    // A leading and a trailing chain member each hold only a drawing, which survives acceptance.
    const source = withBody(clean, `<w:p>${mark}<w:r><w:drawing/></w:r></w:p><w:p>${mark}<w:r><w:t>Figure caption.</w:t></w:r></w:p>`
      + `<w:p><w:bookmarkStart w:id="14" w:name="_dh_p444444444444"/><w:r><w:drawing/></w:r></w:p>`)
    const rows = readDocxCommentParagraphs(source)
    expect(rows).toEqual([{ paragraph_index: 0, paragraph_text: "Figure caption.", pids: ["p444444444444"] }])
    const output = exportDocxComments(source, [comment("figure", { paragraph_index: 0, paragraph_text: "Figure caption." })])
    expect(output.attached).toEqual([{ id: "figure", comment_id: "0", paragraph_index: 0, paragraph_text: "Figure caption." }])
    const body = Array.from(parse(xml(output.bytes, "word/document.xml")).getElementsByTagNameNS(W, "p"))
    expect(body.map(node => ["commentRangeStart", "commentRangeEnd", "commentReference"].map(name => node.getElementsByTagNameNS(W, name).length)))
      .toEqual([[1, 0, 0], [0, 0, 0], [0, 1, 1]])
    expect(readDocxComments(output.bytes)[0].markers.map(marker => [marker.paragraph_index, marker.xml_paragraph_index, marker.accepted_paragraph_text, marker.pids]))
      .toEqual([[0, 0, "Figure caption.", ["p444444444444"]], [0, 2, "Figure caption.", ["p444444444444"]], [0, 2, "Figure caption.", ["p444444444444"]]])
    expect(readDocxCommentParagraphs(output.bytes)).toEqual(rows)
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

  test("a comment spanning split results covers a deleted later paragraph boundary through its merged paragraph and refuses an unjoined one", () => {
    const range = comment("split", { paragraphs: [
      { paragraph_index: 0, paragraph_text: "First result." }, { paragraph_index: 1, paragraph_text: "Second result." },
    ] })
    const joined = withBody(clean, `<w:p><w:r><w:t>First result.</w:t></w:r></w:p>`
      + `<w:p><w:pPr><w:rPr><w:del w:id="8"/></w:rPr></w:pPr><w:r><w:t xml:space="preserve">Second </w:t></w:r></w:p><w:p><w:r><w:t>result.</w:t></w:r></w:p>`)
    const spanning = exportDocxComments(joined, [range])
    expect(spanning.unresolved).toEqual([])
    expect(spanning.attached).toEqual([{ id: "split", comment_id: "0", paragraph_index: 0, paragraph_text: "First result.", paragraphs: range.anchor.paragraphs }])
    const body = Array.from(parse(xml(spanning.bytes, "word/document.xml")).getElementsByTagNameNS(W, "p"))
    expect(body.map(node => ["commentRangeStart", "commentRangeEnd", "commentReference"].map(name => node.getElementsByTagNameNS(W, name).length)))
      .toEqual([[1, 0, 0], [0, 0, 0], [0, 1, 1]])
    const unjoined = withBody(clean, `<w:p><w:r><w:t>First result.</w:t></w:r></w:p>`
      + `<w:p><w:pPr><w:rPr><w:del w:id="8"/></w:rPr></w:pPr><w:r><w:t>Second result.</w:t></w:r></w:p>`)
    const output = exportDocxComments(unjoined, [range])
    expect(output.attached).toEqual([])
    expect(output.unresolved[0]!.reason).toContain("boundary is deleted")
    expect(output.bytes).toEqual(unjoined)
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
