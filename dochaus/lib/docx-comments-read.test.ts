import { expect, test } from "bun:test"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { exportDocxComments, readDocxCommentParagraphs, readDocxComments, readDocxRevisionText, readDocxRevisionProperties } from "./docx-comments"
import { docxodus } from "./docxodus"
import { renderMarkdownToDocx } from "./render"

function replaceBody(bytes: Uint8Array, body: string) {
  const parts = unzipSync(bytes)
  parts["word/document.xml"] = strToU8(strFromU8(parts["word/document.xml"]!).replace(/<w:body>[\s\S]*?<\/w:body>/, `<w:body>${body}</w:body>`))
  return zipSync(parts)
}

test("native exported comments are read separately without altering accepted text or revisions", async () => {
  const original = (await renderMarkdownToDocx("# Record\n\nPrior arrangement remains pending.")).bytes
  const clean = (await renderMarkdownToDocx("# Record\n\nRevised arrangement remains pending.")).bytes
  const redline = await (await docxodus()).compareDocuments(original, clean, { authorName: "Reviewer" })
  expect(readDocxComments(redline)).toEqual([])
  expect(readDocxRevisionText(clean)).toEqual({ inserted_text: [], deleted_text: [] })
  const revisions = readDocxRevisionText(redline)
  expect(revisions.inserted_text.join(" ")).toContain("Revised")
  expect(revisions.deleted_text.join(" ")).toContain("Prior")
  const row = readDocxCommentParagraphs(redline).find(p => p.paragraph_text.includes("Revised"))!
  const output = exportDocxComments(redline, [{ id: "explanation", author: "Reviewer", text: "A qualified explanation.\nThis is commentary, not an agreement.", anchor: { paragraph_text: row.paragraph_text, paragraph_index: row.paragraph_index } }])
  expect(output.unresolved).toEqual([])
  const comments = readDocxComments(output.bytes)
  expect(comments).toHaveLength(1)
  expect(comments[0].text).toBe("A qualified explanation.\nThis is commentary, not an agreement.")
  expect(comments[0].markers.map(mark => mark.kind)).toEqual(["commentRangeStart", "commentRangeEnd", "commentReference"])
  expect(comments[0].markers.every(mark => mark.paragraph_index === row.paragraph_index && mark.accepted_paragraph_text === row.paragraph_text)).toBe(true)
  expect(readDocxCommentParagraphs(output.bytes)).toEqual(readDocxCommentParagraphs(redline))
  expect(readDocxRevisionText(output.bytes)).toEqual(revisions)
})

test("deleted-only comment markers retain artifact PIDs and XML ordinals without inventing accepted ordinals", async () => {
  const bytes = (await renderMarkdownToDocx("Placeholder.")).bytes
  const redline = replaceBody(bytes,
    '<w:p><w:bookmarkStart w:id="1" w:name="_dh_p111111111111"/><w:del w:id="7"><w:r><w:delText>Removed arrangement.</w:delText></w:r></w:del><w:bookmarkEnd w:id="1"/></w:p>' +
    '<w:p><w:bookmarkStart w:id="2" w:name="_dh_p222222222222"/><w:r><w:t>Surviving arrangement.</w:t></w:r><w:bookmarkEnd w:id="2"/></w:p>')
  const output = exportDocxComments(redline, [{ id: "removed", author: "Reviewer", text: "Why it was removed.", anchor: { deleted_paragraph_text: "Removed arrangement." } }])
  expect(output.unresolved).toEqual([])
  expect(readDocxCommentParagraphs(output.bytes)).toEqual([{ paragraph_index: 0, paragraph_text: "Surviving arrangement.", pids: ["p222222222222"] }])
  for (const marker of readDocxComments(output.bytes)[0].markers) {
    expect(marker.paragraph_index).toBeNull()
    expect(marker.xml_paragraph_index).toBe(0)
    expect(marker.pids).toEqual(["p111111111111"])
    expect(marker.accepted_paragraph_text).toBe("")
    expect(marker.deleted_paragraph_text).toBe("Removed arrangement.")
  }
})

test("tracked text preserves tabs and breaks and reads move-from text without XML indentation", async () => {
  const base = (await renderMarkdownToDocx("Placeholder.")).bytes
  const revised = replaceBody(base, `<w:p>
    <w:ins w:id="1"><w:r><w:t>Added</w:t><w:tab/><w:t>terms</w:t><w:br/><w:t>today.</w:t></w:r></w:ins>
    <w:del w:id="2"><w:r><w:delText>Removed</w:delText><w:tab/><w:delText>terms.</w:delText></w:r></w:del>
    <w:moveFrom w:id="3"><w:r><w:t>Moved</w:t><w:br/><w:t>away.</w:t></w:r></w:moveFrom>
  </w:p>`)
  expect(readDocxRevisionText(base)).toEqual({ inserted_text: [], deleted_text: [] })
  expect(readDocxRevisionText(revised)).toEqual({ inserted_text: ["Added\tterms\ntoday."], deleted_text: ["Removed\tterms.", "Moved\naway."] })
  const movedOnly = replaceBody(base, '<w:p><w:moveFrom w:id="3"><w:r><w:t>Moved</w:t><w:br/><w:t>away.</w:t></w:r></w:moveFrom></w:p>')
  const output = exportDocxComments(movedOnly, [{ id: "moved", author: "Reviewer", text: "Moved explanation.", anchor: { deleted_paragraph_text: "Moved\naway." } }])
  expect(output.unresolved).toEqual([])
  expect(readDocxComments(output.bytes)[0].markers[0].deleted_paragraph_text).toBe("Moved\naway.")
})


test("revision properties count native main-body elements, not text labels, and clean artifacts report zero", async () => {
  const base = (await renderMarkdownToDocx("The words ins del and rPrChange are not tracked revisions.")).bytes
  expect(Object.values(readDocxRevisionProperties(base).revision_elements).every(count => count === 0)).toBe(true)
  const bytes = replaceBody(base, `<w:p><w:pPr><w:pPrChange w:id="1"><w:pPr/></w:pPrChange></w:pPr>
    <w:r><w:rPr><w:rPrChange w:id="2"><w:rPr/></w:rPrChange></w:rPr><w:t>Unchanged</w:t></w:r>
    <w:ins w:id="3"><w:r><w:t>Added</w:t></w:r></w:ins>
    <w:del w:id="4"><w:r><w:delText>Removed</w:delText></w:r></w:del>
    <w:moveFrom w:id="5"><w:r><w:t>Moved from</w:t></w:r></w:moveFrom>
    <w:moveTo w:id="6"><w:r><w:t>Moved to</w:t></w:r></w:moveTo>
  </w:p>`)
  expect(readDocxRevisionProperties(bytes)).toEqual({ part: "word/document.xml", revision_elements: {
    ins: 1, del: 1, moveFrom: 1, moveTo: 1, rPrChange: 1, pPrChange: 1,
    tblPrChange: 0, trPrChange: 0, tcPrChange: 0, sectPrChange: 0, tblGridChange: 0, numberingChange: 0,
  } })
})
