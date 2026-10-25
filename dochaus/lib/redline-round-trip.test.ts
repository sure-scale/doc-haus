import { expect, test } from "bun:test"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { docxodus } from "./docxodus"
import { readDocxBodyView, discloseSourceRevisions, sourceRevisionDisclosure } from "./docx-revision-views"
import { readDocxComments } from "./docx-comments"
import { renderMarkdownToDocx } from "./render"
import { assertRedlineRoundTrip, redlineRoundTripMismatch } from "./redline-round-trip"

test("one redline check accepts a real comparison and identifies either wrong projection without comments", async () => {
  const source = (await renderMarkdownToDocx("Original wording.")).bytes
  const intended = (await renderMarkdownToDocx("Revised wording.")).bytes
  const other = (await renderMarkdownToDocx("Different wording.")).bytes
  const redline = await (await docxodus()).compareDocuments(source, intended, { authorName: "Reviewer" })
  expect(await redlineRoundTripMismatch(redline, source, intended)).toBeUndefined()
  expect(await redlineRoundTripMismatch(redline, other, intended)).toMatchObject({ view: "baseline", scope: "body", paragraph: 1 })
  expect(await redlineRoundTripMismatch(redline, source, other)).toMatchObject({ view: "accepted", scope: "body", paragraph: 1 })
  await expect(assertRedlineRoundTrip(redline, source, other)).rejects.toThrow("accept-all text differs at body paragraph 1")
})

test("native projections resolve nested text revisions and moves; markup preserves author and proposed/deleted text", async () => {
  const bytes = await revisedFixture()
  const accepted = (await readDocxBodyView(bytes, "current_body")).map(row => row.paragraph_text)
  const baseline = (await readDocxBodyView(bytes, "baseline_body")).map(row => row.paragraph_text)
  expect(accepted).toEqual(["Start Added", "Moved text", "Inserted row wording"])
  expect(baseline).toEqual(["Start Old", "Moved text", "Deleted row wording"])
  const markup = (await readDocxBodyView(bytes, "marked_up_body")).map(row => row.paragraph_text).join("\n")
  expect(markup).toContain('insert by "Reviewer A": Added⟦delete by "Reviewer B": Later deleted⟧')
  expect(markup).toContain('move from by "Reviewer A": Moved text')
  expect(markup).toContain('move to by "Reviewer A": Moved text')
  expect(markup).toContain('delete row by "Reviewer C": Deleted row wording')
  expect(markup).toContain('insert row by "Reviewer C": Inserted row wording')
  expect((await sourceRevisionDisclosure(bytes))?.authors).toEqual(expect.arrayContaining(["Reviewer A", "Reviewer B", "Reviewer C"]))
  // The structural projection above is separate from comment attachment:
  // replacing its whole table can leave no supported comment anchor. Exercise
  // normal source-revision disclosure on a paragraph-preserving comparison.
  const dx = await docxodus()
  const pending = await dx.compareDocuments((await renderMarkdownToDocx("Earlier wording.")).bytes,
    (await renderMarkdownToDocx("Proposed wording.")).bytes, { authorName: "Reviewer A" })
  const compared = await dx.compareDocuments(pending, (await renderMarkdownToDocx("Final wording.")).bytes, { authorName: "Our reviewer" })
  const disclosed = await discloseSourceRevisions(compared, pending, "Our reviewer")
  expect(disclosed.disclosure?.authors).toEqual(["Reviewer A"])
  expect(readDocxComments(disclosed.bytes).some(row => row.text.includes("not retained as pending changes"))).toBe(true)
  expect(strFromU8(unzipSync(bytes)["word/document.xml"]!)).toContain("Later deleted")
})

async function revisedFixture() {
  const parts = unzipSync((await renderMarkdownToDocx("Seed.")).bytes)
  const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  parts["word/document.xml"] = strToU8(`<w:document xmlns:w="${W}"><w:body>
    <w:p><w:r><w:t xml:space="preserve">Start </w:t></w:r><w:ins w:id="1" w:author="Reviewer A"><w:r><w:t>Added</w:t></w:r><w:del w:id="2" w:author="Reviewer B"><w:r><w:delText>Later deleted</w:delText></w:r></w:del></w:ins><w:del w:id="3" w:author="Reviewer A"><w:r><w:delText>Old</w:delText></w:r></w:del></w:p>
    <w:p><w:moveFrom w:id="4" w:author="Reviewer A"><w:r><w:delText>Moved text</w:delText></w:r></w:moveFrom></w:p>
    <w:p><w:moveTo w:id="5" w:author="Reviewer A"><w:r><w:t>Moved text</w:t></w:r></w:moveTo></w:p>
    <w:tbl><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid>
      <w:tr><w:trPr><w:del w:id="6" w:author="Reviewer C"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Deleted row wording</w:t></w:r></w:p></w:tc></w:tr>
      <w:tr><w:trPr><w:ins w:id="7" w:author="Reviewer C"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Inserted row wording</w:t></w:r></w:p></w:tc></w:tr>
    </w:tbl>
  </w:body></w:document>`)
  return zipSync(parts)
}
