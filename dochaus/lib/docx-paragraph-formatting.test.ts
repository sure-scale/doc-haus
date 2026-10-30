import { expect, test } from "bun:test"
import { strFromU8, unzipSync } from "fflate"
import { renderMarkdownToDocx } from "./render"
import { applyDocxParagraphFormatting } from "./docx-paragraph-formatting"
import { docxodus } from "./docxodus"

test("full document paragraph formatting preserves package and resolves its tracked change", async () => {
  const source = (await renderMarkdownToDocx("First paragraph.\n\nSecond **important** paragraph.")).bytes
  const input = { paragraphText: "Second important paragraph.", format: { spacingAfter: 6 }, author: "Counsel", tracked: true }
  const output = await applyDocxParagraphFormatting(source, input)
  const before = unzipSync(source), after = unzipSync(output)
  for (const name of Object.keys(before).filter(name => name !== "word/document.xml")) expect(after[name]).toEqual(before[name])
  expect(strFromU8(after["word/document.xml"]!)).toContain("pPrChange")
  const session = (await docxodus()).openDocxSession(output, {})
  try { expect(session.listRevisions().some(r => r.type === "format" && r.author === "Counsel")).toBe(true) } finally { session.close() }
  await expect(applyDocxParagraphFormatting(source, { ...input, paragraphText: "missing" })).rejects.toThrow("missing or ambiguous")
})

test("formats a paragraph beside tracked wording and refuses duplicate exact targets", async () => {
  const source = (await renderMarkdownToDocx("First paragraph.\n\nSecond paragraph.")).bytes
  const { createDocxRevision, applyDocxParagraphEdits } = await import("./docx-amendments")
  const revision = await createDocxRevision(source)
  const paragraph = revision.paragraphs.find(p => p.text === "Second paragraph.")!
  const changed = await applyDocxParagraphEdits(source, [{ id: "edit", paragraph_handle: paragraph.paragraph_handle,
    occurrence: 0, before_quote: "Second", replacement: "Updated" }], undefined, { revisionAuthor: "Counsel", pendingRevisions: "layer", ownAuthors: ["Counsel"] })
  const result = await applyDocxParagraphFormatting(changed.bytes, { paragraphText: "Updated paragraph.", format: { spacingAfter: 6 },
    author: "Counsel", tracked: true, options: { pendingRevisions: "layer", ownAuthors: ["Counsel"] } })
  const session = (await docxodus()).openDocxSession(result, {})
  try { expect(session.listRevisions().map(r => r.type)).toContain("format") } finally { session.close() }
  const duplicate = (await renderMarkdownToDocx("Repeated paragraph.\n\nRepeated paragraph.")).bytes
  await expect(applyDocxParagraphFormatting(duplicate, { paragraphText: "Repeated paragraph.", format: { spacingAfter: 6 },
    author: "Counsel", tracked: true })).rejects.toThrow("missing or ambiguous")
})
