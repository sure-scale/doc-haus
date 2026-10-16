// Run from dochaus/: bun test templates/base-style.test.ts
import { describe, expect, test } from "bun:test"
import { readFileSync } from "node:fs"
import { strFromU8, unzipSync } from "fflate"
import { extractRawText } from "mammoth"
import { renderMarkdownToDocx } from "../lib/render"

describe("_base.docx styles", () => {
  test("rendered quotations have a declared Word style and retain their text", async () => {
    const content = "Introduction.\n\n> The agreed notice must be given in writing.\n\nConclusion."
    const rendered = await renderMarkdownToDocx(content)
    const parts = unzipSync(rendered.bytes)
    const document = strFromU8(parts["word/document.xml"]!)
    const styles = strFromU8(parts["word/styles.xml"]!)
    expect(document).toContain('<w:pStyle w:val="Quote"')
    const quote = styles.match(/<w:style[^>]*w:styleId="Quote"[\s\S]*?<\/w:style>/)?.[0]
    expect(quote).toContain('<w:basedOn w:val="Normal"')
    expect(quote).toContain('<w:ind w:left="720" w:right="720"')
    const extracted = await extractRawText({ buffer: Buffer.from(rendered.bytes) })
    expect(extracted.messages).toEqual([])
    expect(extracted.value.trim()).toBe("Introduction.\n\nThe agreed notice must be given in writing.\n\nConclusion.")
  }, 120000)

  test("body paragraphs have clear line spacing", () => {
    const styles = strFromU8(unzipSync(readFileSync("templates/_base.docx"))["word/styles.xml"]!)
    expect(defaultParagraphSpacing(styles)).toEqual({ after: 200, line: 276 })
  })

  test("headings have enough vertical spacing for rendered legal drafts", () => {
    const styles = strFromU8(unzipSync(readFileSync("templates/_base.docx"))["word/styles.xml"]!)
    expect(headingSpacing(styles, "Heading1")).toEqual({ before: 320, after: 160 })
    expect(headingSpacing(styles, "Heading2")).toEqual({ before: 280, after: 140 })
    expect(headingSpacing(styles, "Heading3")).toEqual({ before: 220, after: 120 })
  })
})

function headingSpacing(styles: string, styleId: string) {
  const style = styles.match(new RegExp(`<w:style[^>]*w:styleId="${styleId}"[\\s\\S]*?</w:style>`))?.[0] ?? ""
  const spacing = style.match(/<w:spacing[^>]*w:before="(\d+)"[^>]*w:after="(\d+)"/)
  return spacing ? { before: Number(spacing[1]), after: Number(spacing[2]) } : undefined
}

function defaultParagraphSpacing(styles: string) {
  const spacing = styles.match(/<w:pPrDefault><w:pPr><w:spacing[^>]*w:after="(\d+)"[^>]*w:line="(\d+)"/)
  return spacing ? { after: Number(spacing[1]), line: Number(spacing[2]) } : undefined
}
