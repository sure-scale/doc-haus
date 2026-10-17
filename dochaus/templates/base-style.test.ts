// Run from dochaus/: bun test templates/base-style.test.ts
import { describe, expect, test } from "bun:test"
import { readFileSync } from "node:fs"
import path from "node:path"
import { DOMParser } from "@xmldom/xmldom"
import { strFromU8, unzipSync } from "fflate"
import { extractRawText } from "mammoth"
import { renderMarkdownToDocx } from "../lib/render"

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const R = "http://schemas.openxmlformats.org/package/2006/relationships"
const O = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const CT = "http://schemas.openxmlformats.org/package/2006/content-types"

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

  test("rendered markdown lists are real Word numbering, not literal markers", async () => {
    const rendered = await renderMarkdownToDocx("Steps:\n\n1. Give notice.\n2. Wait seven days.\n\nRecords:\n\n- Keep the original.\n- Send a copy.")
    const parts = unzipSync(rendered.bytes)
    const xml = (name: string) => new DOMParser().parseFromString(strFromU8(parts[name]!), "application/xml")
    const paragraphs = Array.from(xml("word/document.xml").getElementsByTagNameNS(W, "p")).map(paragraph => ({
      text: Array.from(paragraph.getElementsByTagNameNS(W, "t")).map(node => node.textContent).join(""),
      numId: paragraph.getElementsByTagNameNS(W, "numPr")[0]?.getElementsByTagNameNS(W, "numId")[0]?.getAttributeNS(W, "val") ?? null,
    }))
    expect(paragraphs.map(row => row.text)).toEqual(["Steps:", "Give notice.", "Wait seven days.", "Records:", "Keep the original.", "Send a copy."])
    const [ordered, bullet] = [paragraphs[1]!.numId, paragraphs[4]!.numId]
    expect(ordered).not.toBeNull()
    expect(bullet).not.toBeNull()
    expect(ordered).not.toBe(bullet)
    expect(paragraphs.map(row => row.numId)).toEqual([null, ordered, ordered, null, bullet, bullet])
    // The body relates and declares one numbering part that defines each list's format.
    const relationship = Array.from(xml("word/_rels/document.xml.rels").getElementsByTagNameNS(R, "Relationship"))
      .find(row => row.getAttribute("Type") === `${O}/numbering`)
    expect(relationship).toBeDefined()
    const part = path.posix.resolve("/word", relationship!.getAttribute("Target")!).slice(1)
    expect(Array.from(xml("[Content_Types].xml").getElementsByTagNameNS(CT, "Override")).map(row => row.getAttribute("PartName"))).toContain(`/${part}`)
    const numbering = xml(part)
    const format = (numId: string) => {
      const abstract = Array.from(numbering.getElementsByTagNameNS(W, "num")).find(row => row.getAttributeNS(W, "numId") === numId)
        ?.getElementsByTagNameNS(W, "abstractNumId")[0]?.getAttributeNS(W, "val")
      const level = Array.from(numbering.getElementsByTagNameNS(W, "abstractNum")).find(row => row.getAttributeNS(W, "abstractNumId") === abstract)
        ?.getElementsByTagNameNS(W, "lvl")[0]
      return level?.getAttributeNS(W, "ilvl") === "0" ? level.getElementsByTagNameNS(W, "numFmt")[0]?.getAttributeNS(W, "val") : undefined
    }
    expect([format(ordered!), format(bullet!)]).toEqual(["decimal", "bullet"])
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
