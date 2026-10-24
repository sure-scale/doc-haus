// Run from dochaus/: bun test lib/docx-polish.test.ts
import { describe, expect, test } from "bun:test"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { polishMarkdownArtifactsInDocx } from "./docx-polish"
import { renderMarkdownToDocx } from "./render"

describe("polishMarkdownArtifactsInDocx", () => {
  test("declares only missing inline-code styles and preserves existing style inheritance", () => {
    const document = strToU8('<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:rPr><w:rStyle w:val="Code"/></w:rPr><w:t>reference</w:t></w:r></w:p></w:body></w:document>')
    const styles = '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="Normal"><w:rPr><w:rFonts w:ascii="Calibri"/></w:rPr></w:style></w:styles>'
    const missing = zipSync({ "word/document.xml": document, "word/styles.xml": strToU8(styles) })
    const bytes = polishMarkdownArtifactsInDocx(missing)
    const parts = unzipSync(bytes)
    expect(parts["word/document.xml"]).toEqual(document)
    expect(strFromU8(parts["word/styles.xml"]!)).toBe(styles.replace("</w:styles>", '<w:style w:type="character" w:styleId="Code"><w:name w:val="Code"/></w:style></w:styles>'))
    expect(polishMarkdownArtifactsInDocx(bytes)).toEqual(bytes)
    const existing = zipSync({ "word/document.xml": document, "word/styles.xml": strToU8(styles.replace("</w:styles>", '<w:style w:type="character" w:styleId="Code"><w:basedOn w:val="ExistingBase"/><w:rPr><w:rFonts w:ascii="ExistingFont"/></w:rPr></w:style></w:styles>')) })
    expect(polishMarkdownArtifactsInDocx(existing)).toEqual(existing)
    const alternateQuotes = zipSync({ "word/document.xml": document, "word/styles.xml": strToU8(strFromU8(unzipSync(existing)["word/styles.xml"]!).replace('w:styleId="Code"', "w:styleId='Code'")) })
    expect(polishMarkdownArtifactsInDocx(alternateQuotes)).toEqual(alternateQuotes)
    const ordinary = zipSync({ "word/document.xml": strToU8('<w:document><w:body><w:p><w:r><w:t>ordinary</w:t></w:r></w:p></w:body></w:document>'), "word/styles.xml": strToU8(styles) })
    expect(polishMarkdownArtifactsInDocx(ordinary)).toEqual(ordinary)
  })

  test("converts markdown pipe tables into Word tables", () => {
    const docx = zipSync({
      "word/document.xml": strToU8(
        `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
          paragraph("| Clause | Change made |") +
          paragraph("|---|---|") +
          paragraph("| Section 1 | **Verified** &amp; carried |") +
          `</w:body></w:document>`,
      ),
    })

    const xml = strFromU8(unzipSync(polishMarkdownArtifactsInDocx(docx))["word/document.xml"]!)

    expect(xml).toContain("<w:tbl>")
    expect(xml).toContain("<w:b/>")
    expect(xml).toContain("Clause")
    expect(xml).toContain(">Verified</w:t>")
    expect(xml).toContain("> &amp; carried</w:t>")
    expect(xml).not.toContain("|---|---|")
  })

  test("does not add a synthetic blank paragraph before converted markdown tables", () => {
    const docx = zipSync({
      "word/document.xml": strToU8(
        `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
          paragraph("Heading immediately above") +
          paragraph("| Field | Value |") +
          paragraph("|---|---|") +
          paragraph("| Status | Open |") +
          `</w:body></w:document>`,
      ),
    })

    const xml = strFromU8(unzipSync(polishMarkdownArtifactsInDocx(docx))["word/document.xml"]!)

    expect(xml).not.toContain(`<w:p><w:pPr><w:spacing w:before="0" w:after="180"/></w:pPr></w:p><w:tbl>`)
    expect(xml).toContain(`Heading immediately above</w:t></w:r></w:p><w:tbl>`)
  })

  test("adds direct spacing to heading paragraphs", () => {
    const docx = zipSync({
      "word/document.xml": strToU8(
        `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
          paragraphWithStyle("Heading3", "D. Intentionally retained bracketed factual blanks") +
          paragraph("Following text") +
          `</w:body></w:document>`,
      ),
    })

    const xml = strFromU8(unzipSync(polishMarkdownArtifactsInDocx(docx))["word/document.xml"]!)

    expect(xml).toContain(`<w:pStyle w:val="Heading3" /><w:spacing w:before="220" w:after="120"/>`)
  })

  test("converts standalone markdown dividers into Word paragraph borders", () => {
    const docx = zipSync({
      "word/document.xml": strToU8(
        `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
          paragraph("Before") +
          paragraph("---") +
          paragraph("After") +
          `</w:body></w:document>`,
      ),
    })

    const xml = strFromU8(unzipSync(polishMarkdownArtifactsInDocx(docx))["word/document.xml"]!)

    expect(xml).toContain("<w:pBdr>")
    expect(xml).toContain(`w:color="BFBFBF"`)
    expect(xml).not.toContain("<w:t>---</w:t>")
  })

  test("keeps exact table cell text and inline formatting across structural polish", () => {
    const docx = zipSync({
      "word/document.xml": strToU8(
        `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
          paragraph("| Field | Value |") +
          paragraph("|---|---|") +
          `<w:p><w:r><w:t>| </w:t></w:r>` +
          `<w:r><w:rPr><w:b/></w:rPr><w:t>Status</w:t></w:r>` +
          `<w:r><w:t> | </w:t></w:r>` +
          `<w:r><w:rPr><w:i/></w:rPr><w:t>Open</w:t></w:r>` +
          `<w:r><w:t> and </w:t></w:r>` +
          `<w:r><w:rPr><w:rStyle w:val="Code"/></w:rPr><w:t>reviewed</w:t></w:r>` +
          `<w:r><w:t> |</w:t></w:r></w:p>` +
          paragraph("---") +
          `</w:body></w:document>`,
      ),
    })

    const polished = polishMarkdownArtifactsInDocx(docx)
    const xml = strFromU8(unzipSync(polished)["word/document.xml"]!)

    expect(xml).toContain("<w:tbl>")
    expect(xml).toContain("<w:pBdr>")
    expect(xml).toContain("<w:b/>")
    expect(xml).toContain("<w:i/>")
    expect(xml).toContain('<w:rStyle w:val="Code"/>')
    const cellText = ["Field", "Value", "Status", "Open", " and ", "reviewed"]
    cellText.forEach(
      (text) => expect(xml).toContain(`>${text}</w:t>`),
    )
    const rows = [...xml.matchAll(/<w:tr\b[\s\S]*?<\/w:tr>/g)].map(match => match[0])
    expect(rows).toHaveLength(2)
    rows.forEach(row => expect(row).toContain("<w:cantSplit/>"))
    expect(rows[0]).toContain("<w:tblHeader/>")
    expect(rows[1]).not.toContain("<w:tblHeader/>")
  })

  test("conserves hyperlink and combined-format associations", () => {
    const hyperlink =
      '<w:hyperlink r:id="rId9"><w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>Linked term</w:t></w:r></w:hyperlink>'
    const combined =
      '<w:r><w:rPr><w:b/><w:i/></w:rPr><w:t>Combined term</w:t></w:r>'
    const docx = zipSync({
      "word/document.xml": strToU8(
        `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>` +
          paragraph("| Field | Value |") +
          paragraph("|---|---|") +
          `<w:p><w:r><w:t>| Reference | </w:t></w:r>${hyperlink}<w:r><w:t> and </w:t></w:r>${combined}<w:r><w:t> |</w:t></w:r></w:p>` +
          `</w:body></w:document>`,
      ),
    })

    const polished = polishMarkdownArtifactsInDocx(docx)
    const xml = strFromU8(unzipSync(polished)["word/document.xml"]!)

    expect(xml).toContain(
      '<w:hyperlink r:id="rId9"><w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t xml:space="preserve">Linked term</w:t></w:r></w:hyperlink>',
    )
    expect(xml).toContain(
      '<w:r><w:rPr><w:b/><w:i/></w:rPr><w:t xml:space="preserve">Combined term</w:t></w:r>',
    )
  })

  test("fails closed when a markdown table run has unsupported payload", () => {
    const docx = zipSync({
      "word/document.xml": strToU8(
        `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
          paragraph("| Field | Value |") +
          paragraph("|---|---|") +
          `<w:p><w:r><w:t>| Status | Open</w:t><w:tab/><w:t>item |</w:t></w:r></w:p>` +
          `</w:body></w:document>`,
      ),
    })

    expect(() => polishMarkdownArtifactsInDocx(docx)).toThrow(
      "cannot losslessly conserve an unsupported Word run payload",
    )
  })

  test("conserves visible text represented by numeric XML character references", () => {
    const docx = zipSync({
      "word/document.xml": strToU8(
        `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
          paragraph("| Field | Value |") +
          paragraph("|---|---|") +
          paragraph("| Status | A &#x26; B and C &#38; D |") +
          `</w:body></w:document>`,
      ),
    })

    const polished = polishMarkdownArtifactsInDocx(docx)
    const xml = strFromU8(unzipSync(polished)["word/document.xml"]!)

    expect(xml).toContain(">A &amp; B and C &amp; D</w:t>")
    expect(xml).not.toContain("&amp;#x26;")
    expect(xml).not.toContain("&amp;#38;")
  })

  test("ends a table before a bar-delimited line with paragraph properties, which stays as it was", () => {
    const row = `<w:p><w:pPr><w:bidi/></w:pPr><w:r><w:t>| Status | Open |</w:t></w:r></w:p>`
    const docx = zipSync({
      "word/document.xml": strToU8(
        `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
          paragraph("| Field | Value |") +
          paragraph("|---|---|") +
          row +
          `</w:body></w:document>`,
      ),
    })

    const xml = strFromU8(unzipSync(polishMarkdownArtifactsInDocx(docx))["word/document.xml"]!)

    expect(xml.match(/<w:tr>/g)).toHaveLength(1)
    expect(xml).toContain(`${row}</w:body>`)
  })

  test("leaves bar-delimited list items, headings, quotations and code as text", async () => {
    const { bytes } = await renderMarkdownToDocx([
      "# Notice",
      "> | Quoted | Note |\n> | --- | --- |",
      "- | Listed | Note |\n- | --- | --- |",
      "## | Heading | Note |\n\n| --- | --- |",
      "```\n| Code | Note |\n| --- | --- |\n```",
      "| Lead |\n\n- | Item |",
      "| Field | Value |\n| --- | --- |\n| Status | Open |\n> | Quoted | Row |",
    ].join("\n\n"))
    const xml = strFromU8(unzipSync(bytes)["word/document.xml"]!)
    const tables = xml.match(/<w:tbl>[\s\S]*?<\/w:tbl>/g) ?? []

    expect(tables).toHaveLength(1)
    expect(tables[0]).toContain(">Status</w:t>")
    expect(tables[0]).not.toContain("Quoted")
    for (const line of ["| Quoted | Note |", "| Listed | Note |", "| Heading | Note |", "| Code | Note |", "| Lead |", "| Item |", "| Quoted | Row |"])
      expect(xml).toContain(line)
  })

  test("converts a header-only table and a one-column table, but not a bar line without a separator", () => {
    const docx = zipSync({
      "word/document.xml": strToU8(
        `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
          paragraph("| Signed for the Supplier | Signed for the Buyer |") +
          paragraph("| --- | --- |") +
          paragraph("Between the tables") +
          paragraph("| Important notice |") +
          paragraph("| --- |") +
          paragraph("| This agreement limits liability. |") +
          paragraph("After the notice") +
          paragraph("| Not a table |") +
          paragraph("|") +
          paragraph("|---|") +
          `</w:body></w:document>`,
      ),
    })

    const xml = strFromU8(unzipSync(polishMarkdownArtifactsInDocx(docx))["word/document.xml"]!)
    const tables = xml.match(/<w:tbl>[\s\S]*?<\/w:tbl>/g) ?? []

    expect(tables).toHaveLength(2)
    expect(tables[0]!.match(/<w:tr>/g)).toHaveLength(1)
    expect(tables[0]!.match(/<w:tc>/g)).toHaveLength(2)
    expect(tables[1]!.match(/<w:tr>/g)).toHaveLength(2)
    expect(tables[1]!.match(/<w:tc>/g)).toHaveLength(2)
    expect(tables[1]).toContain(">This agreement limits liability.</w:t>")
    expect(xml).not.toContain("| --- |")
    expect(xml).toContain(">| Not a table |</w:t>")
    expect(xml).toContain("<w:t>|</w:t>")
    expect(xml).toContain("<w:t>|---|</w:t>")
  })
})

function paragraph(text: string) {
  return `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`
}

function paragraphWithStyle(style: string, text: string) {
  return `<w:p><w:pPr><w:pStyle w:val="${style}" /></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`
}
