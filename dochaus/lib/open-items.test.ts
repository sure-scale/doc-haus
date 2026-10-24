// Run from dochaus/: bun test lib/open-items.test.ts
import { describe, expect, test } from "bun:test"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { FIELD_HIGHLIGHT_FILL, OPEN_ITEM_HIGHLIGHT_FILL, highlightOpenItemsInDocx } from "./open-items"

describe("highlightOpenItemsInDocx", () => {
  test("highlights bracketed and parenthetical open item markers", () => {
    const docx = zipSync({
      "word/document.xml": strToU8(
        `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>` +
          `<w:r><w:t xml:space="preserve">Before [open item: client address] after.</w:t></w:r>` +
          `<w:r><w:t xml:space="preserve">Need (open item: Maryland notice wording, per R02.) now.</w:t></w:r>` +
          `</w:p></w:body></w:document>`,
      ),
    })

    const xml = strFromU8(unzipSync(highlightOpenItemsInDocx(docx))["word/document.xml"]!)

    expect(xml).toContain(`w:fill="${OPEN_ITEM_HIGHLIGHT_FILL}"`)
    expect(xml.match(new RegExp(`w:fill="${OPEN_ITEM_HIGHLIGHT_FILL}"`, "g"))).toHaveLength(2)
    expect(xml).not.toContain(`w:fill="${FIELD_HIGHLIGHT_FILL}"`)
    expect(xml).toContain("[open item: client address]")
    expect(xml).toContain("(open item: Maryland notice wording, per R02.)")
  })

  test("does not highlight ordinary placeholders", () => {
    const docx = zipSync({
      "word/document.xml": strToU8(
        `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>[insert execution date]</w:t></w:r></w:p></w:body></w:document>`,
      ),
    })

    expect(strFromU8(unzipSync(highlightOpenItemsInDocx(docx))["word/document.xml"]!)).not.toContain(
      `w:fill="${OPEN_ITEM_HIGHLIGHT_FILL}"`,
    )
  })

  test("highlights fillable underscore fields", () => {
    const docx = zipSync({
      "word/document.xml": strToU8(
        `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>` +
          `<w:r><w:t>Signature: _______________________________</w:t></w:r>` +
          `<w:r><w:t>Initial here: ___</w:t></w:r>` +
          `<w:r><w:t>Not a field: __</w:t></w:r>` +
          `</w:p></w:body></w:document>`,
      ),
    })

    const xml = strFromU8(unzipSync(highlightOpenItemsInDocx(docx))["word/document.xml"]!)

    expect(xml).not.toContain(`w:fill="${OPEN_ITEM_HIGHLIGHT_FILL}"`)
    expect(xml.match(new RegExp(`w:fill="${FIELD_HIGHLIGHT_FILL}"`, "g"))).toHaveLength(2)
    expect(xml).toContain("_______________________________")
    expect(xml).toContain("___")
  })

  test("splits a Word run of several texts only at its markers, keeping each text once and its spaces", () => {
    // Word writes a signature cell's lines as one run: two texts with a line break between.
    const run = `<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">By: ________</w:t><w:br/><w:t>Name: ________ (open item: signatory)</w:t></w:r>`
    const docx = zipSync({
      "word/document.xml": strToU8(
        `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>${run}</w:p></w:body></w:document>`,
      ),
    })

    const xml = strFromU8(unzipSync(highlightOpenItemsInDocx(docx))["word/document.xml"]!)
    const pieces = [...xml.matchAll(/<w:r>([\s\S]*?)<\/w:r>/g)].map((match) => match[1]!)

    // The text reads as before, each piece carrying the run's own properties.
    expect(pieces.map((piece) => piece.replace(/<w:br\/>/g, "\n").replace(/<[^>]+>/g, "")).join("")).toBe("By: ________\nName: ________ (open item: signatory)")
    expect(pieces.every((piece) => piece.startsWith("<w:rPr><w:b/>"))).toBe(true)
    expect(xml.match(new RegExp(`w:fill="${FIELD_HIGHLIGHT_FILL}"`, "g"))).toHaveLength(2)
    expect(xml.match(new RegExp(`w:fill="${OPEN_ITEM_HIGHLIGHT_FILL}"`, "g"))).toHaveLength(1)
    // A piece whose text starts or ends with a space keeps it.
    expect(xml).toContain(`<w:t xml:space="preserve">By: </w:t>`)
    expect(xml).toContain(`<w:t xml:space="preserve">Name: </w:t>`)
    expect(xml).toContain(`<w:t xml:space="preserve"> </w:t>`)
  })

  test("recolors existing highlighted underscore fields to field blue", () => {
    const docx = zipSync({
      "word/document.xml": strToU8(
        `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>` +
          `<w:r><w:rPr><w:shd w:val="clear" w:color="auto" w:fill="${OPEN_ITEM_HIGHLIGHT_FILL}"/></w:rPr><w:t>__________</w:t></w:r>` +
          `</w:p></w:body></w:document>`,
      ),
    })

    const xml = strFromU8(unzipSync(highlightOpenItemsInDocx(docx))["word/document.xml"]!)

    expect(xml).not.toContain(`w:fill="${OPEN_ITEM_HIGHLIGHT_FILL}"`)
    expect(xml).toContain(`w:fill="${FIELD_HIGHLIGHT_FILL}"`)
  })
})
