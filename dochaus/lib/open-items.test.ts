// Run from dochaus/: bun test lib/open-items.test.ts
import { describe, expect, test } from "bun:test"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { DOMParser } from "@xmldom/xmldom"
import { FIELD_HIGHLIGHT_FILL, OPEN_ITEM_HIGHLIGHT_FILL, highlightOpenItemsInDocx, shadeOpenItemsInDocx } from "./open-items"

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

  test("highlights an ordinary placeholder whole as an open item, a bracketed blank as a field, and nothing beside them", () => {
    const docx = zipSync({
      "word/document.xml": strToU8(
        `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>` +
          `<w:r><w:t xml:space="preserve">Dated [insert execution date] by the parties; accounts [TBC – outstanding]; signed [____].</w:t></w:r>` +
          `</w:p></w:body></w:document>`,
      ),
    })

    expect(shadedRuns(highlightOpenItemsInDocx(docx))).toEqual([
      { text: "Dated ", fill: undefined },
      { text: "[insert execution date]", fill: OPEN_ITEM_HIGHLIGHT_FILL },
      { text: " by the parties; accounts ", fill: undefined },
      { text: "[TBC – outstanding]", fill: OPEN_ITEM_HIGHLIGHT_FILL },
      { text: "; signed ", fill: undefined },
      { text: "[____]", fill: FIELD_HIGHLIGHT_FILL },
      { text: ".", fill: undefined },
    ])
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

describe("shadeOpenItemsInDocx", () => {
  test("shades a marker whole across runs to its own closing bracket and, in a paragraph it is asked to clear, nothing else as open", () => {
    const shade = (fill: string) => `<w:shd w:val="clear" w:color="auto" w:fill="${fill}"/>`
    const docx = zipSync({
      "word/document.xml": strToU8(
        `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
          // Words written into a marker's run, the marker ending in the next run.
          `<w:p><w:r><w:rPr><w:b/>${shade(OPEN_ITEM_HIGHLIGHT_FILL)}</w:rPr><w:t xml:space="preserve">Words written [open item: see [Disputes</w:t></w:r>` +
          `<w:r><w:t xml:space="preserve">]] and more </w:t></w:r><w:r><w:rPr>${shade("D9D9D9")}</w:rPr><w:t>grey</w:t></w:r></w:p>` +
          // A paragraph not cleared keeps its own shading.
          `<w:p><w:r><w:rPr>${shade(OPEN_ITEM_HIGHLIGHT_FILL)}</w:rPr><w:t>kept as it is</w:t></w:r></w:p>` +
          `</w:body></w:document>`,
      ),
    })

    const written = (paragraph: Element) => (paragraph.textContent ?? "").startsWith("Words")
    const shaded = shadeOpenItemsInDocx(docx, written)
    const doc = new DOMParser().parseFromString(strFromU8(unzipSync(shaded)["word/document.xml"]!), "application/xml")
    const runs = Array.from(doc.getElementsByTagName("w:r")).map(run => ({
      text: run.textContent, bold: run.getElementsByTagName("w:b").length > 0, fill: run.getElementsByTagName("w:shd")[0]?.getAttribute("w:fill"),
    }))

    expect(runs.filter(run => run.fill === OPEN_ITEM_HIGHLIGHT_FILL).map(run => run.text)).toEqual(["[open item: see [Disputes", "]]", "kept as it is"])
    expect(runs.find(run => run.text === "Words written ")).toEqual({ text: "Words written ", bold: true, fill: undefined })
    expect(runs.find(run => run.text === " and more ")?.fill).toBeUndefined()
    expect(runs.find(run => run.text === "grey")?.fill).toBe("D9D9D9")
    // Shaded again, the body is the same bytes.
    expect(shadeOpenItemsInDocx(shaded, written)).toBe(shaded)
  })

  test("shades a drafting placeholder whole across runs, a bracketed blank as a field, an open item naming a placeholder as one open item, and nothing beside them", () => {
    const docx = zipSync({
      "word/document.xml": strToU8(
        `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
          // A placeholder split by a bold word, as drafting renders "[insert the **agreed** period]".
          `<w:p><w:r><w:t xml:space="preserve">Payable within [insert the </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>agreed</w:t></w:r>` +
          `<w:r><w:t xml:space="preserve"> period] of invoice; cap [£ (cap figure to be confirmed)].</w:t></w:r></w:p>` +
          `<w:p><w:r><w:t xml:space="preserve">Sign here: [____] and initial ___.</w:t></w:r></w:p>` +
          `<w:p><w:r><w:t xml:space="preserve">Fee: [open item: confirm whether the fee [TBC] is to be confirmed] now.</w:t></w:r></w:p>` +
          `</w:body></w:document>`,
      ),
    })

    const shaded = shadeOpenItemsInDocx(docx)
    expect(shadedRuns(shaded)).toEqual([
      { text: "Payable within ", fill: undefined },
      { text: "[insert the ", fill: OPEN_ITEM_HIGHLIGHT_FILL },
      { text: "agreed", fill: OPEN_ITEM_HIGHLIGHT_FILL },
      { text: " period]", fill: OPEN_ITEM_HIGHLIGHT_FILL },
      { text: " of invoice; cap ", fill: undefined },
      { text: "[£ (cap figure to be confirmed)]", fill: OPEN_ITEM_HIGHLIGHT_FILL },
      { text: ".", fill: undefined },
      { text: "Sign here: ", fill: undefined },
      { text: "[____]", fill: FIELD_HIGHLIGHT_FILL },
      { text: " and initial ", fill: undefined },
      { text: "___", fill: FIELD_HIGHLIGHT_FILL },
      { text: ".", fill: undefined },
      { text: "Fee: ", fill: undefined },
      { text: "[open item: confirm whether the fee [TBC] is to be confirmed]", fill: OPEN_ITEM_HIGHLIGHT_FILL },
      { text: " now.", fill: undefined },
    ])
    // The bold word inside the placeholder stays bold.
    expect(strFromU8(unzipSync(shaded)["word/document.xml"]!)).toMatch(/<w:rPr><w:b\/><w:shd [^>]*\/><\/w:rPr><w:t[^>]*>agreed<\/w:t>/)
    expect(shadeOpenItemsInDocx(shaded)).toBe(shaded)
  })

  test("leaves a run a lawyer shaded in a fill of their own as it is, its placeholder included", () => {
    const own = `<w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/>`
    const docx = zipSync({
      "word/document.xml": strToU8(
        `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>` +
          `<w:r><w:rPr>${own}</w:rPr><w:t>[insert date]</w:t></w:r><w:r><w:t xml:space="preserve"> and </w:t></w:r>` +
          `<w:r><w:rPr>${own}</w:rPr><w:t xml:space="preserve">due on [TBC] each year</w:t></w:r>` +
          `</w:p></w:body></w:document>`,
      ),
    })

    // Drafting's highlight and the end pass, the paragraph cleared or not, return the body itself.
    for (const shaded of [highlightOpenItemsInDocx(docx), shadeOpenItemsInDocx(docx), shadeOpenItemsInDocx(docx, () => true)]) expect(shaded).toBe(docx)
  })

  test("reads a tab or break in a tracked deletion as no text: it is not shaded and does not move the markers", () => {
    const docx = zipSync({
      "word/document.xml": strToU8(
        `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
          `<w:p><w:r><w:t>[open item: a</w:t></w:r><w:del w:id="1"><w:r><w:tab/></w:r></w:del><w:ins w:id="2"><w:r><w:t>b]</w:t></w:r></w:ins></w:p>` +
          `</w:body></w:document>`,
      ),
    })

    const doc = new DOMParser().parseFromString(strFromU8(unzipSync(shadeOpenItemsInDocx(docx))["word/document.xml"]!), "application/xml")
    const deleted = doc.getElementsByTagName("w:del")[0]!
    expect(deleted.getElementsByTagName("w:shd").length).toBe(0)
    expect(Array.from(doc.getElementsByTagName("w:r")).filter(run => run.getElementsByTagName("w:shd").length).map(run => run.textContent)).toEqual(["[open item: a", "b]"])
  })
})

// Each run of a body with text, in document order, with its shading fill.
function shadedRuns(bytes: Uint8Array) {
  const doc = new DOMParser().parseFromString(strFromU8(unzipSync(bytes)["word/document.xml"]!), "application/xml")
  return Array.from(doc.getElementsByTagName("w:r")).filter(run => run.textContent)
    .map(run => ({ text: run.textContent, fill: run.getElementsByTagName("w:shd")[0]?.getAttribute("w:fill") ?? undefined }))
}
