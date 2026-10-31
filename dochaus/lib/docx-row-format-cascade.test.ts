import { expect, test } from "bun:test"
import { strToU8, zipSync } from "fflate"
import { docxRowFormatCascade } from "./docx-row-format-cascade"

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
type Parts = { defaults?: string; normal?: string; base?: string; character?: string; table?: string; tableNormal?: string; extra?: string; tblPr?: string; styles?: string | null; rels?: string; document?: string }
// One synthetic package per case: a plain table in a body whose cascade is neutral unless a case adds one source.
function docx(change: Parts = {}) {
  const table = (tblPr: string) => `<w:tbl><w:tblPr>${tblPr}</w:tblPr><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`
  const styles = `<w:styles xmlns:w="${W}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri"/><w14:ligatures w14:val="standard"/>${change.defaults ?? ""}</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160"/></w:pPr></w:pPrDefault></w:docDefaults>` +
    `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:basedOn w:val="Base"/>${change.normal ?? ""}</w:style>` +
    `<w:style w:type="paragraph" w:styleId="Base"><w:name w:val="Base"/><w:rPr><w:sz w:val="22"/>${change.base ?? ""}</w:rPr></w:style>` +
    `<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/>${change.character ?? ""}</w:style>` +
    `<w:style w:type="table" w:styleId="Grid"><w:name w:val="Grid"/><w:basedOn w:val="TableNormal"/><w:tblPr><w:tblBorders/></w:tblPr>${change.table ?? ""}</w:style>` +
    `<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/>${change.tableNormal ?? ""}</w:style>` +
    // Unused definitions never reach the new row.
    `<w:style w:type="paragraph" w:styleId="Heading1"><w:link w:val="Heading1Char"/><w:rPr><w:b/></w:rPr></w:style>${change.extra ?? ""}</w:styles>`
  const parts: Record<string, Uint8Array> = {
    "[Content_Types].xml": strToU8('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>'),
    "word/_rels/document.xml.rels": strToU8(change.rels ?? `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="s" Type="${R}/styles" Target="styles.xml"/></Relationships>`),
    "word/document.xml": strToU8(change.document ?? `<w:document xmlns:w="${W}"><w:body><w:p/>${table("")}${table(change.tblPr ?? '<w:tblStyle w:val="Grid"/>')}<w:sectPr/></w:body></w:document>`),
  }
  if (change.styles !== null) parts["word/styles.xml"] = strToU8(change.styles ?? styles)
  return zipSync(parts)
}

test("a neutral cascade is admitted; one inherited bold, italic or underline source refuses it", () => {
  expect(docxRowFormatCascade(docx(), 1)).toEqual({ neutral: true })
  expect(docxRowFormatCascade(docx(), 0)).toEqual({ neutral: true })
  const cases: [Parts, string][] = [
    [{ defaults: "<w:i/>" }, "doc_defaults:i"],
    [{ defaults: '<w:b w:val="0"/>' }, "doc_defaults:b"],
    [{ defaults: '<w:u w:val="none"/>' }, "doc_defaults:u"],
    [{ base: "<w:b/>" }, "paragraph_style:b"],
    [{ normal: "<w:rPr><w:bCs/></w:rPr>" }, "paragraph_style:bCs"],
    [{ table: '<w:rPr><w:u w:val="single"/></w:rPr>' }, "table_style:u"],
    [{ table: '<w:tblStylePr w:type="firstRow"><w:rPr><w:b/></w:rPr></w:tblStylePr>' }, "table_style:b"],
    [{ character: "<w:rPr><w:iCs/></w:rPr>" }, "character_style:iCs"],
  ]
  for (const [change, reason] of cases) expect(docxRowFormatCascade(docx(change), 1)).toEqual({ neutral: false, reason })
  // The default table style governs a table without tblStyle, and is a base of the named one.
  const inherited = docx({ tableNormal: "<w:rPr><w:i/></w:rPr>" })
  for (const index of [0, 1]) expect(docxRowFormatCascade(inherited, index)).toEqual({ neutral: false, reason: "table_style:i" })
})

test("missing, duplicate, cyclic, referenced or unclassifiable styles are refused", () => {
  const cases: [Parts, number, string][] = [
    [{ tblPr: '<w:tblStyle w:val="Absent"/>' }, 1, "missing_style:table_style"],
    [{ extra: '<w:style w:type="paragraph" w:default="1" w:styleId="Other"/>' }, 1, "duplicate_default_style:paragraph"],
    [{ extra: '<w:style w:type="paragraph" w:styleId="Base"/>' }, 1, "duplicate_style:paragraph_style"],
    [{ tblPr: '<w:tblStyle w:val="Normal"/>' }, 1, "style_type_mismatch:table_style"],
    [{ extra: '<w:style w:type="table" w:styleId="Loop"><w:basedOn w:val="Loop"/></w:style>', tblPr: '<w:tblStyle w:val="Loop"/>' }, 1, "cyclic_style:table_style"],
    [{ normal: '<w:link w:val="NormalChar"/>' }, 1, "style_reference:paragraph_style:link"],
    [{ defaults: '<w:rStyle w:val="Strong"/>' }, 1, "style_reference:doc_defaults:rStyle"],
    [{ normal: '<w:basedOn w:val="Other"/>' }, 1, "ambiguous_based_on:paragraph_style"],
    [{ tblPr: '<w:tblStyle w:val="Grid"/><w:tblStyle w:val="TableNormal"/>' }, 1, "ambiguous_table_style"],
    [{ extra: '<w:style w:type="paragraph" w:default="maybe" w:styleId="Odd"/>' }, 1, "unclassified_default_style"],
    [{ defaults: '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"/>' }, 1, "unclassified_markup:doc_defaults"],
    [{ styles: `<w:styles xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main"/>` }, 1, "unclassified_styles"],
    [{ styles: `<!DOCTYPE w:styles><w:styles xmlns:w="${W}"/>` }, 1, "invalid_xml:styles"],
    [{ styles: null }, 1, "styles_part_missing"],
    [{ rels: `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>` }, 1, "unrelated_styles_part"],
    [{ rels: `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="s" Type="${R}/styles" Target="other.xml"/></Relationships>` }, 1, "unclassified_styles_part"],
    // A strict-schema styles relationship must not read as "no styles".
    [{ rels: `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="s" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/styles" Target="other.xml"/></Relationships>`, styles: null }, 1, "unclassified_styles_part"],
    [{}, 2, "table_missing"],
  ]
  for (const [change, index, reason] of cases) expect(docxRowFormatCascade(docx(change), index)).toEqual({ neutral: false, reason })
  expect(docxRowFormatCascade(new Uint8Array([1, 2, 3]), 0)).toEqual({ neutral: false, reason: "unreadable_package" })
})

test("a package without styles applies nothing unless the table names a style", () => {
  const rels = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`
  expect(docxRowFormatCascade(docx({ styles: null, rels, tblPr: "" }), 1)).toEqual({ neutral: true })
  expect(docxRowFormatCascade(docx({ styles: null, rels }), 1)).toEqual({ neutral: false, reason: "missing_style:table_style" })
})
