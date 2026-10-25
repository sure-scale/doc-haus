import { describe, expect, test } from "bun:test"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { safeStyleValuesInXml, withSafeStyleValues } from "../src/components/docx-style-values"

// The rendered effect is covered in Chromium by docx-render-css.browser.ts; these
// cover the parts of the package the viewers' copy changes, and those it keeps.
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const A = "http://schemas.openxmlformats.org/drawingml/2006/main"
const V = "urn:schemas-microsoft-com:vml"

describe("style values in the copy a viewer renders", () => {
  test("a file whose font names and colours are all ordinary renders from its own bytes", () => {
    const names = ["Times New Roman", "MS 明朝", "ＭＳ Ｐゴシック", "맑은 고딕", "Wingdings 2", "Arial Unicode MS", "Helvetica-Bold", "@MS Mincho", "Arial (W1)", "Arial,Bold"]
    const bytes = zipSync({
      "word/document.xml": strToU8(`<w:document xmlns:w="${W}" xmlns:v="${V}"><w:background w:color="FFFFFF"/><w:body>${names.map(name => `<w:p><w:pPr><w:pBdr><w:top w:val="single" w:sz="4" w:space="1" w:color="auto"/></w:pBdr><w:shd w:val="clear" w:color="auto" w:fill="E2EFDA"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="${name}" w:eastAsia="${name}" w:hint="eastAsia"/><w:color w:val="1F3864" w:themeColor="accent1" w:themeShade="BF"/><w:highlight w:val="darkYellow"/><w:u w:val="single" w:color="FF0000"/></w:rPr><w:sym w:font="${name}" w:char="F0B7"/></w:r></w:p>`).join("")}<w:p><w:r><w:pict><v:shape><v:fill color="#4f81bd [3204]"/></v:shape></w:pict></w:r></w:p></w:body></w:document>`),
      "word/fontTable.xml": strToU8(`<w:fonts xmlns:w="${W}">${names.map(name => `<w:font w:name="${name}"><w:altName w:val="${name}"/><w:panose1 w:val="02020603050405020304"/></w:font>`).join("")}</w:fonts>`),
      "word/theme/theme1.xml": strToU8(`<a:theme xmlns:a="${A}"><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:accent1><a:srgbClr val="4472C4"/></a:accent1></a:clrScheme><a:majorFont><a:latin typeface="Calibri Light" panose="020F0302020204030204"/><a:ea typeface=""/><a:font script="Jpan" typeface="游ゴシック Light"/></a:majorFont><a:minorFont><a:latin typeface="+mn-lt"/></a:minorFont></a:theme>`),
    })
    expect(withSafeStyleValues(bytes)).toBe(bytes)
  })

  test("a name that could end its CSS string is cleaned however the XML spells it, and only font names change", () => {
    const cases = [
      `<w:rFonts w:ascii="A&apos;; } B"/>`,
      `<w:rFonts w:ascii="A&#39;B&#x5C;C&#10;D"/>`,
      `<q:rFonts q:hAnsi="A'B"/>`,
      `<w:rFonts w:cs='A"B'/>`,
      `<w:rFonts w:ascii="A>'; } B"/>`,
      `<w:font w:name="A'B">`,
      `<w:altName w:val="A'B"/>`,
      `<w:sym w:font="A'B" w:char="F0B7"/>`,
      `<a:latin typeface="A'B"/>`,
      `<a:font script="Jpan" typeface="A'B"/>`,
    ]
    for (const tag of cases) {
      const cleaned = safeStyleValuesInXml(tag)
      const value = cleaned?.match(/:?(?:ascii|hAnsi|cs|name|val|font|typeface)="([^"]*)"/)?.[1]
      expect({ tag, value }).toEqual({ tag, value: expect.stringMatching(/^[\p{L}\p{M}\p{N} \-_.,+&@()]+$/u) })
    }
    expect(safeStyleValuesInXml(`<w:rFonts w:ascii="A'; } B" w:hint="eastAsia"/>`)).toBe(`<w:rFonts w:ascii="A  B" w:hint="eastAsia"/>`)
    // A colour is written bare, so a semicolon or brace is enough to leave it.
    expect(safeStyleValuesInXml(`<w:color w:val="000000; } p { color: #000000"/>`)).toBe(`<w:color w:val="000000pcolor000000"/>`)
    for (const tag of [`<w:shd w:val="clear" w:fill="FFF;}"/>`, `<w:top w:val="single" w:color="0F0;}"/>`, `<w:highlight w:val="yellow;}"/>`, `<x:background x:color="FFF;}"/>`, `<a:srgbClr val="FFF;}"/>`])
      expect({ tag, cleaned: safeStyleValuesInXml(tag) }).toEqual({ tag, cleaned: tag.replace(";}", "") })
    // Text and other attributes keep their characters.
    expect(safeStyleValuesInXml(`<w:p><w:r><w:t>It's {fine}</w:t></w:r><w:pStyle w:val="A'B"/></w:p>`)).toBeUndefined()
  })

  test("only the parts that change are rewritten into the rendered copy", () => {
    const document = `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>It's agreed.</w:t></w:r></w:p></w:body></w:document>`
    const bytes = zipSync({ "word/document.xml": strToU8(document), "word/styles.xml": strToU8(`<w:styles xmlns:w="${W}"><w:rFonts w:ascii="A'; } B"/></w:styles>`) })
    const parts = unzipSync(withSafeStyleValues(bytes))
    expect(strFromU8(parts["word/document.xml"]!)).toBe(document)
    expect(strFromU8(parts["word/styles.xml"]!)).toBe(`<w:styles xmlns:w="${W}"><w:rFonts w:ascii="A  B"/></w:styles>`)
  })
})
