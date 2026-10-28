import { expect, test } from "bun:test"
import { importedWordChatDifference, reviewedMixedWordChatArtifact, sameWordChatOoxml, verifiedNativeWordChatFormatting, wordChatOoxmlDifference } from "./word-chat-host-ooxml"
import { reviewedWordChatParagraphFormatting } from "../../../dochaus/lib/word-chat-formatting-ooxml"
import { DOMParser } from "@xmldom/xmldom"

const PKG = "http://schemas.microsoft.com/office/2006/xmlPackage"
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const styles = '<w:styles><w:style w:type="paragraph" w:styleId="Normal"><w:pPr><w:spacing w:after="160"/></w:pPr></w:style></w:styles>'
function flat(paragraph = '<w:p><w:r><w:t>Reviewed text.</w:t></w:r></w:p>', dependency = styles, padding = "0") {
  return `<pkg:package xmlns:pkg="${PKG}" xmlns:w="${W}"><pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData><w:document><w:body>${paragraph}</w:body></w:document></pkg:xmlData></pkg:part>${dependency ? `<pkg:part pkg:name="/word/styles.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml" pkg:padding="${padding}"><pkg:xmlData>${dependency}</pkg:xmlData></pkg:part>` : ""}</pkg:package>`
}

test("dependency diagnostics identify package attributes without ignoring their values", () => {
  const before = flat(), changed = flat(undefined, styles, "hidden-package-value")
  expect(sameWordChatOoxml(before, before, before)).toBe(true)
  expect(sameWordChatOoxml(before, changed, before)).toBe(false)
  const reason = wordChatOoxmlDifference(before, changed, before)!
  expect(reason).toContain("/word/styles.xml/part/@padding")
  expect(reason).toContain("attribute changed")
  expect(reason).not.toContain("hidden-package-value")
})

test("dependency diagnostics identify meaningful nested style properties and missing captures", () => {
  const before = flat(), changed = flat(undefined, styles.replace('after="160"', 'after="240"'))
  expect(wordChatOoxmlDifference(before, before, before)).toBeUndefined()
  expect(wordChatOoxmlDifference(before, changed, before)).toContain("/styles[1]/style[1]/pPr[1]/spacing/@after")
  expect(sameWordChatOoxml(before, changed, before)).toBe(false)
  expect(wordChatOoxmlDifference(before, flat(undefined, ""), before)).toBe("/word/styles.xml: dependency missing from current Word capture")
  expect(wordChatOoxmlDifference(flat(undefined, ""), before, before)).toBe("/word/styles.xml: dependency missing from reviewed Word capture")
})

test("main paragraph diagnostics retain the exact existing write-ID guard", () => {
  const before = flat('<w:p w:rsidR="1"><w:r><w:t>Reviewed text.</w:t></w:r></w:p>')
  const observedWrite = before.replace('rsidR="1"', 'rsidR="2"')
  expect(sameWordChatOoxml(before, observedWrite, before)).toBe(true)
  const changed = before.replace('rsidR="1"', 'rsidR="1" w:rsidP="unverified-write-value"')
  expect(sameWordChatOoxml(before, changed, before)).toBe(false)
  const reason = wordChatOoxmlDifference(before, changed, before)!
  expect(reason).toContain("/word/document.xml[1]/p/@rsidP")
  expect(reason).toContain("attribute added")
  expect(reason).not.toContain("unverified-write-value")
})

test("main paragraph text diagnostics omit source text and values", () => {
  const before = flat('<w:p><w:r><w:t>Private original wording.</w:t></w:r></w:p>')
  const changed = flat('<w:p><w:r><w:t>Private changed wording.</w:t></w:r></w:p>')
  expect(sameWordChatOoxml(before, changed, before)).toBe(false)
  const reason = wordChatOoxmlDifference(before, changed, before)!
  expect(reason).toContain("/word/document.xml[1]/p[1]/r[1]/t[1]: text changed")
  expect(reason).not.toContain("Private")
  expect(reason).not.toContain("wording")
})

test("imported view diagnostics identify the failed check without source text", () => {
  const before = flat('<w:p><w:r><w:t>Private original wording.</w:t></w:r></w:p>')
  const wording = flat('<w:p><w:r><w:t>Private altered wording.</w:t></w:r></w:p>')
  const properties = flat('<w:p><w:pPr><w:ind w:left="300"/></w:pPr><w:r><w:t>Private original wording.</w:t></w:r></w:p>')
  expect(importedWordChatDifference(before, before)).toBeUndefined()
  expect(importedWordChatDifference(before, wording)).toBe("current paragraph 1 runs")
  expect(importedWordChatDifference(before, properties)).toBe("current paragraph 1 properties")
})

test("native formatting must retain the prior properties and inline revisions and use the approved properties", () => {
  const before = flat('<w:p><w:pPr><w:spacing w:after="120"/></w:pPr><w:ins w:id="4" w:author="Earlier"><w:r><w:t>Reviewed text.</w:t></w:r></w:ins></w:p>')
  const result = flat('<w:p><w:pPr><w:spacing w:after="240"/><w:pPrChange w:id="5" w:author="doc.haus" w:date="2026-01-01T00:00:00Z"><w:pPr><w:spacing w:after="120"/></w:pPr></w:pPrChange></w:pPr><w:ins w:id="4" w:author="Earlier"><w:r><w:t>Reviewed text.</w:t></w:r></w:ins></w:p>')
  const native = result.replace('w:author="doc.haus"', 'w:author="Word user"')
  const operations = [{ paragraph: 0, spacingAfter: 12 }]
  expect(verifiedNativeWordChatFormatting(before, result, native, operations)).toEqual(["Word user"])
  const renumbered = native.replace('w:id="4"', 'w:id="24"').replace('<w:p>', '<w:p w:rsidP="12345678">')
  expect(verifiedNativeWordChatFormatting(before, result, renumbered, operations)).toEqual(["Word user"])
  for (const corrupted of [native.replace('w:after="120"', 'w:after="140"'), native.replace('w:after="240"', 'w:after="260"'),
    native.replace('w:author="Earlier"', 'w:author="Other"'), native.replace('w:author="Word user"', 'w:author=""'),
    renumbered.replace('w:author="Earlier"', 'w:author="Other"'), renumbered.replace('w:rsidP="12345678"', 'w:rsidP="12345678" w:rsidR="unexpected"')])
    expect(() => verifiedNativeWordChatFormatting(before, result, corrupted, operations)).toThrow()
})

test("native original properties only materialize matching inherited list indentation", () => {
  const style = '<w:styles><w:style w:type="paragraph" w:styleId="ListParagraph"/></w:styles>'
  const numbering = '<w:numbering><w:abstractNum w:abstractNumId="4"><w:lvl w:ilvl="0"><w:pPr><w:ind w:left="360" w:hanging="360"/></w:pPr></w:lvl><w:lvl w:ilvl="1"><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="2"><w:abstractNumId w:val="4"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num></w:numbering>'
  const pack = (paragraph: string, number = numbering, styleXml = style) => flat(paragraph, styleXml).replace('</pkg:package>', `<pkg:part pkg:name="/word/numbering.xml"><pkg:xmlData>${number}</pkg:xmlData></pkg:part></pkg:package>`)
  for (const level of ["0", "1"]) {
    const original = `<w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="${level}"/><w:numId w:val="2"/></w:numPr></w:pPr>`
    const nativeOld = original.replace('</w:pPr>', `<w:ind w:left="${level === "0" ? "360" : "720"}" w:hanging="360"/></w:pPr>`).replace(level === "0" ? '<w:ilvl w:val="0"/>' : 'unchanged', '')
    const body = (properties: string) => `<w:p>${properties}<w:r><w:t>Reviewed text.</w:t></w:r></w:p>`
    const after = (old: string, author: string) => `<w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="${level}"/><w:numId w:val="2"/></w:numPr><w:spacing w:after="240"/><w:pPrChange w:id="5" w:author="${author}" w:date="2026-01-01T00:00:00Z">${old}</w:pPrChange></w:pPr>`
    const before = pack(body(original)), authored = pack(body(after(original, "doc.haus"))), native = pack(body(after(nativeOld, "Word user")))
    const rawParagraph = (xml: string) => new DOMParser().parseFromString(xml, "application/xml").getElementsByTagNameNS(W, "p")[0]!
    expect(() => reviewedWordChatParagraphFormatting([rawParagraph(before)], [rawParagraph(pack(body(after(nativeOld, "doc.haus"))))], [{ paragraph: 0, spacingAfter: 12 }])).toThrow("does not preserve")
    expect(verifiedNativeWordChatFormatting(before, authored, native, [{ paragraph: 0, spacingAfter: 12 }])).toEqual(["Word user"])
    for (const bad of [native.replace('w:hanging="360"', 'w:hanging="400"'), native.replace('w:val="2"', 'w:val="3"'),
      native.replace('w:left="360" w:hanging="360"/></w:pPr></w:lvl>', 'w:left="400" w:hanging="360"/></w:pPr></w:lvl>'),
      native.replace('<w:style w:type="paragraph" w:styleId="ListParagraph"/>', '<w:style w:type="paragraph" w:styleId="ListParagraph"><w:pPr><w:ind w:left="300"/></w:pPr></w:style>')])
      expect(() => verifiedNativeWordChatFormatting(before, authored, bad, [{ paragraph: 0, spacingAfter: 12 }])).toThrow()
    const competing = '<w:styles><w:style w:type="paragraph" w:styleId="ListParagraph"><w:pPr><w:ind w:left="300"/></w:pPr></w:style></w:styles>'
    expect(() => verifiedNativeWordChatFormatting(pack(body(original), numbering, competing), pack(body(after(original, "doc.haus")), numbering, competing), pack(body(after(nativeOld, "Word user")), numbering, competing), [{ paragraph: 0, spacingAfter: 12 }])).toThrow()
    if (level === "1") expect(() => verifiedNativeWordChatFormatting(before, authored, native.replace('<w:ilvl w:val="1"/>', ''), [{ paragraph: 0, spacingAfter: 12 }])).toThrow()
  }
})

test("inserted heading preformatting is proved against a constrained baseline", () => {
  const old = '<w:p><w:r><w:t>Retained.</w:t></w:r></w:p>'
  const insert = '<w:p><w:pPr><w:rPr><w:ins w:id="1" w:author="doc.haus"/></w:rPr></w:pPr><w:ins w:id="2" w:author="doc.haus"><w:r><w:t>Heading</w:t></w:r></w:ins></w:p>'
  const base = flat(insert + old), delivered = flat(insert.replace('<w:pPr>', '<w:pPr><w:pStyle w:val="Heading1"/>') + old)
  const target = { id: "target", text: "Retained.", paragraphs: [{ id: "p", text: "Retained." }], startOffset: 0, endOffset: 9, saved: false, capturedAt: "2026-01-01T00:00:00Z" }
  const revisions = [{ author: "doc.haus", type: "Added" as const, text: "Heading" }, { author: "doc.haus", type: "Added" as const, text: "\r" }]
  const formatting = [{ paragraph: 0, style: "Heading1" as const }]
  const artifact = { proposalId: "proposal", author: "doc.haus" as const, paragraphs: ["Heading", "Retained."], formatting,
    revisions: [...revisions, { author: "doc.haus", type: "Formatted" as const, text: "Heading" }], ooxml: delivered,
    wording: { ooxml: delivered, unformattedOoxml: base, revisions } }
  const verify = (value: typeof artifact) => reviewedMixedWordChatArtifact(value, "proposal", target, artifact.paragraphs, flat(old), formatting)
  const proof = verify(artifact)
  expect(proof.inserted.has(0)).toBe(true)
  expect(proof.nativeFormatCount).toBe(0)
  expect(verifiedNativeWordChatFormatting(delivered, delivered, delivered, formatting, proof.inserted)).toEqual([])
  expect(() => verifiedNativeWordChatFormatting(delivered, delivered, delivered.replace('w:val="Heading1"', 'w:val="Heading2"'), formatting, proof.inserted)).toThrow()
  expect(() => verify({ ...artifact, wording: { ...artifact.wording, unformattedOoxml: base.replace('<w:pPr>', '<w:pPr><w:ind w:left="300"/>') } })).toThrow()
  expect(() => verify({ ...artifact, wording: { ...artifact.wording, ooxml: delivered.replace('w:val="Heading1"', 'w:val="Heading2"') } })).toThrow()
  expect(() => verify({ ...artifact, wording: { ...artifact.wording, ooxml: delivered.replace('<w:p><w:r><w:t>Retained.', '<w:p><w:pPr><w:ind w:left="300"/></w:pPr><w:r><w:t>Retained.') } })).toThrow()
  const bold = (xml: string) => xml.replace('<w:r><w:t>Heading</w:t></w:r>', '<w:r><w:rPr><w:b/></w:rPr><w:t>Heading</w:t></w:r>')
  expect(() => verify({ ...artifact, ooxml: bold(artifact.ooxml), wording: { ...artifact.wording, ooxml: bold(delivered), unformattedOoxml: bold(base) } })).toThrow("Unreviewed insertion run formatting")
})
