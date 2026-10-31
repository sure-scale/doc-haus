import { expect, test } from "bun:test"
import { reviewedWordChatInlineControl, importedWordChatDifference, reviewedMixedWordChatArtifact, sameWordChatOoxml, sameNativeWordChatParagraphContent, sameWordChatParagraphContent, sameWordChatParagraphProperties, sameWordChatReplacementFormatting, sameWordChatTargetSnapshot, sameNativeWordChatBodySnapshot, uniformWordChatFormatting, verifiedNativeWordChatFormatting, wordChatOoxmlDifference } from "./word-chat-host-ooxml"
import { reviewedWordChatParagraphFormatting } from "../../../dochaus/lib/word-chat-formatting-ooxml"
import { DOMParser } from "@xmldom/xmldom"

const PKG = "http://schemas.microsoft.com/office/2006/xmlPackage"
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const styles = '<w:styles><w:style w:type="paragraph" w:styleId="Normal"><w:pPr><w:spacing w:after="160"/></w:pPr></w:style></w:styles>'
function flat(paragraph = '<w:p><w:r><w:t>Reviewed text.</w:t></w:r></w:p>', dependency = styles, padding = "0") {
  return `<pkg:package xmlns:pkg="${PKG}" xmlns:w="${W}"><pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData><w:document><w:body>${paragraph}</w:body></w:document></pkg:xmlData></pkg:part>${dependency ? `<pkg:part pkg:name="/word/styles.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml" pkg:padding="${padding}"><pkg:xmlData>${dependency}</pkg:xmlData></pkg:part>` : ""}</pkg:package>`
}

test("native whole-body snapshot compares mixed structure and protected content", () => {
  const body = '<w:p><w:pPr><w:spacing w:after="120"/></w:pPr><w:r><w:t>Unrelated</w:t></w:r><w:ins w:id="3" w:author="Prior"><w:r><w:t>Old revision</w:t></w:r></w:ins></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr><w:pgSz w:w="12000"/></w:sectPr>'
  const source = flat(body)
  expect(() => sameWordChatTargetSnapshot(source, source)).toThrow()
  expect(sameNativeWordChatBodySnapshot(source, source)).toBe(true)
  for (const changed of [body.replace('Unrelated', 'Changed'), body.replace('w:after="120"', 'w:after="240"'), body.replace('</w:tr>', '<w:tc><w:p/></w:tc></w:tr>'), body.replace('w:author="Prior"', 'w:author="Other"'), body.replace('w:w="12000"', 'w:w="13000"')])
    expect(sameNativeWordChatBodySnapshot(source, flat(changed))).toBe(false)
  expect(sameNativeWordChatBodySnapshot(source, '<invalid')).toBe(false)
})

test("native whole-body snapshot ignores only terminal section export rsid", () => {
  const body = '<w:p w:rsidP="11111111"><w:r><w:t>Text</w:t></w:r></w:p><w:sectPr w:rsidSect="11111111"><w:pgSz w:w="12000"/></w:sectPr>'
  const source = flat(body)
  expect(sameNativeWordChatBodySnapshot(source, flat(body.replace('w:rsidSect="11111111"', 'w:rsidSect="22222222"')))).toBe(true)
  expect(sameNativeWordChatBodySnapshot(source, flat(body.replace('w:rsidP="11111111"', 'w:rsidP="22222222"')))).toBe(false)
  expect(sameNativeWordChatBodySnapshot(source, flat(body.replace('w:w="12000"', 'w:w="13000"')))).toBe(false)
  const nonterminal = flat('<w:sectPr w:rsidSect="11111111"/><w:p/>')
  expect(sameNativeWordChatBodySnapshot(nonterminal, nonterminal.replace('11111111', '22222222'))).toBe(false)
})

test("native body projection binds table order and retains real empty paragraph and properties", async () => {
  const { projectNativeWordChatBody } = await import("./word-chat-host-ooxml")
  const body = '<w:p><w:r><w:t>Same</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:pPr><w:spacing w:after="360"/></w:pPr><w:ins w:id="3" w:author="Prior"><w:r><w:rPr><w:b/></w:rPr><w:t>Same</w:t></w:r></w:ins></w:p></w:tc></w:tr></w:tbl><w:p/><w:p/><w:sectPr/>'
  const projection = projectNativeWordChatBody(flat(body), [{id:"a",text:"Same",tableNestingLevel:0},{id:"b",text:"Same",tableNestingLevel:1},{id:"c",text:"",tableNestingLevel:0}])
  expect(projection.paragraphs.map(p => p.id)).toEqual(["a", "b", "c"])
  expect(projection.paragraphs[1]!.path).toContain("tbl")
  expect(projection.scope([1])).toContain('w:after="360"')
  expect(projection.scope([1])).toContain("<w:b/>")
  expect(projection.scope([1])).toContain('w:author="Prior"')
  expect(projection.scope([1])).toContain("/word/styles.xml")
  expect(projection.scope([2])).toContain("<w:p/>")
  const changed = projectNativeWordChatBody(flat(body.replace('<w:tbl>', '<w:tbl><w:tblPr><w:tblW w:w="2000"/></w:tblPr>')), [{id:"a",text:"Same",tableNestingLevel:0},{id:"b",text:"Same",tableNestingLevel:1},{id:"c",text:"",tableNestingLevel:0}])
  expect(changed.geometry).not.toBe(projection.geometry)
  const layout = projectNativeWordChatBody(flat(body.replace('<w:sectPr/>', '<w:sectPr><w:pgSz w:w="12000"/></w:sectPr>')), [{id:"a",text:"Same",tableNestingLevel:0},{id:"b",text:"Same",tableNestingLevel:1},{id:"c",text:"",tableNestingLevel:0}])
  expect(layout.geometry).not.toBe(projection.geometry)
})

test("native body geometry ignores Word's export renumbering of tracked row IDs, not the row revision", async () => {
  const { projectNativeWordChatBody } = await import("./word-chat-host-ooxml")
  const inv = [{id:"a",text:"Same"},{id:"b",text:"Cell",tableNestingLevel:1}]
  const body = (id: string, mark = "ins", author = "Prior") => `<w:p><w:r><w:t>Same</w:t></w:r></w:p><w:tbl><w:tr><w:trPr><w:${mark} w:id="${id}" w:author="${author}" w:date="2026-01-01T00:00:00Z"/></w:trPr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/>`
  const geometry = (xml: string) => projectNativeWordChatBody(flat(xml), inv).geometry
  // A new revision earlier in the story shifts every later export ID.
  expect(geometry(body("11"))).toBe(geometry(body("10")))
  expect(geometry(body("10", "del"))).not.toBe(geometry(body("10")))
  expect(geometry(body("10", "ins", "Other"))).not.toBe(geometry(body("10")))
})

test("native body projection refuses unbound or substantive terminal carriers", async () => {
  const { projectNativeWordChatBody } = await import("./word-chat-host-ooxml")
  const inv = [{id:"a",text:"One"}], p = '<w:p><w:r><w:t>One</w:t></w:r></w:p>'
  expect(() => projectNativeWordChatBody(flat(p + '<w:p><w:bookmarkStart w:id="1" w:name="x"/></w:p><w:sectPr/>'), inv)).toThrow()
  expect(() => projectNativeWordChatBody(flat(p + '<w:p><w:pPr/></w:p><w:sectPr/>'), inv)).toThrow()
  expect(() => projectNativeWordChatBody(flat(p + '<w:p><!-- marker --></w:p><w:sectPr/>'), inv)).toThrow()
  expect(() => projectNativeWordChatBody(flat(p + '<w:p/><w:p/><w:sectPr/>'), inv)).toThrow()
  expect(() => projectNativeWordChatBody(flat(p + '<w:p/>'), inv)).toThrow()
  expect(() => projectNativeWordChatBody(flat(p), [{id:"a",text:"Other"}])).toThrow()
  expect(() => projectNativeWordChatBody(flat(p), [{id:"a",text:"One",tableNestingLevel:1}])).toThrow()
  expect(() => projectNativeWordChatBody(flat(p+p), [{id:"a",text:"One"},{id:"a",text:"One"}])).toThrow()
  expect(() => projectNativeWordChatBody(flat('<w:sdt><w:sdtContent>'+p+'</w:sdtContent></w:sdt>'), inv)).toThrow()
})

test("native wording checks paragraph properties inside the selected cell wrapper", () => {
  const cell = (spacing: number) => flat(`<w:tbl><w:tr><w:tc><w:p><w:pPr><w:spacing w:after="${spacing}"/></w:pPr><w:r><w:t>Terms.</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`)
  expect(sameWordChatParagraphProperties(cell(120), cell(120))).toBe(true)
  expect(sameWordChatParagraphProperties(cell(120), cell(240))).toBe(false)
})

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
  const churn = renumbered.replace('<w:p ', '<w:p w:rsidR="00ABCDEF" w:rsidRDefault="00123456" ')
  expect(verifiedNativeWordChatFormatting(before, result, churn, operations)).toEqual(["Word user"])
  for (const corrupted of [native.replace('w:after="120"', 'w:after="140"'), native.replace('w:after="240"', 'w:after="260"'),
    native.replace('w:author="Earlier"', 'w:author="Other"'), native.replace('w:author="Word user"', 'w:author=""'),
    native.replace('Reviewed text.', 'Altered text.'), renumbered.replace('w:author="Earlier"', 'w:author="Other"'), renumbered.replace('w:rsidP="12345678"', 'w:rsidP="12345678" w:rsidR="unexpected"')])
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


test("exact inline control verifier binds the shared authored artifact and rejects tampering", async () => {
  const authorModule: string = "../../../dochaus/lib/docx-flat-opc"
  const { authorWordChatTarget } = await import(authorModule)
  const rel = "http://schemas.openxmlformats.org/package/2006/relationships", office = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  const control = '<w:sdt><w:sdtPr><w:id w:val="41"/><w:tag w:val="colour"/></w:sdtPr><w:sdtContent><w:r><w:rPr><w:b/></w:rPr><w:t>red</w:t></w:r></w:sdtContent></w:sdt>'
  const source = flat(`<w:p><w:r><w:t xml:space="preserve">A </w:t></w:r>${control}<w:r><w:t xml:space="preserve"> clause</w:t></w:r></w:p>`, "")
    .replace('</pkg:package>', `<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData><Relationships xmlns="${rel}"><Relationship Id="r1" Type="${office}/officeDocument" Target="word/document.xml"/></Relationships></pkg:xmlData></pkg:part></pkg:package>`)
  const target = { id:"target", paragraphs:[{id:"paragraph",text:"A red clause"}],text:"red",startOffset:2,endOffset:5,saved:false,capturedAt:"2026-01-01T00:00:00Z" }
  const authored = { ...await authorWordChatTarget(source, target, "blue"), proposalId:"proposal" }
  const verify = (xml = authored.ooxml, input = source) => reviewedWordChatInlineControl({ ...authored, ooxml:xml }, "proposal", target, "blue", input)
  expect(verify()).toEqual({paragraphId:"paragraph",controlId:"41",startOffset:2,endOffset:5,before:"red",after:"blue"})
  for (const bad of [authored.ooxml.replace('w:val="41"', 'w:val="42"'), authored.ooxml.replace('w:val="colour"', 'w:val="other"'),
    authored.ooxml.replace('<w:b/>', '<w:i/>'), authored.ooxml.replace('w:author="doc.haus"', 'w:author="Other"'),
    authored.ooxml.replace('red</w:delText>', 'tan</w:delText>'), authored.ooxml.replace('blue</w:t>', 'gold</w:t>'),
    authored.ooxml.replace(' clause</w:t>', ' wrong</w:t>'),
    authored.ooxml.replace('w:author="doc.haus"', 'w:author="doc.haus" w:custom="bad"'),
    authored.ooxml.replace('</w:sdt>', '</w:sdt><w:sdt><w:sdtPr/><w:sdtContent/></w:sdt>')]) {
    expect(bad).not.toBe(authored.ooxml)
    expect(() => verify(bad)).toThrow()
  }
  const deleted = authored.ooxml.match(/<w:del\b[^>]*>[\s\S]*?<\/w:del>/)?.[0]
  const inserted = authored.ooxml.match(/<w:ins\b[^>]*>[\s\S]*?<\/w:ins>/)?.[0]
  expect(deleted && inserted).toBeTruthy()
  const swapped = authored.ooxml.replace(deleted! + inserted!, inserted! + deleted!)
  expect(swapped).not.toBe(authored.ooxml)
  expect(() => verify(swapped)).toThrow()
  const deleteId = deleted!.match(/w:id="(\d+)"/)![1]!, insertId = inserted!.match(/w:id="(\d+)"/)![1]!
  const duplicateId = authored.ooxml.replace(inserted!, inserted!.replace(`w:id="${insertId}"`, `w:id="${deleteId}"`))
  expect(duplicateId).not.toBe(authored.ooxml)
  expect(() => verify(duplicateId)).toThrow()
  const extra = authored.ooxml.replace(inserted!, inserted! + inserted!)
  expect(extra).not.toBe(authored.ooxml)
  expect(() => verify(extra)).toThrow()
  expect(() => verify(authored.ooxml, source.replace('</w:sdtPr>', '<w:lock w:val="sdtContentLocked"/></w:sdtPr>'))).toThrow()
  expect(() => reviewedWordChatInlineControl(authored, "different", target, "blue", source)).toThrow()
  expect(() => reviewedWordChatInlineControl(authored, "proposal", { ...target, startOffset: 1 }, "blue", source)).toThrow()
  const table = source.replace('<w:body><w:p>', '<w:body><w:tbl><w:tr><w:tc><w:p>').replace('</w:p></w:body>', '</w:p></w:tc></w:tr></w:tbl></w:body>')
  const tableArtifact = { ...await authorWordChatTarget(table, target, "blue"), proposalId: "proposal" }
  expect(reviewedWordChatInlineControl(tableArtifact, "proposal", target, "blue", table).controlId).toBe("41")
  const wrapped = table.replace('<w:tbl><w:tr>', '<w:tbl><w:sdt><w:sdtPr><w:lock w:val="sdtLocked"/></w:sdtPr><w:sdtContent><w:tr>')
    .replace('</w:tr></w:tbl>', '</w:tr></w:sdtContent></w:sdt></w:tbl>')
  expect(() => reviewedWordChatInlineControl(tableArtifact, "proposal", target, "blue", wrapped)).toThrow()
})


test("native wording requires uniform formatting and preserves the replaced span's formatting", () => {
  const plain = '<w:r><w:t xml:space="preserve">Charge £</w:t></w:r>'
  const bold = (text: string) => `<w:r><w:rPr><w:b/></w:rPr><w:t>${text}</w:t></w:r>`
  const italic = '<w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve"> per month</w:t></w:r>'
  const before = flat(`<w:p>${plain}${bold("2,500")}${italic}</w:p>`)
  const after = flat(`<w:p>${plain}${bold("2,750")}${italic}</w:p>`)
  const wronglyStyled = flat(`<w:p>${plain}<w:r><w:rPr><w:i/></w:rPr><w:t>2,750</w:t></w:r>${italic}</w:p>`)
  expect(uniformWordChatFormatting(before, 8, 13)).toBe(true)
  expect(uniformWordChatFormatting(before, 0, 13)).toBe(false)
  expect(sameWordChatReplacementFormatting(before, after, 8, 13, "2,750")).toBe(true)
  expect(sameWordChatReplacementFormatting(before, wronglyStyled, 8, 13, "2,750")).toBe(false)
  expect(sameWordChatReplacementFormatting(before, flat(`<w:p>${plain}${bold("2,500")}${italic}</w:p>`), 8, 13, "2,750")).toBe(false)
})

test("authored granular hunks preserve styled boundary whitespace and reject unsafe groups", async () => {
  const { authoredWordChatWordingHunks, sameWordChatHunkFormatting } = await import("./word-chat-host-ooxml")
  const b = (value: string) => `<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">${value}</w:t></w:r>`
  const i = (value: string) => `<w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve">${value}</w:t></w:r>`
  const plain = (value: string) => `<w:r><w:t xml:space="preserve">${value}</w:t></w:r>`
  const del = (value: string, fmt: string) => `<w:del w:id="1" w:author="doc.haus"><w:r><w:rPr>${fmt}</w:rPr><w:delText>${value}</w:delText></w:r></w:del>`
  const ins = (value: string, fmt: string) => `<w:ins w:id="2" w:author="doc.haus"><w:r><w:rPr>${fmt}</w:rPr><w:t>${value}</w:t></w:r></w:ins>`
  const source = flat(`<w:p>${b("Cell bold ")}${plain("Cell focus")}${i(" italic tail")}</w:p>`)
  const authored = flat(`<w:p>${del("Cell bold", "<w:b/>")}${ins("Cell important", "<w:b/>")}${b(" ")}${plain("Cell focus")}${i(" ")}${del("italic tail", "<w:i/>")}${ins("italic ending", "<w:i/>")}</w:p>`)
  const before = "Cell bold Cell focus italic tail", after = "Cell important Cell focus italic ending"
  const hunks = authoredWordChatWordingHunks(authored, before, after, 0, before.length)
  expect(hunks).toEqual([{ start: 0, end: 9, before: "Cell bold", after: "Cell important" }, { start: 21, end: 32, before: "italic tail", after: "italic ending" }])
  expect(sameWordChatHunkFormatting(source, authored, hunks)).toBe(true)
  expect(sameWordChatHunkFormatting(source, authored.replace(i(" "), plain(" ")), hunks)).toBe(false)
  expect(() => authoredWordChatWordingHunks(authored, before, after, 1, before.length)).toThrow()
  expect(() => authoredWordChatWordingHunks(authored.replace(del("Cell bold", "<w:b/>"), ""), before, after, 0, before.length)).toThrow()
})

test("one-sided authored hunks are scoped and preserve source formatting", async () => {
  const { authoredWordChatWordingHunks, sameWordChatHunkFormatting, unambiguousWordChatInsertionFormatting } = await import("./word-chat-host-ooxml")
  const run = (v: string) => `<w:r><w:t xml:space="preserve">${v}</w:t></w:r>`
  const add = (v: string) => `<w:ins w:id="1" w:author="doc.haus">${run(v)}</w:ins>`
  const del = (v: string) => `<w:del w:id="2" w:author="doc.haus"><w:r><w:delText>${v}</w:delText></w:r></w:del>`
  const before = flat(`<w:p>${run("Alpha beta")}</w:p>`)
  const added = flat(`<w:p>${run("Alpha ")}${add("new ")}${run("beta")}</w:p>`)
  const addedHunks = authoredWordChatWordingHunks(added, "Alpha beta", "Alpha new beta", 0, 10)
  expect(addedHunks).toEqual([{ start: 6, end: 6, before: "", after: "new " }])
  expect(sameWordChatHunkFormatting(before, added, addedHunks)).toBe(true)
  expect(() => authoredWordChatWordingHunks(added, "Alpha beta", "Alpha new beta", 0, 5)).toThrow()
  const deleted = flat(`<w:p>${run("Alpha ")}${del("beta")}</w:p>`)
  const deletedHunks = authoredWordChatWordingHunks(deleted, "Alpha beta", "Alpha ", 0, 10)
  expect(deletedHunks).toEqual([{ start: 6, end: 10, before: "beta", after: "" }])
  expect(sameWordChatHunkFormatting(before, deleted, deletedHunks)).toBe(true)
  const boundary = flat(`<w:p>${run("Alpha ")}<w:r><w:rPr><w:b/></w:rPr><w:t>beta</w:t></w:r></w:p>`)
  expect(unambiguousWordChatInsertionFormatting(boundary, 6)).toBe(false)
  expect(unambiguousWordChatInsertionFormatting(boundary, 0)).toBe(true)
  expect(uniformWordChatFormatting(boundary, 6, 6)).toBe(false)
  expect(sameWordChatHunkFormatting(boundary, added, addedHunks)).toBe(false)
  const emoji = flat(`<w:p>${run("A😀B")}</w:p>`)
  expect(unambiguousWordChatInsertionFormatting(emoji, 2)).toBe(false)
  expect(() => authoredWordChatWordingHunks(added, "Alpha beta", "Alpha new beta", 0, 11)).toThrow()
})

test("real author output supports a pure insertion and deletion", async () => {
  const { authorWordChatTarget } = await import(new URL("../../../dochaus/lib/docx-flat-opc.ts", import.meta.url).href) as { authorWordChatTarget(xml: string, target: import("../../../dochaus/lib/word-chat-amendment-contract").WordChatTarget, replacement: string): Promise<Omit<import("../../../dochaus/lib/word-chat-amendment-contract").WordChatAuthoredAmendment, "proposalId">> }
  const { authoredWordChatWordingHunks, sameWordChatHunkFormatting } = await import("./word-chat-host-ooxml")
  const rel = "http://schemas.openxmlformats.org/package/2006/relationships", office = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  const source = flat('<w:p><w:r><w:t xml:space="preserve">Alpha beta</w:t></w:r></w:p>', "")
    .replace('</pkg:package>', `<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData><Relationships xmlns="${rel}"><Relationship Id="r1" Type="${office}/officeDocument" Target="word/document.xml"/></Relationships></pkg:xmlData></pkg:part></pkg:package>`)
  const target = { id: "target", paragraphs: [{ id: "p", text: "Alpha beta" }], text: "beta", startOffset: 6, endOffset: 10, saved: false, capturedAt: "" }
  for (const replacement of ["new beta", ""]) {
    const artifact = await authorWordChatTarget(source, target, replacement)
    const hunks = authoredWordChatWordingHunks(artifact.ooxml, "Alpha beta", `Alpha ${replacement}`, 6, 10)
    expect(hunks.some(hunk => !hunk.before || !hunk.after)).toBe(true)
    expect(sameWordChatHunkFormatting(source, artifact.ooxml, hunks)).toBe(true)
  }
})

test("real author preserves original formatting for a mixed-run deletion", async () => {
  const { authorWordChatTarget } = await import(new URL("../../../dochaus/lib/docx-flat-opc.ts", import.meta.url).href) as { authorWordChatTarget(xml: string, target: import("../../../dochaus/lib/word-chat-amendment-contract").WordChatTarget, replacement: string): Promise<Omit<import("../../../dochaus/lib/word-chat-amendment-contract").WordChatAuthoredAmendment, "proposalId">> }
  const { authoredWordChatWordingHunks, sameWordChatHunkFormatting } = await import("./word-chat-host-ooxml")
  const rel = "http://schemas.openxmlformats.org/package/2006/relationships", office = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  const source = flat('<w:p><w:r><w:t xml:space="preserve">Keep </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r><w:r><w:t xml:space="preserve"> plain </w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>italic</w:t></w:r><w:r><w:t xml:space="preserve"> end</w:t></w:r></w:p>', "")
    .replace('</pkg:package>', `<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData><Relationships xmlns="${rel}"><Relationship Id="r1" Type="${office}/officeDocument" Target="word/document.xml"/></Relationships></pkg:xmlData></pkg:part></pkg:package>`)
  const original = "Keep bold plain italic end", selected = "bold plain italic"
  const target = { id: "target", paragraphs: [{ id: "p", text: original }], text: selected, startOffset: 5, endOffset: 22, saved: false, capturedAt: "" }
  const artifact = await authorWordChatTarget(source, target, "")
  const hunks = authoredWordChatWordingHunks(artifact.ooxml, original, "Keep  end", 5, 22)
  expect(hunks).toEqual([{ start: 5, end: 22, before: selected, after: "" }])
  expect(sameWordChatHunkFormatting(source, artifact.ooxml, hunks)).toBe(true)
  const tampered = artifact.ooxml.replace("<w:b/>", "<w:i/>")
  expect(tampered).not.toBe(artifact.ooxml)
  expect(sameWordChatHunkFormatting(source, tampered, hunks)).toBe(false)
})


test("post-format neighbour comparison accepts only validated inline revision ID reassignment", () => {
  const paragraph = '<w:p><w:pPr><w:spacing w:after="120"/></w:pPr><w:del w:id="3" w:author="Reviewer" w:date="2026-01-01T00:00:00Z"><w:r><w:rPr><w:b/></w:rPr><w:delText>old</w:delText></w:r></w:del><w:ins w:id="4" w:author="Reviewer" w:date="2026-01-01T00:00:00Z"><w:r><w:t>new</w:t></w:r></w:ins><w:bookmarkStart w:id="9" w:name="mark"/><w:bookmarkEnd w:id="9"/></w:p>'
  const before = flat(paragraph), after = flat(paragraph.replace('w:id="3"', 'w:id="13"').replace('w:id="4"', 'w:id="14"'))
  expect(sameWordChatParagraphContent(before, after)).toBe(false)
  expect(sameNativeWordChatParagraphContent(before, after)).toBe(true)
  for (const drift of [after.replace('>new<', '>wrong<'), after.replace('w:author="Reviewer"', 'w:author="Other"'), after.replace('2026-01-01', '2026-01-02'), after.replace('<w:b/>', '<w:i/>'), after.replace('w:id="14"', 'w:id="13"'), after.replace('w:id="14"', 'w:id="bad"'), after.replace('w:id="9" w:name="mark"', 'w:id="10" w:name="mark"'), after.replace('w:after="120"', 'w:after="240"')])
    expect(sameNativeWordChatParagraphContent(before, drift)).toBe(false)
  const alternatePrefix = after.replace('xmlns:w="' + W + '"', 'xmlns:w="' + W + '" xmlns:ns0="' + W + '"').replace('w:id="13"', 'ns0:id="13"').replace('w:id="14"', 'ns0:id="14"')
  expect(sameNativeWordChatParagraphContent(before, alternatePrefix)).toBe(true)
  expect(sameNativeWordChatParagraphContent(before, alternatePrefix.replace('w:author="Reviewer"', 'w:author="Other"'))).toBe(false)
  const duplicate = flat(paragraph.replace('w:id="4"', 'w:id="3"'))
  expect(sameNativeWordChatParagraphContent(duplicate, duplicate)).toBe(false)
})

test("native row remainder checks bound geometry and all-depth prior revisions", async () => {
  const { verifyNativeWordChatRowRemainder: verify } = await import("./word-chat-host-ooxml")
  const cell = (text: string) => `<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`
  const old = `<w:p><w:pPr><w:pPrChange w:id="1" w:author="Earlier" w:date="2026-01-01"><w:pPr/></w:pPrChange></w:pPr><w:ins w:id="2" w:author="Earlier" w:date="2026-01-01"><w:r><w:rPr><w:b/><w:rPrChange w:id="3" w:author="Earlier" w:date="2026-01-01"><w:rPr/></w:rPrChange></w:rPr><w:t>Prior</w:t></w:r></w:ins></w:p>`
  const row = `<w:tr>${cell("First")}${cell("Second")}</w:tr>`, inserted = `<w:tr>${cell("New")}${cell("Value")}</w:tr>`
  const body = `${old}<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>${row}</w:tbl><w:p><w:r><w:t>Tail</w:t></w:r></w:p>`
  const after = body.replace(row, row+inserted).replace('w:id="1"','w:id="11"').replace('w:id="2"','w:id="12"').replace('w:id="3"','w:id="13"')
  const beforeInv = [{id:"p",text:"Prior",tableNestingLevel:0},{id:"a",text:"First",tableNestingLevel:1},{id:"b",text:"Second",tableNestingLevel:1},{id:"tail",text:"Tail",tableNestingLevel:0}]
  const afterInv = [...beforeInv.slice(0,3),{id:"new1",text:"New",tableNestingLevel:1},{id:"new2",text:"Value",tableNestingLevel:1},beforeInv[3]!]
  const binding = {tableIndex:0,rows:[[{id:"a",text:"First"},{id:"b",text:"Second"}]]}
  const check = (xml:string, inv=afterInv) => verify(flat(body),flat(xml),beforeInv,inv,binding,["New","Value"])
  expect(check(after).carrierMap.map(c => [c.before,c.after])).toEqual([["1","11"],["2","12"],["3","13"]])
  for (const drift of [after.replace('>Prior<','>Wrong<'),after.replace('w:author="Earlier"','w:author="Other"'),after.replace('2026-01-01','2026-01-02'),after.replace('<w:b/>','<w:i/>'),after.replace('w:id="13"','w:id="12"'),after.replace('w:w="2400"','w:w="2300"'),after.replace(inserted,inserted+inserted),after.replace('>New<','>Other<')]) expect(() => check(drift)).toThrow()
  expect(() => check(after,[afterInv[0]!,afterInv[2]!,afterInv[1]!,...afterInv.slice(3)])).toThrow()
})


test("reviewed row oracle verifies the authored package and rejects tampering", async () => {
  const {zipSync,strToU8,strFromU8,unzipSync} = await import("fflate")
  const rowWriterPath = new URL("../../../dochaus/lib/docx-table-row-insertion.ts", import.meta.url).href
  const {docxTableRowInsertionTargets,authorDocxTableRowInsertion} = await import(rowWriterPath)
  const {verifyReviewedWordChatRowOracle: verify} = await import("./word-chat-host-ooxml")
const ROW_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const xml = (text: string) => text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
const p = (text: string) => `<w:p><w:r><w:t>${xml(text)}</w:t></w:r></w:p>`
const cell = (content: string, props = "") => `<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/>${props}</w:tcPr>${content}</w:tc>`
const row = (content: string, props = "") => `<w:tr><w:trPr>${props}</w:trPr>${content}</w:tr>`
function fixture(kind = "simple") {
  const stamp = 'w:author="Existing Editor" w:date="2026-01-01T00:00:00Z"'
  const retained = `<w:p><w:bookmarkStart w:id="40" w:name="_dh_source"/><w:commentRangeStart w:id="2"/><w:del w:id="100" ${stamp}><w:r><w:delText>Five</w:delText></w:r></w:del><w:ins w:id="101" ${stamp}><w:r><w:t>Ten</w:t></w:r></w:ins><w:commentRangeEnd w:id="2"/><w:r><w:commentReference w:id="2"/></w:r><w:bookmarkEnd w:id="40"/></w:p>`
  const format = `<w:p><w:r><w:rPr><w:b/><w:rPrChange w:id="102" ${stamp}><w:rPr/></w:rPrChange></w:rPr><w:t>Prior bold</w:t></w:r><w:hyperlink r:id="link"><w:r><w:t>link</w:t></w:r></w:hyperlink><w:fldSimple w:instr=" DATE "><w:r><w:t>Fixed cache</w:t></w:r></w:fldSimple></w:p>`
  const rows = kind === "merged" ? row(cell(p("Merged"), '<w:gridSpan w:val="2"/>'))
    : row(cell(p("Reference")) + cell(retained), kind === "pending" ? `<w:ins w:id="103" ${stamp}/>` : "") + row(cell(format) + cell(p("Right")))
  const table = `<w:tbl><w:tblPr>${kind === "styled" ? '<w:tblStyle w:val="Conditional"/>' : ""}</w:tblPr><w:tblGrid>${'<w:gridCol w:w="2400"/>'.repeat(2)}</w:tblGrid>${rows}</w:tbl>`
  const body = kind === "nested" ? `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4800"/></w:tblGrid>${row(cell(p("Outer") + table))}</w:tbl>` : table
  return zipSync({
    "[Content_Types].xml": strToU8('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/><Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>'),
    "_rels/.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="main" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/_rels/document.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="${R}/styles" Target="styles.xml"/><Relationship Id="settings" Type="${R}/settings" Target="settings.xml"/><Relationship Id="comments" Type="${R}/comments" Target="comments.xml"/><Relationship Id="link" Type="${R}/hyperlink" Target="https://example.com" TargetMode="External"/></Relationships>`),
    "word/styles.xml": strToU8(`<w:styles xmlns:w="${ROW_W}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>${kind === "conditional" ? '<w:style w:type="table" w:styleId="Conditional"><w:tblStylePr w:type="firstRow"><w:rPr><w:b/></w:rPr></w:tblStylePr></w:style>' : ""}</w:styles>`),
    "word/settings.xml": strToU8(`<w:settings xmlns:w="${ROW_W}"/>`),
    "word/comments.xml": strToU8(`<w:comments xmlns:w="${ROW_W}"><w:comment w:id="2" w:author="Commenter"><w:p><w:r><w:t>Keep original comment.</w:t></w:r></w:p></w:comment></w:comments>`),
    "word/document.xml": strToU8(`<w:document xmlns:w="${ROW_W}" xmlns:r="${R}"><w:body>${p("Before")}${body}${p("After")}<w:sectPr/></w:body></w:document>`),
  }, { mtime: new Date(1980, 0, 1) })
}

  const source=fixture(), packet=await docxTableRowInsertionTargets(source), target=packet.targets[1]!
  const table={tableIndex:0,rows:[[{id:"a",text:"Reference"},{id:"b",text:"Ten"}],[{id:"c",text:"Prior boldlinkFixed cache"},{id:"d",text:"Right"}]]}
  const proposal={kind:"table_row_insert" as const,id:"p",documentId:"d",explanation:"Add",openItems:[],createdAt:"now",position:"end" as const,table,cells:["New","Value"]}
  const writerRequest={source_sha256:packet.source_sha256,target_handle:target.target_handle,position:"after" as const,cells:proposal.cells,author:"doc.haus"}
  const authored=await authorDocxTableRowInsertion(source,writerRequest)
  const encode=(bytes:Uint8Array)=>btoa(Array.from(bytes,x=>String.fromCharCode(x)).join(""))
  const oracle={kind:"table_row_insert" as const,proposalId:"p",sourceSha256:packet.source_sha256,table,cells:proposal.cells,author:"doc.haus" as const,writerRequest,authoredBase64:encode(authored.bytes),revisions:authored.revisions}
  expect((await verify(source,oracle,proposal)).cells).toEqual(proposal.cells)
  await expect(verify(source,{...oracle,proposalId:"other"},proposal)).rejects.toThrow()
  await expect(verify(source,{...oracle,sourceSha256:"0".repeat(64)},proposal)).rejects.toThrow()
  await expect(verify(source,{...oracle,writerRequest:{...writerRequest,position:"before" as "after"}},proposal)).rejects.toThrow()
  await expect(verify(source,oracle,{...proposal,cells:["Wrong","Value"]})).rejects.toThrow()
  for (const name of ["word/document.xml","word/styles.xml","word/settings.xml"]) {
    const parts=unzipSync(authored.bytes), value=strFromU8(parts[name]!)
    parts[name]=strToU8(name==="word/document.xml"?value.replace("Reference","Changed"):value.replace(name==="word/styles.xml"?"Normal":"trackRevisions",name==="word/styles.xml"?"Changed":"differentSetting"))
    await expect(verify(source,{...oracle,authoredBase64:encode(zipSync(parts))},proposal)).rejects.toThrow()
  }
})

test("reviewed row oracle binds formatted spans and refuses a table whose styles already format text", async () => {
  const {zipSync,strToU8,strFromU8,unzipSync} = await import("fflate")
  const {docxTableRowInsertionTargets,authorDocxTableRowInsertion} = await import(new URL("../../../dochaus/lib/docx-table-row-insertion.ts", import.meta.url).href)
  const {verifyReviewedWordChatRowOracle: verify, WordChatRowFormatRefusal} = await import("./word-chat-host-ooxml")
  const RW = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", RR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  const cell = (text: string) => `<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`
  const styles = (defaults = "") => strToU8(`<w:styles xmlns:w="${RW}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri"/>${defaults}</w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`)
  const source = zipSync({
    "[Content_Types].xml": strToU8('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'),
    "_rels/.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="main" Type="${RR}/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/_rels/document.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="${RR}/styles" Target="styles.xml"/><Relationship Id="settings" Type="${RR}/settings" Target="settings.xml"/></Relationships>`),
    "word/styles.xml": styles(), "word/settings.xml": strToU8(`<w:settings xmlns:w="${RW}"/>`),
    "word/document.xml": strToU8(`<w:document xmlns:w="${RW}"><w:body><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid><w:tr>${cell("Term")}${cell("Meaning")}</w:tr></w:tbl><w:p/><w:sectPr/></w:body></w:document>`),
  }, { mtime: new Date(1980, 0, 1) })
  const span = (text: string, flags = "") => ({ text, bold: flags.includes("b"), italic: flags.includes("i"), underline: flags.includes("u") })
  const runs = [[span("Bold", "b"), span(" then plain")], [span("a"), span("tail", "iu")]], cells = ["Bold then plain", "atail"]
  const table = {tableIndex:0,rows:[[{id:"a",text:"Term"},{id:"b",text:"Meaning"}]]}
  const proposal = {kind:"table_row_insert" as const,id:"p",documentId:"d",explanation:"Add",openItems:[],createdAt:"now",position:"end" as const,table,cells,runs}
  const packet = await docxTableRowInsertionTargets(source)
  const writerRequest = {source_sha256:packet.source_sha256,target_handle:packet.targets[0]!.target_handle,position:"after" as const,cells,runs,author:"doc.haus"}
  const authored = await authorDocxTableRowInsertion(source, writerRequest)
  const encode = (bytes: Uint8Array) => btoa(Array.from(bytes, x => String.fromCharCode(x)).join(""))
  const oracle = {kind:"table_row_insert" as const,proposalId:"p",sourceSha256:packet.source_sha256,table,cells,runs,author:"doc.haus" as const,writerRequest,authoredBase64:encode(authored.bytes),revisions:authored.revisions}
  expect((await verify(source, oracle, proposal)).cells).toEqual(cells)
  const tampered = (change: (xml: string) => string) => {
    const parts = unzipSync(authored.bytes), xml = strFromU8(parts["word/document.xml"]!), next = change(xml)
    expect(next).not.toBe(xml)
    parts["word/document.xml"] = strToU8(next)
    return {...oracle, authoredBase64:encode(zipSync(parts))}
  }
  for (const change of [(x: string) => x.replace(">Bold<", ">Bol<").replace("> then plain<", ">d then plain<"), (x: string) => x.replace(/<w:b ?\/>/, "<w:i/>"),
    (x: string) => x.replace(/<w:b ?\/>/, "<w:b/><w:bCs/>"), (x: string) => x.replace(/<w:rPr ?\/><\/w:rPrChange>/, "<w:rPr><w:b/></w:rPr></w:rPrChange>"),
    (x: string) => x.replace(/<w:rPrChange [^>]*><w:rPr ?\/><\/w:rPrChange>/, "")])
    await expect(verify(source, tampered(change), proposal)).rejects.toThrow("Reviewed row oracle mismatch")
  // Spans must equal the proposal's and travel unchanged through the writer request.
  await expect(verify(source, {...oracle, runs:[runs[0]!, [span("atail")]]}, proposal)).rejects.toThrow("Reviewed row oracle mismatch")
  await expect(verify(source, {...oracle, writerRequest:{...writerRequest, runs:undefined}}, proposal)).rejects.toThrow("Reviewed row oracle mismatch")
  const {runs: _runs, ...plain} = proposal
  await expect(verify(source, oracle, plain)).rejects.toThrow("Reviewed row oracle mismatch")
  // Same oracle, one inherited italic in the source styles: the host refuses on its own.
  const restyle = async (defaults: string) => {
    const s = unzipSync(source), a = unzipSync(authored.bytes)
    s["word/styles.xml"] = a["word/styles.xml"] = styles(defaults)
    const bytes = zipSync(s), digest = Array.from(new Uint8Array(await crypto.subtle.digest("SHA-256", bytes)), x => x.toString(16).padStart(2, "0")).join("")
    return [bytes, {...oracle, sourceSha256:digest, writerRequest:{...writerRequest, source_sha256:digest}, authoredBase64:encode(zipSync(a))}] as const
  }
  const [neutral, neutralOracle] = await restyle('<w:sz w:val="22"/>')
  expect((await verify(neutral, neutralOracle, proposal)).cells).toEqual(cells)
  const [italic, italicOracle] = await restyle("<w:i/>")
  const refused = await verify(italic, italicOracle, proposal).catch((error: unknown) => error)
  expect(refused).toBeInstanceOf(WordChatRowFormatRefusal)
  expect((refused as Error).message).toBe("This table's style already applies bold, italic or underline, so a mixed-format row can't be added safely. Your document is unchanged.")
})

test("native inline verification preserves a marker within the reviewed span and rejects anchor drift", async () => {
  const { assertNativeWordChatInlineTarget, verifiedNativeWordChatInlineFormatting } = await import("./word-chat-host-ooxml")
  const marker = '<w:bookmarkStart w:id="4" w:name="place"/>'
  const before = flat(`<w:p><w:r><w:t>ab</w:t></w:r>${marker}<w:r><w:t>cd</w:t></w:r></w:p>`)
  const run = (text: string, id: number, author: string) => `<w:r><w:rPr><w:b/><w:rPrChange w:id="${id}" w:author="${author}" w:date="2026-01-01T00:00:00Z"><w:rPr/></w:rPrChange></w:rPr><w:t>${text}</w:t></w:r>`
  const after = (author: string, anchor = marker) => flat(`<w:p>${run("ab", 8, author)}${anchor}${run("cd", 9, author)}</w:p>`)
  const target = {id:"target", text:"abcd", paragraphs:[{id:"p",text:"abcd"}], startOffset:0,endOffset:4,capturedAt:"now",saved:false}
  assertNativeWordChatInlineTarget(before,target,{bold:true})
  expect(verifiedNativeWordChatInlineFormatting(before,after("Native user"),target,{bold:true},{bold:true})).toEqual(["Native user"])
  expect(() => verifiedNativeWordChatInlineFormatting(before,after("Native user", ""),target,{bold:true},{bold:true})).toThrow()
  expect(() => verifiedNativeWordChatInlineFormatting(before,after("Native user", marker.replace('w:id="4"','w:id="5"')),target,{bold:true},{bold:true})).toThrow()
})

test("native inline bridge admits evidenced session metadata and matched prior revision renumbering only", async () => {
  const { verifiedNativeWordChatInlineFormatting } = await import("./word-chat-host-ooxml")
  const prior = (id:string, author="Prior", date="2026-01-01T00:00:00Z") => `<w:ins w:id="${id}" w:author="${author}" w:date="${date}"><w:r><w:t>Old</w:t></w:r></w:ins>`
  const mark = '<w:rPrChange w:id="2" w:author="Native" w:date="2026-01-01T00:00:00Z"><w:rPr/></w:rPrChange>'
  const para = (id:string, rsid:string, runAttr:string, props:string, old=prior(id), num="5", paraId="AABBCCDD") => `<w:p xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" w14:paraId="${paraId}" w:rsidR="${rsid}" w:rsidRDefault="${rsid}"><w:pPr><w:numPr><w:numId w:val="${num}"/></w:numPr></w:pPr>${old}<w:r ${runAttr}><w:rPr>${props}</w:rPr><w:t>Target</w:t></w:r></w:p>`
  const before = flat(para("2","00112233","","")), after = flat(para("3","AABBCCDD",'w:rsidRPr="01020304"',`<w:b/>${mark}`))
  const target = {id:"target",text:"Target",paragraphs:[{id:"p",text:"OldTarget"}],startOffset:3,endOffset:9,capturedAt:"now",saved:false}
  const verify = (value:string) => verifiedNativeWordChatInlineFormatting(before,value,target,{bold:true},{bold:true})
  expect(verify(after)).toEqual(["Native"])
  expect(()=>verify(after.replace('w:rsidRPr="01020304"','w:rsidRPr="bad"'))).toThrow()
  expect(()=>verify(after.replace('w14:paraId="AABBCCDD"','w14:paraId="11223344"'))).toThrow()
  expect(()=>verify(after.replace('w:numId w:val="5"','w:numId w:val="6"'))).toThrow()
  expect(()=>verify(after.replace('w:author="Prior"','w:author="Other"'))).toThrow()
  expect(()=>verify(after.replace('w:date="2026-01-01T00:00:00Z"','w:date="2026-01-02T00:00:00Z"'))).toThrow()
  expect(()=>verify(after.replace('<w:b/>','<w:b/><w:i/>'))).toThrow()
  expect(()=>verify(after.replace('w:ins w:id="3"','w:ins w:id="2"'))).toThrow()
})

test("inline formatting preserves exact formatting definition parts", async () => {
  const { assertNativeWordChatInlineDependencies } = await import("./word-chat-host-ooxml")
  const part = (name:string, body:string, type="application/xml") => `<pkg:part pkg:name="${name}" pkg:contentType="${type}"><pkg:xmlData>${body}</pkg:xmlData></pkg:part>`
  const stylesPart = part("/word/styles.xml",'<w:styles><w:style w:styleId="Normal"/></w:styles>')
  const numbering = part("/word/numbering.xml",'<w:numbering><w:num w:numId="5"/></w:numbering>')
  const theme = part("/word/theme/theme1.xml",'<theme xmlns="urn:example"><color>blue</color></theme>')
  const before = flat(undefined,"", "0").replace('</pkg:package>',`${stylesPart}${numbering}${theme}</pkg:package>`)
  expect(()=>assertNativeWordChatInlineDependencies(before,before)).not.toThrow()
  expect(()=>assertNativeWordChatInlineDependencies(before,before.replace('w:styleId="Normal"','w:styleId="Other"'))).toThrow()
  expect(()=>assertNativeWordChatInlineDependencies(before,before.replace('w:numId="5"','w:numId="6"'))).toThrow()
  expect(()=>assertNativeWordChatInlineDependencies(before,before.replace(numbering,""))).toThrow()
  expect(()=>assertNativeWordChatInlineDependencies(before,before.replace('application/xml','text/xml'))).toThrow()
  expect(()=>assertNativeWordChatInlineDependencies(before,before.replace('>blue<','>red<'))).toThrow()
})

test("inline unchanged paragraph matches prior formatting revisions only by exact payload and position", async () => {
  const {sameNativeInlineWordChatParagraph: same} = await import("./word-chat-host-ooxml")
  const change = (id:string, author="Prior", date="2026-01-01T00:00:00Z", old="") => `<w:rPrChange w:id="${id}" w:author="${author}" w:date="${date}"><w:rPr>${old}</w:rPr></w:rPrChange>`
  const para = (id:string, author="Prior", date="2026-01-01T00:00:00Z", old="") => `<w:p xmlns:w="${W}"><w:r><w:rPr><w:b/>${change(id,author,date,old)}</w:rPr><w:t>Prior</w:t></w:r><w:r><w:t>Tail</w:t></w:r></w:p>`
  const before=para("8"), after=para("9")
  expect(same(before,after)).toBe(true)
  expect(same(before,para("9","Other"))).toBe(false)
  expect(same(before,para("9","Prior","2026-01-02T00:00:00Z"))).toBe(false)
  expect(same(before,para("9","Prior","2026-01-01T00:00:00Z","<w:i/>"))).toBe(false)
  expect(same(before,after.replace('<w:t>Prior</w:t>','<w:t>Changed</w:t>'))).toBe(false)
  expect(same(before,after.replace('<w:r><w:t>Tail</w:t></w:r>','<w:ins w:id="9" w:author="Prior"><w:r><w:t>Tail</w:t></w:r></w:ins>'))).toBe(false)
})

test("inline target retains prior formatting outside selection through ID reassignment", async () => {
  const {verifiedNativeWordChatInlineFormatting: verify} = await import("./word-chat-host-ooxml")
  const prior = (id:string) => `<w:rPrChange w:id="${id}" w:author="Prior" w:date="2026-01-01T00:00:00Z"><w:rPr/></w:rPrChange>`
  const fresh = '<w:rPrChange w:id="8" w:author="Native" w:date="2026-01-01T00:01:00Z"><w:rPr/></w:rPrChange>'
  const para = (id:string, newProps="") => `<w:p><w:r><w:rPr><w:b/>${prior(id)}</w:rPr><w:t>Old</w:t></w:r><w:r><w:rPr>${newProps}</w:rPr><w:t>Target</w:t></w:r></w:p>`
  const before=flat(para("8")), after=flat(para("9",`<w:i/>${fresh}`)), target={id:"t",text:"Target",paragraphs:[{id:"p",text:"OldTarget"}],startOffset:3,endOffset:9,capturedAt:"now",saved:false}
  expect(verify(before,after,target,{italic:true},{italic:true})).toEqual(["Native"])
  expect(()=>verify(before,after.replace('w:author="Prior"','w:author="Other"'),target,{italic:true},{italic:true})).toThrow()
  expect(()=>verify(before,after.replace('w:id="9"','w:id="8"'),target,{italic:true},{italic:true})).toThrow()
})

test("inline review IDs remain unique across the whole main story", async () => {
  const { assertNativeInlineWordChatRevisionIds: verify } = await import("./word-chat-host-ooxml")
  const prior = '<w:p><w:ins w:id="3" w:author="Earlier" w:date="2026-01-01T00:00:00Z"><w:r><w:t>Earlier</w:t></w:r></w:ins></w:p>'
  const next = '<w:p><w:r><w:rPr><w:b/><w:rPrChange w:id="4" w:author="Native" w:date="2026-01-01T00:00:00Z"><w:rPr/></w:rPrChange></w:rPr><w:t>Target</w:t></w:r></w:p>'
  expect(() => verify(flat(prior+next))).not.toThrow()
  expect(() => verify(flat(prior+next.replace('w:id="4"','w:id="3"')))).toThrow("inline_revision_id_collision")
})
