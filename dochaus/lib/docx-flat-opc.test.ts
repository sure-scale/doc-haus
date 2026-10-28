import { expect, test } from "bun:test"
import { DOMParser } from "@xmldom/xmldom"
import { authorWordChatTarget, MAX_WORD_AMENDMENT_OOXML_BYTES, WordFormattingError } from "./docx-flat-opc"
import type { WordChatTarget } from "./word-chat-amendment-contract"
import { reviewedWordChatArtifact, sameImportedWordChat } from "../../apps/word-addin/src/word-chat-host-ooxml"

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const PKG = "http://schemas.microsoft.com/office/2006/xmlPackage"
const REL = "http://schemas.openxmlformats.org/package/2006/relationships"
const O = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const escape = (s: string) => s.replaceAll("&", "&amp;").replaceAll("<", "&lt;")
const paragraph = (s: string) => `<w:p><w:r><w:t xml:space="preserve">${escape(s)}</w:t></w:r></w:p>`
function flat(body: string, extra = "", relations = "") {
  return `<pkg:package xmlns:pkg="${PKG}" xmlns:w="${W}" xmlns:r="${O}"><pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData><Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${O}/officeDocument" Target="word/document.xml"/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData><w:document><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document></pkg:xmlData></pkg:part>${relations ? `<pkg:part pkg:name="/word/_rels/document.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData><Relationships xmlns="${REL}">${relations}</Relationships></pkg:xmlData></pkg:part>` : ""}${extra}</pkg:package>`
}
function target(texts: string[], start = 0, end = texts.at(-1)!.length): WordChatTarget {
  return { id: "local-range", text: texts.map((text, i) => text.slice(i === 0 ? start : 0, i === texts.length - 1 ? end : undefined)).join("\r"),
    paragraphs: texts.map((text, i) => ({ id: `native-${i}`, text })), startOffset: start, endOffset: end,
    saved: false, capturedAt: "2026-01-01T00:00:00Z" }
}
const parsed = (xml: string) => new DOMParser().parseFromString(xml, "application/xml")

test("shared authoring preserves inherited namespaces/run formatting and exact selected occurrence", async () => {
  const xml = flat('<w:p><w:pPr><w:spacing w:after="120"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Monday Monday</w:t></w:r></w:p>')
  const result = await authorWordChatTarget(xml, target(["Monday Monday"], 7, 13), "Tuesday")
  expect(result.paragraphs).toEqual(["Monday Tuesday"])
  expect(result.author).toBe("doc.haus")
  expect(result.revisions).toEqual([{ type: "Deleted", text: "Monday", author: "doc.haus" }, { type: "Added", text: "Tuesday", author: "doc.haus" }])
  const doc = parsed(result.ooxml)
  expect(doc.getElementsByTagNameNS(W, "p")).toHaveLength(1)
  expect(doc.getElementsByTagNameNS(W, "b").length).toBeGreaterThan(0)
  expect(doc.getElementsByTagNameNS(W, "spacing")[0]!.getAttributeNS(W, "after")).toBe("120")
  expect(doc.getElementsByTagNameNS(W, "sectPr")).toHaveLength(0)
})

test("same-count multiline authoring preserves outside selected prefixes/suffixes and unchanged paragraphs", async () => {
  const texts = ["Lead Monday tail", "Keep this paragraph.", "Other Monday tail"], xml = flat(texts.map(paragraph).join(""))
  const result = await authorWordChatTarget(xml, target(texts, 5, 5), "Tuesday tail\rKeep this paragraph.\rFinal")
  expect(result.paragraphs).toEqual(["Lead Tuesday tail", "Keep this paragraph.", "Final Monday tail"])
  expect(result.revisions.every(r => r.author === "doc.haus")).toBe(true)
  expect(result.revisions.map(r => r.text).join(" ")).not.toContain("Keep this paragraph")
  expect(parsed(result.ooxml).getElementsByTagNameNS(W, "p")).toHaveLength(3)
})

test("reviewed paragraph insertion and deletion use shared tracked authoring with original numbering and runs intact", async () => {
  const texts = ["Working note Tuesday.", "Unrelated BLUE FERN."], note = "Working notes — non-operative — retain, move or delete:"
  const body = '<w:p><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="7"/></w:numPr><w:spacing w:after="120"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Working note Tuesday.</w:t></w:r></w:p>' + paragraph(texts[1]!)
  const xml = flat(body), selected = target(texts), inserted = await authorWordChatTarget(xml, selected, note + "\r" + selected.text)
  expect(inserted.paragraphs).toEqual([note, ...texts])
  expect(inserted.revisions).toEqual([{ type: "Added", author: "doc.haus", text: note }, { type: "Added", author: "doc.haus", text: "\r" }])
  const authored = { ...inserted, proposalId: "add-note" }
  expect(reviewedWordChatArtifact(authored, "add-note", selected, inserted.paragraphs, xml)).toHaveLength(3)
  expect(sameImportedWordChat(inserted.ooxml, inserted.ooxml)).toBe(true)
  const doc = parsed(inserted.ooxml)
  expect(doc.getElementsByTagNameNS(W, "ilvl")[0]!.getAttributeNS(W, "val")).toBe("1")
  expect(doc.getElementsByTagNameNS(W, "numId")[0]!.getAttributeNS(W, "val")).toBe("7")
  expect(doc.getElementsByTagNameNS(W, "b")).toHaveLength(1)
  const deleted = await authorWordChatTarget(xml, selected, texts[1]!)
  expect(deleted.paragraphs).toEqual([texts[1]!])
  expect(deleted.revisions).toEqual([{ type: "Deleted", author: "doc.haus", text: texts[0]! }, { type: "Deleted", author: "doc.haus", text: "\r" }])
  expect(reviewedWordChatArtifact({ ...deleted, proposalId: "remove-note" }, "remove-note", selected, deleted.paragraphs, xml)).toHaveLength(2)
  // Losing the paragraph-mark revision or its attribution is a failed import,
  // even when all visible words and ordinary run revisions still match.
  expect(sameImportedWordChat(inserted.ooxml, inserted.ooxml.replace(/<w:rPr><w:ins[^>]*\/><\/w:rPr>/, ""))).toBe(false)
  expect(() => reviewedWordChatArtifact({ ...authored, ooxml: authored.ooxml.replace('<w:spacing w:after="120"', '<w:spacing w:after="240"') }, "add-note", selected, inserted.paragraphs, xml)).toThrow("formatting")
})

test("several inserted paragraphs keep reviewed punctuation literal and ordered between retained paragraphs", async () => {
  const texts = ["Retained opening.", "Retained ending."], selected = target(texts)
  const inserted = ["# A literal *note* [reference](target).", "- Another note with \\ punctuation."]
  const expected = [texts[0]!, ...inserted, texts[1]!]
  const xml = flat(texts.map(paragraph).join("")), result = await authorWordChatTarget(xml, selected, expected.join("\r"))
  expect(result.paragraphs).toEqual(expected)
  expect(reviewedWordChatArtifact({ ...result, proposalId: "several" }, "several", selected, expected, xml)).toHaveLength(4)
  expect(result.revisions.filter(r => r.text !== "\r").map(r => r.text)).toEqual(inserted)
  expect(parsed(result.ooxml).getElementsByTagNameNS(W, "hyperlink")).toHaveLength(0)
})

const formatStyles = `<pkg:part pkg:name="/word/styles.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"><pkg:xmlData><w:styles><w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/></w:rPr></w:style></w:styles></pkg:xmlData></pkg:part>`
const styleRelation = `<Relationship Id="rStyle" Type="${O}/styles" Target="styles.xml"/>`

test("wording and formatting preserve a separate wording artifact", async () => {
  const texts = ["The parties shall deliver the materials promptly.", "This qualification remains unchanged."], selected = target(texts)
  const xml = flat(texts.map(paragraph).join(""), formatStyles, styleRelation)
  const expected = ["The parties shall deliver the materials promptly and in writing.", texts[1]!]
  const result = await authorWordChatTarget(xml, selected, expected.join("\r"), [{ paragraph: 0, spacingAfter: 6 }])
  expect(result.paragraphs).toEqual(expected)
  expect(result.wording).toBeDefined()
  expect(result.revisions.some(r => r.type === "Formatted" && r.text === expected[0])).toBe(true)
  expect(parsed(result.ooxml).getElementsByTagNameNS(W, "pPrChange")).toHaveLength(1)
  expect(parsed(result.wording!.ooxml).getElementsByTagNameNS(W, "pPrChange")).toHaveLength(0)
})

test("mixed wording leaves earlier inline revisions in an untouched paragraph", async () => {
  const texts = ["The obligation is ten days.", "Deliver the materials promptly.", "Retained qualification."], selected = target(texts)
  const body = '<w:p><w:r><w:t>The obligation is </w:t></w:r><w:del w:id="42" w:author="Earlier reviewer" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>five</w:delText></w:r></w:del><w:ins w:id="43" w:author="Earlier reviewer" w:date="2026-01-01T00:00:00Z"><w:r><w:t>ten</w:t></w:r></w:ins><w:r><w:t> days.</w:t></w:r></w:p>' + texts.slice(1).map(paragraph).join("")
  const xml = flat(body, formatStyles, styleRelation), expected = [texts[0]!, "Deliver the materials promptly and in writing.", texts[2]!]
  const result = await authorWordChatTarget(xml, selected, expected.join("\r"), [{ paragraph: 1, spacingAfter: 6 }])
  for (const tag of ["ins", "del"]) expect(Array.from(parsed(result.ooxml).getElementsByTagNameNS(W, tag)).some(e => e.getAttributeNS(W, "author") === "Earlier reviewer")).toBe(true)
  await expect(authorWordChatTarget(xml, selected, ["The obligation is twenty days.", texts[1]!, texts[2]!].join("\r"), [{ paragraph: 1, spacingAfter: 6 }])).rejects.toThrow()
})

test("layered wording changes beside an earlier own revision and refuses overlap", async () => {
  const texts = ["A will supply B Market.", "Retained qualification."], selected = target(texts)
  const body = '<w:p><w:r><w:t>A will supply B </w:t></w:r><w:del w:id="42" w:author="doc.haus" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>Street</w:delText></w:r></w:del><w:ins w:id="43" w:author="doc.haus" w:date="2026-01-01T00:00:00Z"><w:r><w:t>Market</w:t></w:r></w:ins><w:r><w:t>.</w:t></w:r></w:p>' + paragraph(texts[1]!)
  const xml = flat(body, formatStyles, styleRelation)
  const result = await authorWordChatTarget(xml, selected, "A shall supply B Market.\r" + texts[1], [{ paragraph: 0, spacingAfter: 6 }])
  const doc = parsed(result.ooxml)
  for (const [tag, id] of [["del", "42"], ["ins", "43"]]) {
    const prior = Array.from(doc.getElementsByTagNameNS(W, tag)).find(e => e.getAttributeNS(W, "id") === id)
    expect(prior?.getAttributeNS(W, "author")).toBe("doc.haus")
    expect(prior?.getAttributeNS(W, "date")).toBe("2026-01-01T00:00:00Z")
    expect(prior?.textContent).toBe(tag === "del" ? "Street" : "Market")
  }
  expect(result.revisions.filter(r => r.type !== "Formatted").map(r => r.text).join(" ")).not.toContain("Market")
  const wordingOnly = await authorWordChatTarget(xml, selected, "A shall supply B Market.\r" + texts[1])
  expect(wordingOnly.revisions.map(r => r.text).join(" ")).not.toContain("Market")
  expect(Array.from(parsed(wordingOnly.ooxml).getElementsByTagNameNS(W, "ins")).some(e => e.getAttributeNS(W, "id") === "43")).toBe(true)
  await expect(authorWordChatTarget(xml, selected, "A will supply B Square.\r" + texts[1], [{ paragraph: 0, spacingAfter: 6 }])).rejects.toThrow("docx_own_revision")
})

test("inserted heading formatting travels with its tracked insertion", async () => {
  const texts = ["The parties shall deliver the materials promptly.", "This qualification remains unchanged."], selected = target(texts)
  const xml = flat(texts.map(paragraph).join(""), formatStyles, styleRelation)
  const expected = ["Delivery", "The parties shall deliver the materials promptly and in writing.", texts[1]!]
  const result = await authorWordChatTarget(xml, selected, expected.join("\r"), [{ paragraph: 0, style: "Heading1" }, { paragraph: 1, spacingAfter: 6 }])
  const doc = parsed(result.ooxml)
  expect(doc.getElementsByTagNameNS(W, "pStyle")[0]!.getAttributeNS(W, "val")).toBe("Heading1")
  expect(doc.getElementsByTagNameNS(W, "pPrChange")).toHaveLength(1)
  expect(result.revisions.some(r => r.type === "Formatted" && r.text === "Delivery")).toBe(true)
  const wording = parsed(result.wording!.ooxml)
  expect(wording.getElementsByTagNameNS(W, "pStyle")[0]!.getAttributeNS(W, "val")).toBe("Heading1")
  expect(parsed(result.wording!.unformattedOoxml!).getElementsByTagNameNS(W, "pStyle")).toHaveLength(0)
  expect(wording.getElementsByTagNameNS(W, "pPrChange")).toHaveLength(0)
  expect(wording.getElementsByTagNameNS(W, "spacing")).toHaveLength(0)
  expect(result.wording!.revisions).toEqual(result.revisions.filter(revision => revision.type !== "Formatted"))
})

test("combined paragraph formatting preserves inline text revisions and proves native accept/reject properties", async () => {
  const texts = ["The obligation is ten days.", "Retained qualification."], selected = target(texts)
  const body = '<w:p><w:pPr><w:spacing w:after="120"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>The obligation is </w:t></w:r><w:del w:id="42" w:author="Earlier reviewer" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>five</w:delText></w:r></w:del><w:ins w:id="43" w:author="Earlier reviewer" w:date="2026-01-01T00:00:00Z"><w:r><w:t>ten</w:t></w:r></w:ins><w:r><w:t> days.</w:t></w:r></w:p>' + paragraph(texts[1]!)
  const xml = flat(body, formatStyles, styleRelation), formatting = [{ paragraph: 0, style: "Heading1" as const, spacingBefore: 12, spacingAfter: 8, lineSpacing: 1.25 }, { paragraph: 1, spacingAfter: 6 }]
  const result = await authorWordChatTarget(xml, selected, selected.text, formatting)
  expect(result.paragraphs).toEqual(texts)
  expect(result.formatting).toEqual(formatting)
  expect(result.revisions).toEqual(texts.map(text => ({ author: "doc.haus", type: "Formatted", text })))
  const doc = parsed(result.ooxml), changes = Array.from(doc.getElementsByTagNameNS(W, "pPrChange"))
  expect(changes).toHaveLength(2)
  expect(changes.every(node => node.getAttributeNS(W, "author") === "doc.haus")).toBe(true)
  expect(doc.getElementsByTagNameNS(W, "rPrChange")).toHaveLength(0)
  for (const tag of ["ins", "del"]) {
    const existing = doc.getElementsByTagNameNS(W, tag)
    expect(existing).toHaveLength(1)
    expect(existing[0]!.getAttributeNS(W, "author")).toBe("Earlier reviewer")
    expect(existing[0]!.getAttributeNS(W, "date")).toBe("2026-01-01T00:00:00Z")
  }
  const authored = { ...result, proposalId: "format" }
  expect(reviewedWordChatArtifact(authored, "format", selected, texts, xml, formatting).flat()).toEqual(result.revisions)
  expect(sameImportedWordChat(result.ooxml, result.ooxml)).toBe(true)
  expect(() => reviewedWordChatArtifact({ ...authored, ooxml: authored.ooxml.replace(/<w:b\s*\/>/, '<w:i/>') }, "format", selected, texts, xml, formatting)).toThrow("inline")
  expect(() => reviewedWordChatArtifact({ ...authored, ooxml: authored.ooxml.replace('w:after="160"', 'w:after="180"') }, "format", selected, texts, xml, formatting)).toThrow("spacing")
  expect(() => reviewedWordChatArtifact({ ...authored, ooxml: authored.ooxml.replace('w:author="Earlier reviewer"', 'w:author="doc.haus"') }, "format", selected, texts, xml, formatting)).toThrow("inline")
})

test("formatting refuses unavailable styles, partial selections and changes that remove inherited numbering", async () => {
  const selected = target(["A numbered paragraph."])
  const styles = formatStyles.replace('</w:styles>', '<w:style w:type="paragraph" w:styleId="Numbered"><w:name w:val="Numbered"/><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="7"/></w:numPr></w:pPr></w:style></w:styles>')
  const numbering = `<pkg:part pkg:name="/word/numbering.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"><pkg:xmlData><w:numbering><w:abstractNum w:abstractNumId="7"><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="(%2)"/></w:lvl></w:abstractNum><w:num w:numId="7"><w:abstractNumId w:val="7"/></w:num></w:numbering></pkg:xmlData></pkg:part>`
  const xml = flat('<w:p><w:pPr><w:pStyle w:val="Numbered"/></w:pPr><w:r><w:t>A numbered paragraph.</w:t></w:r></w:p>', styles + numbering, styleRelation + `<Relationship Id="rNumber" Type="${O}/numbering" Target="numbering.xml"/>`)
  await expect(authorWordChatTarget(xml, selected, selected.text, [{ paragraph: 0, style: "Heading2" }])).rejects.toBeInstanceOf(WordFormattingError)
  await expect(authorWordChatTarget(xml, selected, selected.text, [{ paragraph: 0, style: "Heading1" }])).rejects.toThrow("numbering")
  const result = await authorWordChatTarget(xml, selected, selected.text, [{ paragraph: 0, spacingAfter: 8 }])
  expect(result.revisions).toEqual([{ author: "doc.haus", type: "Formatted", text: selected.text }])
  expect(parsed(result.ooxml).getElementsByTagNameNS(W, "pStyle")[0]!.getAttributeNS(W, "val")).toBe("Numbered")
  const direct = xml.replace('<w:pStyle w:val="Numbered"/>', '<w:numPr><w:ilvl w:val="1"/><w:numId w:val="7"/></w:numPr>')
  const heading = await authorWordChatTarget(direct, selected, selected.text, [{ paragraph: 0, style: "Heading1", spacingAfter: 8 }])
  expect(parsed(heading.ooxml).getElementsByTagNameNS(W, "numId")[0]!.getAttributeNS(W, "val")).toBe("7")
  expect(parsed(heading.ooxml).getElementsByTagNameNS(W, "ilvl")[0]!.getAttributeNS(W, "val")).toBe("1")
  expect(parsed(heading.ooxml).getElementsByTagNameNS(W, "pStyle")[0]!.getAttributeNS(W, "val")).toBe("Heading1")
  const partial = target(["A numbered paragraph."], 2)
  await expect(authorWordChatTarget(xml, partial, partial.text, [{ paragraph: 0, spacingAfter: 8 }])).rejects.toThrow("word_snapshot_unsupported")
})

test("formatting dependencies remain while unrelated header projections/settings/add-in context are excluded", async () => {
  const extra = `<pkg:part pkg:name="/word/header1.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"><pkg:xmlData><w:hdr>${paragraph("same")}</w:hdr></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/styles.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"><pkg:xmlData><w:styles><w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles></pkg:xmlData></pkg:part>`
  const rels = `<Relationship Id="rHead" Type="${O}/header" Target="header1.xml"/><Relationship Id="rStyle" Type="${O}/styles" Target="styles.xml"/>`
  const result = await authorWordChatTarget(flat(paragraph("same"), extra, rels), target(["same"]), "different")
  expect(result.paragraphs).toEqual(["different"])
  const names = Array.from(parsed(result.ooxml).getElementsByTagNameNS(PKG, "part")).map(p => p.getAttributeNS(PKG, "name"))
  expect(names).toContain("/word/styles.xml")
  expect(names).not.toContain("/word/header1.xml")
  expect(names).not.toContain("/word/settings.xml")
  expect(result.ooxml).not.toContain('Target="header1.xml"')
})

test("delete-only selected wording remains a native deletion and soft breaks/tabs stay in place", async () => {
  const deleted = await authorWordChatTarget(flat(paragraph("Keep obsolete wording.")), target(["Keep obsolete wording."], 5, 14), "")
  expect(deleted.paragraphs).toEqual(["Keep wording."])
  expect(deleted.revisions).toEqual([{ type: "Deleted", text: "obsolete ", author: "doc.haus" }])
  const text = "Keep\tMonday\vwhen agreed."
  const xml = flat('<w:p><w:r><w:t>Keep</w:t><w:tab/><w:t>Monday</w:t><w:br/><w:t>when agreed.</w:t></w:r></w:p>')
  const result = await authorWordChatTarget(xml, target([text]), "Keep\tTuesday\vwhen agreed.")
  expect(result.paragraphs).toEqual(["Keep\tTuesday\vwhen agreed."])
  expect(parsed(result.ooxml).getElementsByTagNameNS(W, "tab")).toHaveLength(1)
  expect(parsed(result.ooxml).getElementsByTagNameNS(W, "br")).toHaveLength(1)
  await expect(authorWordChatTarget(xml, target([text]), "Keep Tuesday when agreed.")).rejects.toThrow("word_snapshot_unsupported")
})

test("picture-bullet numbering retains its internal image and relationship dependency", async () => {
  const image = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jV2kAAAAASUVORK5CYII="
  const extra = `<pkg:part pkg:name="/word/numbering.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"><pkg:xmlData><w:numbering xmlns:v="urn:schemas-microsoft-com:vml"><w:numPicBullet w:numPicBulletId="1"><w:pict><v:shape style="width:9pt;height:9pt"><v:imagedata r:id="rBullet"/></v:shape></w:pict></w:numPicBullet><w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="singleLevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlPicBulletId w:val="1"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num></w:numbering></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/_rels/numbering.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData><Relationships xmlns="${REL}"><Relationship Id="rBullet" Type="${O}/image" Target="media/bullet.png"/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/media/bullet.png" pkg:contentType="image/png"><pkg:binaryData>${image}</pkg:binaryData></pkg:part>`
  const body = '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Monday</w:t></w:r></w:p>'
  const result = await authorWordChatTarget(flat(body, extra, `<Relationship Id="rNumbering" Type="${O}/numbering" Target="numbering.xml"/>`), target(["Monday"]), "Tuesday")
  expect(result.paragraphs).toEqual(["Tuesday"])
  const doc = parsed(result.ooxml), parts = Array.from(doc.getElementsByTagNameNS(PKG, "part"))
  const part = (name: string) => parts.find(p => p.getAttributeNS(PKG, "name") === name)!
  expect(part("/word/numbering.xml")).toBeDefined()
  expect(part("/word/_rels/numbering.xml.rels").getElementsByTagNameNS(REL, "Relationship")[0]!.getAttribute("Target")).toBe("media/bullet.png")
  expect(part("/word/media/bullet.png").getElementsByTagNameNS(PKG, "binaryData")[0]!.textContent).toBe(image)
  expect(part("/word/media/bullet.png").getAttributeNS(PKG, "contentType")).toBe("image/png")
})

test("paragraph tab-stop properties are formatting while actual run tabs remain text", async () => {
  const xml = flat('<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="720"/><w:tab w:val="right" w:pos="4320"/></w:tabs></w:pPr><w:r><w:t>Keep</w:t><w:tab/><w:t>Monday</w:t></w:r></w:p>')
  const result = await authorWordChatTarget(xml, target(["Keep\tMonday"]), "Keep\tTuesday")
  expect(result.paragraphs).toEqual(["Keep\tTuesday"])
  const doc = parsed(result.ooxml), tabs = Array.from(doc.getElementsByTagNameNS(W, "tab"))
  expect(tabs).toHaveLength(3)
  expect(tabs.filter(tab => tab.getAttributeNS(W, "pos")).map(tab => tab.getAttributeNS(W, "pos"))).toEqual(["720", "4320"])
  expect(tabs.filter(tab => !tab.getAttributeNS(W, "pos"))).toHaveLength(1)
})

test("one complete cell becomes paragraph-only delivery while extra table paragraphs cannot masquerade as the target", async () => {
  const xml = flat(`<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr/>${paragraph("Monday")}</w:tc></w:tr></w:tbl>`)
  const result = await authorWordChatTarget(xml, target(["Monday"]), "Tuesday")
  expect(result.paragraphs).toEqual(["Tuesday"])
  expect(parsed(result.ooxml).getElementsByTagNameNS(W, "tbl")).toHaveLength(0)
  const twoCells = xml.replace('</w:tc></w:tr>', `</w:tc><w:tc>${paragraph("Unselected")}</w:tc></w:tr>`)
  await expect(authorWordChatTarget(twoCells, target(["Monday"]), "Tuesday")).rejects.toThrow("word_snapshot_unsupported")
})

test("changed text/offsets, protected content, structural changes and oversized snapshots refuse without authoring", async () => {
  const xml = flat(paragraph("Monday"))
  await expect(authorWordChatTarget(xml, target(["Tuesday"]), "Wednesday")).rejects.toThrow("word_snapshot_changed")
  await expect(authorWordChatTarget(xml, { ...target(["Monday"]), startOffset: 2 }, "Tuesday")).rejects.toThrow("word_snapshot_changed")
  await expect(authorWordChatTarget(xml, target(["Monday"]), "Tuesday\rWednesday")).rejects.toThrow("word_snapshot_unsupported")
  for (const content of ['<w:r><w:fldChar w:fldCharType="begin"/><w:t>Monday</w:t></w:r>', '<w:commentRangeStart w:id="1"/><w:r><w:t>Monday</w:t></w:r>']) {
    await expect(authorWordChatTarget(flat(`<w:p>${content}</w:p>`), target(["Monday"]), "Tuesday")).rejects.toThrow("word_snapshot_unsupported")
  }
  await expect(authorWordChatTarget(flat('<w:p><w:ins w:id="42" w:author="Someone" w:date="2026-01-01T00:00:00Z"><w:r><w:t>Monday</w:t></w:r></w:ins></w:p>'), target(["Monday"]), "Tuesday")).rejects.toThrow("docx_pending_revision")
  await expect(authorWordChatTarget(" ".repeat(MAX_WORD_AMENDMENT_OOXML_BYTES + 1), target(["Monday"]), "Tuesday")).rejects.toThrow("word_snapshot_large")
})
