import { expect, test } from "bun:test"
import { DOMParser } from "@xmldom/xmldom"
import { authorWordChatTarget, MAX_WORD_AMENDMENT_OOXML_BYTES } from "./docx-flat-opc"
import type { WordChatTarget } from "./word-chat-amendment-contract"

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
  for (const content of ['<w:ins w:author="Someone"><w:r><w:t>Monday</w:t></w:r></w:ins>', '<w:r><w:fldChar w:fldCharType="begin"/><w:t>Monday</w:t></w:r>', '<w:commentRangeStart w:id="1"/><w:r><w:t>Monday</w:t></w:r>']) {
    await expect(authorWordChatTarget(flat(`<w:p>${content}</w:p>`), target(["Monday"]), "Tuesday")).rejects.toThrow("word_snapshot_unsupported")
  }
  await expect(authorWordChatTarget(" ".repeat(MAX_WORD_AMENDMENT_OOXML_BYTES + 1), target(["Monday"]), "Tuesday")).rejects.toThrow("word_snapshot_large")
})
