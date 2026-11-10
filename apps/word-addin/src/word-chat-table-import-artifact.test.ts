import { expect, test } from "bun:test"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strToU8, unzipSync, zipSync } from "fflate"
import { verifyWordChatTableImportArtifact } from "./word-chat-table-import-artifact"

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", O = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const REL = "http://schemas.openxmlformats.org/package/2006/relationships", CT = "http://schemas.openxmlformats.org/package/2006/content-types", PKG = "http://schemas.microsoft.com/office/2006/xmlPackage"
const MAIN = "word/document.xml", MAIN_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
const xml = (node: Node) => new XMLSerializer().serializeToString(node)
const parse = (value: string) => new DOMParser().parseFromString(value, "application/xml")
const kids = (node: Node) => Array.from(node.childNodes).filter((n): n is Element => n.nodeType === 1)
const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`
const row = (texts: string[]) => `<w:tr>${texts.map(text => `<w:tc><w:tcPr><w:tcW w:w="1700" w:type="dxa"/></w:tcPr>${p(text)}</w:tc>`).join("")}</w:tr>`
const rel = (id: string, type: string, target: string) => `<Relationship Id="${id}" Type="${O}/${type}" Target="${target}"/>`
const rels = (content: string) => `<Relationships xmlns="${REL}">${content}</Relationships>`
let cached: ReturnType<typeof makeFixture> | undefined
async function makeFixture() {
  // Real server authoring is fixture setup, not part of the browser module's
  // compilation graph. Keep server-only ES2024 libraries out of Word's build.
  const writerPath = "../../../dochaus/lib/docx-table-row-deletion", builderPath = "../../../dochaus/lib/docx-table-row-deletion-import"
  type Request = { source_sha256: string; target_handle: string; author: string }
  const { authorDocxTableRowDeletion, docxTableRowDeletionTargets } = await import(writerPath) as {
    authorDocxTableRowDeletion: (source: Uint8Array, input: Request) => Promise<{ bytes: Uint8Array }>
    docxTableRowDeletionTargets: (source: Uint8Array) => Promise<{ source_sha256: string; targets: Array<{ row: number; target_handle: string }> }>
  }
  const { prepareDocxTableRowDeletionImport } = await import(builderPath) as {
    prepareDocxTableRowDeletionImport: (source: Uint8Array, authored: Uint8Array, input: Request) => Promise<{ beforeOoxml: string; ooxml: string }>
  }
  const fontType = "application/vnd.openxmlformats-officedocument.obfuscatedFont"
  const source = zipSync({
    "[Content_Types].xml": strToU8(`<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="odttf" ContentType="${fontType}"/><Override PartName="/${MAIN}" ContentType="${MAIN_TYPE}"/>${["styles", "settings", "fontTable", "numbering"].map(name => `<Override PartName="/word/${name}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${name}+xml"/>`).join("")}</Types>`),
    "_rels/.rels": strToU8(rels(rel("original-root", "officeDocument", MAIN))),
    "word/_rels/document.xml.rels": strToU8(rels(rel("s", "styles", "styles.xml") + rel("f", "fontTable", "fontTable.xml") + rel("n", "numbering", "numbering.xml") + rel("settings", "settings", "settings.xml"))),
    "word/styles.xml": strToU8(`<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="23"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`),
    "word/fontTable.xml": strToU8(`<w:fonts xmlns:w="${W}" xmlns:r="${O}"><w:font w:name="Example"><w:embedRegular r:id="font" w:fontKey="{11223344-5566-7788-99AA-BBCCDDEEFF00}"/></w:font></w:fonts>`),
    "word/_rels/fontTable.xml.rels": strToU8(rels(rel("font", "font", "fonts/example.odttf"))),
    "word/fonts/example.odttf": new Uint8Array([7, 8, 9]),
    "word/numbering.xml": strToU8(`<w:numbering xmlns:w="${W}" xmlns:r="${O}" xmlns:v="urn:schemas-microsoft-com:vml"><w:numPicBullet w:numPicBulletId="1"><w:pict><v:shape><v:imagedata r:id="picture"/></v:shape></w:pict></w:numPicBullet></w:numbering>`),
    "word/_rels/numbering.xml.rels": strToU8(rels(rel("picture", "image", "media/picture.png"))),
    "word/media/picture.png": new Uint8Array(Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/l9sAAAAASUVORK5CYII=", "base64")),
    "word/settings.xml": strToU8(`<w:settings xmlns:w="${W}"/>`),
    [MAIN]: strToU8(`<w:document xmlns:w="${W}"><w:body>${p("Outside opening")}<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1700"/><w:gridCol w:w="1700"/></w:tblGrid>${row(["Kind", "Meaning"])}${row(["Remove", "Entry"])}</w:tbl>${p("Outside ending")}<w:sectPr/></w:body></w:document>`),
  })
  const packet = await docxTableRowDeletionTargets(source), target = packet.targets.find(t => t.row === 2)!
  const request = { source_sha256: packet.source_sha256, target_handle: target.target_handle, author: "doc.haus" }
  const authored = await authorDocxTableRowDeletion(source, request), pair = await prepareDocxTableRowDeletionImport(source, authored.bytes, request)
  const paragraphs = ["Kind", "Meaning", "Remove", "Entry"].map((text, i) => ({ id: `cell-${i}`, text }))
  const proposal = { kind: "table_row_delete" as const, id: "reviewed", documentId: "document", explanation: "Remove row", openItems: [], createdAt: "now",
    table: { tableIndex: 0, rows: [paragraphs.slice(0, 2), paragraphs.slice(2)] }, rowIndex: 1, cells: ["Remove", "Entry"] }
  return { source: unzipSync(source), authored: unzipSync(authored.bytes), pair, proposal }
}
const fixture = () => cached ??= makeFixture()
function mutate(ooxml: string, name: string, change: (root: Element) => void) {
  const doc = parse(ooxml), part = kids(doc.documentElement).find(p => p.getAttributeNS(PKG, "name") === `/${name}`)!
  change(kids(kids(part)[0]!)[0]!); return xml(doc)
}
function drop(ooxml: string, name: string) {
  const doc = parse(ooxml), part = kids(doc.documentElement).find(p => p.getAttributeNS(PKG, "name") === `/${name}`)!
  part.parentNode!.removeChild(part); return xml(doc)
}

test("binds a real paired import to the verified source and authored complete table and implicit dependencies", async () => {
  const f = await fixture()
  expect(() => verifyWordChatTableImportArtifact(f.source, f.authored, f.pair, f.proposal)).not.toThrow()
  const pair = { ...f.pair, ooxml: mutate(f.pair.ooxml, "_rels/.rels", root => kids(root)[0]!.setAttribute("Id", "another-root")) }
  expect(() => verifyWordChatTableImportArtifact(f.source, f.authored, pair, f.proposal)).not.toThrow()
})

test("forged after tables cannot change header wording, geometry or deletion carriers", async () => {
  const f = await fixture()
  for (const change of [
    (root: Element) => root.getElementsByTagNameNS(W, "t")[0]!.appendChild(root.ownerDocument!.createTextNode(" forged")),
    (root: Element) => root.getElementsByTagNameNS(W, "gridCol")[0]!.setAttributeNS(W, "w:w", "1900"),
    (root: Element) => root.getElementsByTagNameNS(W, "del")[0]!.setAttributeNS(W, "w:author", "Someone else"),
  ]) expect(() => verifyWordChatTableImportArtifact(f.source, f.authored, { ...f.pair, ooxml: mutate(f.pair.ooxml, MAIN, change) }, f.proposal)).toThrow("part_xml_changed")
})

test("source preimages and document or body shells cannot be forged even when table text matches", async () => {
  const f = await fixture()
  const changedHeader = mutate(f.pair.beforeOoxml, MAIN, root => root.getElementsByTagNameNS(W, "t")[0]!.appendChild(root.ownerDocument!.createTextNode(" altered")))
  expect(() => verifyWordChatTableImportArtifact(f.source, f.authored, { ...f.pair, beforeOoxml: changedHeader }, f.proposal)).toThrow()
  for (const field of ["beforeOoxml", "ooxml"] as const) {
    const body = mutate(f.pair[field], MAIN, root => root.getElementsByTagNameNS(W, "body")[0]!.setAttributeNS("http://www.w3.org/XML/1998/namespace", "xml:space", "preserve"))
    const shell = mutate(f.pair[field], MAIN, root => root.setAttributeNS(W, "w:conformance", "strict"))
    for (const value of [body, shell]) expect(() => verifyWordChatTableImportArtifact(f.source, f.authored, { ...f.pair, [field]: value }, f.proposal)).toThrow()
  }
})

test("section properties, adjacent paragraphs and arbitrary parts never enter either payload", async () => {
  const f = await fixture()
  for (const field of ["beforeOoxml", "ooxml"] as const) {
    for (const name of ["p", "sectPr"]) {
      const leaked = mutate(f.pair[field], MAIN, root => root.getElementsByTagNameNS(W, "body")[0]!.appendChild(root.ownerDocument!.createElementNS(W, `w:${name}`)))
      expect(() => verifyWordChatTableImportArtifact(f.source, f.authored, { ...f.pair, [field]: leaked }, f.proposal)).toThrow()
    }
    const doc = parse(f.pair[field]), extra = doc.createElementNS(PKG, "pkg:part")
    extra.setAttributeNS(PKG, "pkg:name", "/word/comments.xml"); extra.setAttributeNS(PKG, "pkg:contentType", "application/xml")
    const data = doc.createElementNS(PKG, "pkg:xmlData"); data.appendChild(doc.createElementNS(W, "w:comments")); extra.appendChild(data); doc.documentElement.appendChild(extra)
    expect(() => verifyWordChatTableImportArtifact(f.source, f.authored, { ...f.pair, [field]: xml(doc) }, f.proposal)).toThrow("part_set_changed")
  }
})

test("omitting implicit defaults or transitive font and image dependencies cannot forge a smaller allowed closure", async () => {
  const f = await fixture()
  for (const name of ["word/styles.xml", "word/fonts/example.odttf", "word/media/picture.png", "word/_rels/numbering.xml.rels"]) {
    let pair = { beforeOoxml: drop(f.pair.beforeOoxml, name), ooxml: drop(f.pair.ooxml, name) }
    if(name === "word/styles.xml") {
      // The forged smaller pair is internally consistent and has no explicit
      // table style reference. Source implicit defaults still MUST travel.
      const removeLink = (root: Element) => { const link=kids(root).find(n=>n.getAttribute("Id")==="s")!;root.removeChild(link) }
      pair={beforeOoxml:mutate(pair.beforeOoxml,"word/_rels/document.xml.rels",removeLink),ooxml:mutate(pair.ooxml,"word/_rels/document.xml.rels",removeLink)}
    }
    expect(() => verifyWordChatTableImportArtifact(f.source, f.authored, pair, f.proposal)).toThrow("part_set_changed")
  }
})

test("changed formatting definitions, binary data, relationship modes and content types refuse", async () => {
  const f = await fixture()
  for (const field of ["beforeOoxml", "ooxml"] as const) {
    const style = mutate(f.pair[field], "word/styles.xml", root => root.getElementsByTagNameNS(W, "sz")[0]!.setAttributeNS(W, "w:val", "31"))
    const external = mutate(f.pair[field], "word/_rels/numbering.xml.rels", root => kids(root)[0]!.setAttribute("TargetMode", "External"))
    const missing = mutate(f.pair[field], "word/_rels/numbering.xml.rels", root => kids(root)[0]!.setAttribute("Target", "media/missing.png"))
    for (const value of [style, external, missing]) expect(() => verifyWordChatTableImportArtifact(f.source, f.authored, { ...f.pair, [field]: value }, f.proposal)).toThrow()
    const doc = parse(f.pair[field]), font = kids(doc.documentElement).find(p => p.getAttributeNS(PKG, "name") === "/word/fonts/example.odttf")!
    kids(font)[0]!.textContent = "AAAA"
    expect(() => verifyWordChatTableImportArtifact(f.source, f.authored, { ...f.pair, [field]: xml(doc) }, f.proposal)).toThrow("part_bytes_changed")
    font.setAttributeNS(PKG, "pkg:contentType", "application/xml")
    expect(() => verifyWordChatTableImportArtifact(f.source, f.authored, { ...f.pair, [field]: xml(doc) }, f.proposal)).toThrow()
  }
})

test("unresolvable source dependencies and oversized pairs are refused", async () => {
  const f = await fixture(), bad = { ...f.source }
  bad["word/_rels/numbering.xml.rels"] = strToU8(rels(rel("picture", "image", "media/missing.png")))
  expect(() => verifyWordChatTableImportArtifact(bad, f.authored, f.pair, f.proposal)).toThrow("dependency_missing")
  const external = parse(rels(rel("picture", "image", "media/picture.png"))); kids(external.documentElement)[0]!.setAttribute("TargetMode", "External")
  bad["word/_rels/numbering.xml.rels"] = strToU8(xml(external))
  expect(() => verifyWordChatTableImportArtifact(bad, f.authored, f.pair, f.proposal)).toThrow("external_dependency")
  expect(() => verifyWordChatTableImportArtifact(f.source, f.authored, { beforeOoxml: "x".repeat(1024 * 1024 + 1), ooxml: "x".repeat(1024 * 1024) }, f.proposal)).toThrow("pair_large")
})
