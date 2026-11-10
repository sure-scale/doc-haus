import { expect, test } from "bun:test"
import { createHash } from "node:crypto"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { authorDocxTableRowDeletion, docxTableRowDeletionTargets, resolveDocxTableRowDeletion, verifyDocxTableRowDeletion, type DocxTableRowDeletion } from "./docx-table-row-deletion"
import { prepareDocxTableRowDeletionImport } from "./docx-table-row-deletion-import"

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", O = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const REL = "http://schemas.openxmlformats.org/package/2006/relationships", CT = "http://schemas.openxmlformats.org/package/2006/content-types", PKG = "http://schemas.microsoft.com/office/2006/xmlPackage"
const A = "http://schemas.openxmlformats.org/drawingml/2006/main", V = "urn:schemas-microsoft-com:vml", MAIN = "word/document.xml", SETTINGS = "word/settings.xml"
const parse = (value: string | Uint8Array) => new DOMParser().parseFromString(typeof value === "string" ? value : strFromU8(value), "application/xml")
const xml = (node: Node) => new XMLSerializer().serializeToString(node)
const kids = (node: Node) => Array.from(node.childNodes).filter((n): n is Element => n.nodeType === 1)
const direct = (node: Node, name: string) => kids(node).filter(n => n.namespaceURI === W && n.localName === name)
const paragraph = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`
const cell = (body: string, width: number) => `<w:tc><w:tcPr><w:tcW w:type="dxa" w:w="${width}"/></w:tcPr>${body}</w:tc>`
const row = (a: string, b: string) => `<w:tr>${cell(a, 2100)}${cell(b, 3300)}</w:tr>`
const relationship = (id: string, kind: string, target: string) => `<Relationship Id="${id}" Type="${O}/${kind}" Target="${target}"/>`
const relationships = (value: string) => `<Relationships xmlns="${REL}">${value}</Relationships>`
function fixture({ trailing = true, leadingTable = false } = {}) {
  const prior = `<w:p><w:ins w:id="40" w:author="Earlier Editor" w:date="2026-02-04T10:00:00Z"><w:r><w:rPr><w:b/></w:rPr><w:t>Reviewed header</w:t></w:r></w:ins></w:p>`
  const mixed = `<w:p><w:pPr><w:spacing w:after="160"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Heavy</w:t></w:r><w:r><w:t xml:space="preserve"> plain </w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>slanted</w:t></w:r></w:p>`
  const table = (old: string) => `<w:tbl w:rsidR="00112233"><w:tblPr><w:tblW w:type="dxa" w:w="5400"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="2100"/><w:gridCol w:w="3300"/></w:tblGrid>${row(old, paragraph("Stable heading"))}${row(mixed, paragraph("Removed entry"))}${row(paragraph("Stable ending"), paragraph("Peer entry"))}</w:tbl>`
  const parts: Record<string, Uint8Array> = {
    [MAIN]: strToU8(`<w:document xmlns:w="${W}" xmlns:r="${O}" xmlns:a="${A}" xmlns:v="${V}"><w:body><w:p><w:commentRangeStart w:id="9"/><w:r><w:t>Opening outside table</w:t></w:r><w:commentRangeEnd w:id="9"/><w:r><w:commentReference w:id="9"/></w:r></w:p>${leadingTable ? table(paragraph("Other header")) : ""}${table(prior)}${trailing ? paragraph("Ending outside table") : ""}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`),
    [SETTINGS]: strToU8(`<w:settings xmlns:w="${W}"/>`),
    "word/styles.xml": strToU8(`<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Example Serif"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`),
    "word/fontTable.xml": strToU8(`<w:fonts xmlns:w="${W}" xmlns:r="${O}"><w:font w:name="Example Serif"><w:embedRegular r:id="font" w:fontKey="{00112233-4455-6677-8899-AABBCCDDEEFF}"/></w:font></w:fonts>`),
    "word/theme/theme1.xml": strToU8(`<a:theme xmlns:a="${A}" name="Simple"><a:themeElements><a:clrScheme name="Plain"/><a:fontScheme name="Plain"/><a:fmtScheme name="Plain"/></a:themeElements></a:theme>`),
    "word/numbering.xml": strToU8(`<w:numbering xmlns:w="${W}" xmlns:r="${O}" xmlns:v="${V}"><w:numPicBullet w:numPicBulletId="3"><w:pict><v:shape><v:imagedata r:id="picture"/></v:shape></w:pict></w:numPicBullet></w:numbering>`),
    "word/fonts/example.odttf": new Uint8Array([2, 4, 6, 8]),
    "word/media/bullet.png": new Uint8Array(Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/l9sAAAAASUVORK5CYII=", "base64")),
    "word/comments.xml": strToU8(`<w:comments xmlns:w="${W}"><w:comment w:id="9" w:author="Outside Reviewer">${paragraph("Keep the outside review note.")}</w:comment></w:comments>`),
    "customXml/item1.xml": strToU8('<private xmlns="urn:synthetic:outside">Excluded metadata</private>'),
    "word/webextensions/taskpanes.xml": strToU8('<taskpanes xmlns="urn:synthetic:taskpanes"/>'),
    "_rels/.rels": strToU8(relationships(relationship("main", "officeDocument", MAIN))),
    "word/_rels/document.xml.rels": strToU8(relationships(relationship("styles", "styles", "styles.xml") + relationship("fonts", "fontTable", "fontTable.xml") + relationship("theme", "theme", "theme/theme1.xml") + relationship("numbering", "numbering", "numbering.xml") + relationship("settings", "settings", "settings.xml") + relationship("comments", "comments", "comments.xml") + relationship("custom", "customXml", "../customXml/item1.xml"))),
    "word/_rels/fontTable.xml.rels": strToU8(relationships(relationship("font", "font", "fonts/example.odttf"))),
    "word/_rels/numbering.xml.rels": strToU8(relationships(relationship("picture", "image", "media/bullet.png"))),
  }
  const type = (name: string) => name === MAIN ? "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
    : name === "word/theme/theme1.xml" ? "application/vnd.openxmlformats-officedocument.theme+xml"
    : ["styles", "fontTable", "numbering", "settings", "comments"].some(n => name === `word/${n}.xml`) ? `application/vnd.openxmlformats-officedocument.wordprocessingml.${name.slice(5, -4)}+xml` : "application/xml"
  parts["[Content_Types].xml"] = strToU8(`<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="odttf" ContentType="application/vnd.openxmlformats-officedocument.obfuscatedFont"/>${Object.keys(parts).filter(n => n.endsWith(".xml")).map(n => `<Override PartName="/${n}" ContentType="${type(n)}"/>`).join("")}</Types>`)
  return zipSync(parts, { mtime: new Date(1980, 0, 1) })
}
function changed(bytes: Uint8Array, action: (doc: Document) => void, part = MAIN) {
  const parts = unzipSync(bytes), doc = parse(parts[part]!); action(doc); parts[part] = strToU8(xml(doc)); return zipSync(parts)
}
function tableAt(doc: Document, index = 1) { return direct(direct(doc.documentElement, "body")[0]!, "tbl")[index - 1]! }
function canonical(node: Node): string {
  if (node.nodeType !== 1) return JSON.stringify([node.nodeType, node.nodeValue])
  const el = node as Element
  return JSON.stringify([el.namespaceURI, el.nodeName, Array.from(el.attributes).map(a => [a.namespaceURI, a.name, a.value]).sort(), Array.from(el.childNodes).map(canonical)])
}
async function author(source: Uint8Array, table = 1) {
  const packet = await docxTableRowDeletionTargets(source), target = packet.targets.find(t => t.table === table && t.row === 2)!
  expect(target).toBeDefined()
  const request: DocxTableRowDeletion = { source_sha256: packet.source_sha256, target_handle: target.target_handle, author: "Current Editor" }
  return { request, authored: await authorDocxTableRowDeletion(source, request) }
}
function unpackFlat(ooxml: string) {
  const doc = parse(ooxml), parts: Record<string, Uint8Array> = {}
  for (const part of kids(doc.documentElement)) {
    const name = part.getAttributeNS(PKG, "name")!.slice(1), data = kids(part)[0]!
    parts[name] = data.localName === "xmlData" ? strToU8(xml(kids(data)[0]!)) : new Uint8Array(Buffer.from(data.textContent ?? "", "base64"))
  }
  return parts
}

test("paired complete-table imports reconstruct the verified full package with mixed runs, geometry and old review intact", async () => {
  const source = fixture({ leadingTable: true }), { request, authored } = await author(source, 2)
  const result = await prepareDocxTableRowDeletionImport(source, authored.bytes, request)
  expect(result.tableIndex).toBe(2)
  expect(result.revisions.map(r => r.family)).toEqual(["row_delete"])
  const before = unpackFlat(result.beforeOoxml), after = unpackFlat(result.ooxml), original = unzipSync(source), fullAuthored = unzipSync(authored.bytes)
  expect(Object.keys(after)).toEqual(Object.keys(before))
  for (const name of Object.keys(before).filter(n => n !== MAIN)) expect(after[name]).toEqual(before[name])
  expect(Object.keys(after).sort()).toEqual(["_rels/.rels", MAIN, "word/_rels/document.xml.rels", "word/styles.xml", "word/fontTable.xml", "word/_rels/fontTable.xml.rels", "word/theme/theme1.xml", "word/numbering.xml", "word/_rels/numbering.xml.rels", "word/fonts/example.odttf", "word/media/bullet.png"].sort())
  expect(after["word/media/bullet.png"]).toEqual(original["word/media/bullet.png"])
  expect(after["word/fonts/example.odttf"]).toEqual(original["word/fonts/example.odttf"])
  expect(canonical(tableAt(parse(before[MAIN]!)))).toBe(canonical(tableAt(parse(original[MAIN]!), 2)))
  expect(canonical(tableAt(parse(after[MAIN]!)))).toBe(canonical(tableAt(parse(fullAuthored[MAIN]!), 2)))
  const fragment = parse(result.beforeTableXml)
  expect(fragment.documentElement.namespaceURI).toBe(W)
  expect(fragment.getElementsByTagNameNS(W, "ins")[0]!.getAttributeNS(W, "author")).toBe("Earlier Editor")
  const doc = parse(original[MAIN]!), previous = tableAt(doc, result.tableIndex)
  previous.parentNode!.replaceChild(doc.importNode(tableAt(parse(after[MAIN]!)), true), previous)
  const rebuilt = zipSync({ ...original, [MAIN]: strToU8(xml(doc)), [SETTINGS]: fullAuthored[SETTINGS]! })
  expect(canonical(parse(unzipSync(rebuilt)[MAIN]!).documentElement)).toBe(canonical(parse(fullAuthored[MAIN]!).documentElement))
  expect((await verifyDocxTableRowDeletion(source, rebuilt, request)).revisions.map(r => r.family)).toEqual(["row_delete"])
  expect(await resolveDocxTableRowDeletion(source, rebuilt, request, "reject")).toEqual(source)
  expect(canonical(parse(unzipSync(await resolveDocxTableRowDeletion(source, rebuilt, request, "accept"))[MAIN]!).documentElement)).toBe(canonical(parse(unzipSync(await resolveDocxTableRowDeletion(source, authored.bytes, request, "accept"))[MAIN]!).documentElement))
  expect(unzipSync(rebuilt)["word/comments.xml"]).toEqual(original["word/comments.xml"])
})

test("only the enclosing table travels: outside comments, section properties and trailing paragraphs never leak", async () => {
  for (const trailing of [true, false]) {
    const source = fixture({ trailing }), { request, authored } = await author(source)
    const result = await prepareDocxTableRowDeletionImport(source, authored.bytes, request)
    for (const ooxml of [result.beforeOoxml, result.ooxml]) {
      const flat = unpackFlat(ooxml), doc = parse(flat[MAIN]!)
      expect(kids(direct(doc.documentElement, "body")[0]!).map(n => n.localName)).toEqual(["tbl"])
      expect(doc.getElementsByTagNameNS(W, "sectPr")).toHaveLength(0)
      expect(doc.getElementsByTagNameNS(W, "commentRangeStart")).toHaveLength(0)
      expect(doc.getElementsByTagNameNS(W, "commentReference")).toHaveLength(0)
      expect(ooxml).not.toContain("Opening outside table")
      expect(ooxml).not.toContain("Ending outside table")
      expect(ooxml).not.toContain("Keep the outside review note")
      expect(flat[SETTINGS]).toBeUndefined()
    }
  }
})

test("annotations in a retained cell or crossing the enclosing boundary refuse even though the deleted row is eligible", async () => {
  for (const crossing of [false, true]) {
    const source = changed(fixture(), doc => {
      const table = tableAt(doc), header = table.getElementsByTagNameNS(W, "p")[0]!, opening = doc.getElementsByTagNameNS(W, "p")[0]!
      const start = doc.createElementNS(W, "w:bookmarkStart"), end = doc.createElementNS(W, "w:bookmarkEnd")
      start.setAttributeNS(W, "w:id", "66"); start.setAttributeNS(W, "w:name", "RetainedReference"); end.setAttributeNS(W, "w:id", "66")
      const first = crossing ? opening : header
      first.insertBefore(start, first.firstChild); header.appendChild(end)
    })
    const { request, authored } = await author(source)
    await expect(prepareDocxTableRowDeletionImport(source, authored.bytes, request)).rejects.toThrow("table_annotation")
  }
})

test("stale sources, changed peers and changed dependencies cannot become table imports", async () => {
  const source = fixture(), { request, authored } = await author(source)
  const peer = changed(authored.bytes, doc => doc.getElementsByTagNameNS(W, "t")[0]!.appendChild(doc.createTextNode(" changed")))
  const style = changed(authored.bytes, doc => doc.documentElement.setAttribute("changed", "true"), "word/styles.xml")
  await expect(prepareDocxTableRowDeletionImport(source, peer, request)).rejects.toThrow("docx_row_deletion:")
  await expect(prepareDocxTableRowDeletionImport(source, style, request)).rejects.toThrow("unrelated_part_changed")
  const stale = changed(source, doc => doc.getElementsByTagNameNS(W, "t")[0]!.appendChild(doc.createTextNode(" changed")))
  await expect(prepareDocxTableRowDeletionImport(stale, authored.bytes, request)).rejects.toThrow("stale_source")
})

test("merged or nested source tables never produce an import payload", async () => {
  const source = fixture(), { request } = await author(source)
  const merged = changed(source, doc => {
    const properties = doc.getElementsByTagNameNS(W, "tcPr")[0]!, span = doc.createElementNS(W, "w:gridSpan")
    span.setAttributeNS(W, "w:val", "2"); properties.appendChild(span)
  })
  const nested = changed(source, doc => doc.getElementsByTagNameNS(W, "tc")[0]!.appendChild(tableAt(doc).cloneNode(true)))
  for (const bytes of [merged, nested]) {
    const input = { ...request, source_sha256: createHash("sha256").update(bytes).digest("hex") }
    await expect(prepareDocxTableRowDeletionImport(bytes, bytes, input)).rejects.toThrow(/target_not_eligible|nested_or_wrapped_table_unsupported/)
  }
})

test("dangling, external and non-formatting dependency relationships never enter the closure", async () => {
  const baseline = fixture(), control = await author(baseline)
  for (const kind of ["missing", "external", "forbidden"]) {
    const corrupt = (doc: Document) => {
      const rel = kids(doc.documentElement)[0]!
      if (kind === "external") { rel.setAttribute("TargetMode", "External"); rel.setAttribute("Target", "https://example.invalid/bullet.png") }
      else if (kind === "forbidden") { rel.setAttribute("Type", `${O}/font`); rel.setAttribute("Target", "comments.xml") }
      else rel.setAttribute("Target", "media/missing.png")
    }
    // Preserve the same malformed dependency on both sides of an otherwise
    // verified deletion, so this exercises import closure rather than SDK save.
    const source = changed(baseline, corrupt, "word/_rels/numbering.xml.rels"), authored = changed(control.authored.bytes, corrupt, "word/_rels/numbering.xml.rels")
    const packet = await docxTableRowDeletionTargets(source), target = packet.targets.find(t => t.table === 1 && t.row === 2)!
    const request = { ...control.request, source_sha256: packet.source_sha256, target_handle: target.target_handle }
    await expect(prepareDocxTableRowDeletionImport(source, authored, request)).rejects.toThrow(kind === "external" ? "external_dependency" : kind === "forbidden" ? "dependency_content_type_unsupported" : "dependency_unavailable")
  }
})

test("table relationship references cannot silently lose unsupported or missing targets", async () => {
  const source = changed(fixture(), doc => {
    const properties = doc.getElementsByTagNameNS(W, "rPr")[0]!
    const font = doc.createElementNS(W, "w:rFonts"); font.setAttributeNS(O, "r:id", "missing"); properties.appendChild(font)
  })
  const { request, authored } = await author(source)
  await expect(prepareDocxTableRowDeletionImport(source, authored.bytes, request)).rejects.toThrow("unsupported_or_dangling_relationship")
})

test("oversized dependency closure is refused before delivery", async () => {
  const parts = unzipSync(fixture()); parts["word/media/bullet.png"] = new Uint8Array(1024 * 1024 + 1)
  const source = zipSync(parts), { request, authored } = await author(source)
  await expect(prepareDocxTableRowDeletionImport(source, authored.bytes, request)).rejects.toThrow("dependency_closure_large")
})
