import { expect, test } from "bun:test"
import { Document, Packer, Paragraph } from "docx"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { DOMParser } from "@xmldom/xmldom"
import { linkedWordDocx, readWordLink as boundedReadWordLink, stripWordLink, wordLastModifiedBy, WordPackageError } from "../../../dochaus/lib/word-linked-docx"
import { WORD_ADDIN_VERSION, WORD_LINK_SETTING, type WordDocumentLink } from "../../../dochaus/lib/word-addin-contract"

const link: WordDocumentLink = { version: 1, origin: "http://localhost:5190", apiOrigin: "https://localhost:4612", matter: "synthetic-matter", matterName: "Synthetic matter",
  document: "synthetic-id", documentName: "Advice.docx", baseVersion: "synthetic-base", token: "synthetic-capability" }
const plain = () => Packer.toBuffer(new Document({ lastModifiedBy: "Synthetic reviewer", sections: [{ children: [new Paragraph("Neutral sample advice.")] }] }))
const xmls = (bytes: Uint8Array) => Object.values(unzipSync(bytes)).map(bytes => strFromU8(bytes)).join("\n")

export function readWordLink(bytes: Uint8Array) {
  return boundedReadWordLink(bytes) as WordDocumentLink
}

test("reading a handoff rejects ambiguous owned extensions and duplicate link properties", async () => {
  const bytes = await plain()
  expect(boundedReadWordLink(bytes)).toBeNull()
  const parts = unzipSync(linkedWordDocx(bytes, link)), extension = parts["word/webextensions/dochaus.xml"]!
  expect(() => boundedReadWordLink(zipSync({ ...parts, "word/webextensions/duplicate.xml": extension }))).toThrow("ambiguous")
  const duplicate = strFromU8(extension).replace("</we:properties>", `<we:property name="${WORD_LINK_SETTING}" value="{}"/></we:properties>`)
  expect(() => boundedReadWordLink(zipSync({ ...parts, "word/webextensions/dochaus.xml": strToU8(duplicate) }))).toThrow("one doc.haus link")
})

test("linking changes only extension/package parts; stripping removes the capability and preserves unrelated add-ins", async () => {
  const original = unzipSync(await plain())
  const REL = "http://schemas.openxmlformats.org/package/2006/relationships"
  original["_rels/.rels"] = strToU8(strFromU8(original["_rels/.rels"]).replace("</Relationships>",
    `<Relationship Id="emptyPane" Type="http://schemas.microsoft.com/office/2011/relationships/webextensiontaskpanes" Target="NULL"/><Relationship Id="otherPane" Type="http://schemas.microsoft.com/office/2011/relationships/webextensiontaskpanes" Target="word/webextensions/taskpanes.xml"/></Relationships>`))
  original["word/webextensions/taskpanes.xml"] = strToU8(`<wetp:taskpanes xmlns:wetp="http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11"><wetp:taskpane dockstate="left" visibility="1" width="200" row="0"><wetp:webextensionref xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="other"/></wetp:taskpane></wetp:taskpanes>`)
  original["word/webextensions/_rels/taskpanes.xml.rels"] = strToU8(`<Relationships xmlns="${REL}"><Relationship Id="other" Type="http://schemas.microsoft.com/office/2011/relationships/webextension" Target="other.xml"/></Relationships>`)
  original["word/webextensions/other.xml"] = strToU8(`<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" id="other"><we:reference id="other" version="1" store="developer" storeType="Registry"/></we:webextension>`)
  const linked = linkedWordDocx(zipSync(original), link)
  expect(readWordLink(linked)).toEqual(link)
  const linkedParts = unzipSync(linked)
  const extension = new DOMParser().parseFromString(strFromU8(linkedParts["word/webextensions/dochaus.xml"]), "application/xml")
  const reference = extension.getElementsByTagNameNS("http://schemas.microsoft.com/office/webextensions/webextension/2010/11", "reference")[0]
  expect(reference?.getAttribute("version")).toBe(WORD_ADDIN_VERSION)
  // Word starts the manifest's shared runtime when the file opens, so the pane can show itself.
  const background = extension.getElementsByTagNameNS("http://schemas.microsoft.com/office/webextensions/webextension/2010/11", "backgroundApp")[0]
  expect([background?.getAttribute("state"), background?.getAttribute("runtimeId")]).toEqual(["1", "Pane.Url"])
  for (const [name, bytes] of Object.entries(original)) if (!name.endsWith("/") && !["_rels/.rels", "[Content_Types].xml", "word/webextensions/taskpanes.xml", "word/webextensions/_rels/taskpanes.xml.rels"].includes(name)) expect(linkedParts[name]).toEqual(bytes)
  const clean = stripWordLink(linked)
  expect(xmls(clean)).not.toContain(link.token)
  expect(xmls(clean)).not.toContain(WORD_LINK_SETTING)
  const cleanParts = unzipSync(clean)
  expect(cleanParts["word/document.xml"]).toEqual(original["word/document.xml"])
  expect(cleanParts["word/webextensions/other.xml"]).toEqual(original["word/webextensions/other.xml"])
  expect(strFromU8(cleanParts["word/webextensions/taskpanes.xml"])).toContain('r:id="other"')
  expect(strFromU8(cleanParts["word/webextensions/_rels/taskpanes.xml.rels"])).toContain('Target="other.xml"')
  expect(stripWordLink(linked)).toEqual(clean)
  expect(stripWordLink(clean)).toEqual(clean)
  expect(wordLastModifiedBy(linked)).toBe("Synthetic reviewer")
})

test("a relink replaces only doc.haus linkage and ordinary clean bytes are unchanged", async () => {
  const bytes = await plain()
  expect(stripWordLink(bytes)).toEqual(bytes)
  const linked = linkedWordDocx(linkedWordDocx(bytes, link), { ...link, token: "new-synthetic-capability" })
  expect(readWordLink(linked).token).toBe("new-synthetic-capability")
  expect(xmls(linked)).not.toContain('value="synthetic-capability"')
  expect(Object.keys(unzipSync(stripWordLink(linked))).filter(name => name.includes("webextension"))).toEqual([])
})

test("relinking repairs only an internal orphaned NULL taskpane relationship", async () => {
  const type = "http://schemas.microsoft.com/office/2011/relationships/webextensiontaskpanes"
  const namespace = "http://schemas.openxmlformats.org/package/2006/relationships"
  for (const mode of ["", ' TargetMode="Internal"']) {
    const original = unzipSync(await plain())
    original["_rels/.rels"] = strToU8(strFromU8(original["_rels/.rels"]).replace("</Relationships>",
      `<Relationship Id="orphanedPane" Type="${type}" Target="NULL"${mode}/></Relationships>`))
    const input = zipSync(original)
    expect(stripWordLink(input)).toEqual(input)
    const linked = linkedWordDocx(input, link)
    const parts = unzipSync(linked)
    const root = new DOMParser().parseFromString(strFromU8(parts["_rels/.rels"]), "application/xml")
    const panes = Array.from(root.getElementsByTagNameNS(namespace, "Relationship")).filter(node => node.getAttribute("Type") === type)
    expect(panes.map(node => node.getAttribute("Target"))).toEqual(["word/webextensions/taskpanes.xml"])
    expect(readWordLink(linked)).toEqual(link)
    for (const [name, bytes] of Object.entries(original))
      if (!name.endsWith("/") && !["_rels/.rels", "[Content_Types].xml"].includes(name)) expect(parts[name]).toEqual(bytes)
    const clean = stripWordLink(linked)
    expect(xmls(clean)).not.toContain(link.token)
    expect(xmls(clean)).not.toContain(type)
    expect(Object.keys(unzipSync(clean)).filter(name => name.includes("webextension"))).toEqual([])
    expect(stripWordLink(clean)).toEqual(clean)
    const fresh = { ...link, token: "new-synthetic-capability" }
    const relinked = linkedWordDocx(clean, fresh)
    expect(readWordLink(relinked)).toEqual(fresh)
    expect(unzipSync(relinked)["word/document.xml"]).toEqual(original["word/document.xml"])
  }
})

test("relinking still rejects live, external and other unsupported taskpane targets", async () => {
  const type = "http://schemas.microsoft.com/office/2011/relationships/webextensiontaskpanes"
  const taskpanes = strToU8('<wetp:taskpanes xmlns:wetp="http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11"/>')
  const cases: { target: string; mode: string; parts: Record<string, Uint8Array> }[] = [
    { target: "NULL", mode: "", parts: { NULL: taskpanes } },
    { target: "NULL", mode: ' TargetMode="External"', parts: {} },
    { target: "NULL", mode: ' TargetMode=""', parts: {} },
    { target: "null", mode: "", parts: {} },
    { target: "word/webextensions/missing.xml", mode: "", parts: {} },
    { target: "custom/taskpanes.xml", mode: "", parts: { "custom/taskpanes.xml": taskpanes } },
  ]
  for (const row of cases) {
    const original = { ...unzipSync(await plain()), ...row.parts }
    original["_rels/.rels"] = strToU8(strFromU8(original["_rels/.rels"]).replace("</Relationships>",
      `<Relationship Id="unsupportedPane" Type="${type}" Target="${row.target}"${row.mode}/></Relationships>`))
    const input = zipSync(original)
    expect(() => linkedWordDocx(input, link)).toThrow("unsupported taskpane location")
    expect(stripWordLink(input)).toEqual(input)
  }
})

test("unsafe paths, malformed XML, entity declarations and inflated package sizes fail before writing", async () => {
  const original = unzipSync(await plain())
  const invalid: Record<string, Uint8Array>[] = [{ "../outside.xml": strToU8("<outside/>") }, { "word/document.xml": strToU8("<w:document>") },
    { "docProps/core.xml": strToU8('<!DOCTYPE x [<!ENTITY secret "unsafe">]><x>&secret;</x>') }]
  for (const changes of invalid) {
    expect(() => stripWordLink(zipSync({ ...original, ...changes }))).toThrow(WordPackageError)
  }
  const archive = zipSync(original)
  const view = new DataView(archive.buffer, archive.byteOffset, archive.byteLength)
  for (let at = 0; at < archive.length - 46; at++) if (view.getUint32(at, true) === 0x02014b50 && archive[at + 46 + view.getUint16(at + 28, true) - 1] !== 47) {
    view.setUint32(at + 24, 300 * 1024 * 1024, true); break
  }
  expect(() => stripWordLink(archive)).toThrow("package size limit")
})
