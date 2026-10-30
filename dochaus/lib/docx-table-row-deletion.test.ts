import { expect, spyOn, test } from "bun:test"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { docxodus } from "./docxodus"
import { docxTableRowInsertionTargets } from "./docx-table-row-insertion"
import { authorDocxTableRowDeletion, docxTableRowDeletionTargets, resolveDocxTableRowDeletion, type DocxTableRowDeletion } from "./docx-table-row-deletion"

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", O = "http://schemas.openxmlformats.org/officeDocument/2006/relationships", R = "http://schemas.openxmlformats.org/package/2006/relationships"
const BODY = "word/document.xml", AUTHOR = "Review Editor", stamp = `w:author="${AUTHOR}" w:date="2026-01-01T00:00:00Z"`
const escape = (value: string) => value.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;")
const p = (value: string) => `<w:p><w:r><w:t xml:space="preserve">${escape(value)}</w:t></w:r></w:p>`
const cell = (content: string) => `<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>${content}</w:tc>`
const row = (a: string, b: string, props = "") => `<w:tr>${props ? `<w:trPr>${props}</w:trPr>` : ""}${cell(a)}${cell(b)}</w:tr>`
function fixture(existingMinimumHeight = false) {
  const old = `<w:p><w:commentRangeStart w:id="4"/><w:del w:id="100" ${stamp}><w:r><w:delText>Former</w:delText></w:r></w:del><w:ins w:id="101" ${stamp}><w:r><w:t>Current</w:t></w:r></w:ins><w:commentRangeEnd w:id="4"/><w:r><w:commentReference w:id="4"/></w:r></w:p>`
  const formatted = `<w:p><w:pPr><w:spacing w:after="120"/></w:pPr><w:r><w:t xml:space="preserve"> Keep </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r><w:r><w:rPr><w:i/><w:color w:val="123456"/></w:rPr><w:t xml:space="preserve"> **literal** 😀</w:t></w:r></w:p>`
  return zipSync({
    "[Content_Types].xml": strToU8('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/><Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>'),
    "_rels/.rels": strToU8(`<Relationships xmlns="${R}"><Relationship Id="main" Type="${O}/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/_rels/document.xml.rels": strToU8(`<Relationships xmlns="${R}"><Relationship Id="styles" Type="${O}/styles" Target="styles.xml"/><Relationship Id="settings" Type="${O}/settings" Target="settings.xml"/><Relationship Id="comments" Type="${O}/comments" Target="comments.xml"/></Relationships>`),
    "word/styles.xml": strToU8(`<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`),
    "word/settings.xml": strToU8(`<w:settings xmlns:w="${W}"/>`),
    "word/comments.xml": strToU8(`<w:comments xmlns:w="${W}"><w:comment w:id="4" w:author="Previous Reviewer"><w:p><w:r><w:t>Retain this exact review note.</w:t></w:r></w:p></w:comment></w:comments>`),
    [BODY]: strToU8(`<w:document xmlns:w="${W}" xmlns:r="${O}"><w:body>${p("Before")}${`<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>${row(old, p("Heading"), "<w:tblHeader/>")}${row(formatted, p(""), existingMinimumHeight ? '<w:trHeight w:val="480" w:hRule="atLeast"/>' : "")}${row(p("Last row"), p("Stable peer"))}</w:tbl>`}<w:p><w:ins w:id="102" ${stamp}><w:r><w:t>After the table</w:t></w:r></w:ins></w:p><w:sectPr/></w:body></w:document>`),
  }, { mtime: new Date(1980, 0, 1) })
}
const parse = (bytes: Uint8Array) => new DOMParser().parseFromString(strFromU8(bytes), "application/xml")
const serialize = (node: Node) => new XMLSerializer().serializeToString(node)
function changed(bytes: Uint8Array, action: (doc: Document) => void, part = BODY) {
  const parts = unzipSync(bytes), doc = parse(parts[part]!)
  action(doc); parts[part] = strToU8(serialize(doc)); return zipSync(parts)
}
const body = (bytes: Uint8Array) => parse(unzipSync(bytes)[BODY]!)
const direct = (node: Node, name: string) => Array.from(node.childNodes).filter((n): n is Element => n.nodeType === 1 && (n as Element).namespaceURI === W && (n as Element).localName === name)
const rows = (bytes: Uint8Array) => Array.from(body(bytes).getElementsByTagNameNS(W, "tr"))
function expectedBody(source: Uint8Array, index: number) { const doc = body(source), row = doc.getElementsByTagNameNS(W, "tr")[index]!; row.parentNode!.removeChild(row); return doc.documentElement!.toString() }
async function input(source: Uint8Array, index = 1): Promise<DocxTableRowDeletion> {
  const packet = await docxTableRowInsertionTargets(source)
  return { source_sha256: packet.source_sha256, target_handle: packet.targets[index]!.target_handle, author: AUTHOR }
}
async function revisions(bytes: Uint8Array) {
  const s = (await docxodus()).openDocxSession(bytes, { persistAnchorIds: false })
  try { return s.listRevisions().map(({ anchorId: _anchor, affectedAnchors: _affected, ...record }) => record) } finally { s.close() }
}

test("middle and last-of-several deletion retain source words until selective accept, and reject restores every source byte", async () => {
  for (const { index, minimumHeight } of [{ index: 1, minimumHeight: false }, { index: 2, minimumHeight: false }, { index: 1, minimumHeight: true }]) {
    const source = fixture(minimumHeight), original = new Uint8Array(source), request = await input(source, index), authored = await authorDocxTableRowDeletion(source, request)
    expect(authored.revisions.map(r => r.family)).toEqual(["row_delete"])
    expect(rows(authored.bytes)).toHaveLength(3)
    const deleted = rows(authored.bytes)[index]!
    expect(deleted.getElementsByTagNameNS(W, "delText").length).toBeGreaterThan(0)
    expect(deleted.getElementsByTagNameNS(W, "t").length).toBe(0)
    const accepted = await resolveDocxTableRowDeletion(source, authored.bytes, request, "accept")
    expect(body(accepted).documentElement!.toString()).toBe(expectedBody(source, index))
    expect(rows(accepted)).toHaveLength(2)
    expect(await revisions(accepted)).toEqual(await revisions(source))
    const old = unzipSync(source), output = unzipSync(accepted)
    expect(Object.keys(output).sort()).toEqual(Object.keys(old).sort())
    for (const name of Object.keys(old).filter(n => n !== BODY)) expect(output[name]).toEqual(old[name])
    expect(await resolveDocxTableRowDeletion(source, authored.bytes, request, "reject")).toEqual(source)
    expect(source).toEqual(original)
  }
})

test("exact row handles choose the second identical table and cannot rebase after an unrelated formatting change", async () => {
  const source = changed(fixture(), doc => { const table = doc.getElementsByTagNameNS(W, "tbl")[0]!; table.parentNode!.insertBefore(table.cloneNode(true), table.nextSibling) })
  // Give the second table no copied revision/comment markers: identical data
  // rows still exercise exact table identity without malformed duplicate IDs.
  const clean = changed(source, doc => {
    const first = doc.getElementsByTagNameNS(W, "tr")[3]!, cells = direct(first, "tc")
    cells[0]!.replaceChild(doc.importNode(parse(strToU8(p("Current"))).documentElement!, true), direct(cells[0]!, "p")[0]!)
  })
  const request = await input(clean, 4), authored = await authorDocxTableRowDeletion(clean, request), accepted = await resolveDocxTableRowDeletion(clean, authored.bytes, request, "accept")
  expect(Array.from(body(accepted).getElementsByTagNameNS(W, "tbl")).map(t => direct(t, "tr").length)).toEqual([3, 2])
  const modified = changed(clean, doc => doc.getElementsByTagNameNS(W, "spacing")[0]!.setAttributeNS(W, "w:after", "180"))
  await expect(authorDocxTableRowDeletion(modified, request)).rejects.toThrow("stale_source")
  await expect(resolveDocxTableRowDeletion(modified, authored.bytes, request, "accept")).rejects.toThrow("stale_source")
})

test("headers, sole rows, target annotations and crossing zero-text ranges refuse before any deletion", async () => {
  const source = fixture(), examples: Array<{ bytes: Uint8Array; index: number }> = [{ bytes: source, index: 0 }]
  examples.push({ bytes: changed(source, doc => {
    const target = doc.getElementsByTagNameNS(W, "tr")[1]!
    target.insertBefore(doc.createElementNS(W, "w:trPr"), target.firstChild)
  }), index: 1 })
  examples.push({ bytes: changed(source, doc => { const table = doc.getElementsByTagNameNS(W, "tbl")[0]!; direct(table, "tr").filter((_, index) => index !== 1).forEach(row => table.removeChild(row)) }), index: 0 })
  for (const family of ["bookmark", "commentRange", "perm"]) examples.push({ bytes: changed(source, doc => {
    const first = doc.getElementsByTagNameNS(W, "p")[0]!, last = doc.getElementsByTagNameNS(W, "p").item(doc.getElementsByTagNameNS(W, "p").length - 1)!
    const start = doc.createElementNS(W, `w:${family}Start`), end = doc.createElementNS(W, `w:${family}End`)
    start.setAttributeNS(W, "w:id", "50"); end.setAttributeNS(W, "w:id", "50")
    if (family === "bookmark") start.setAttributeNS(W, "w:name", "CrossingRange")
    first.insertBefore(start, first.firstChild); last.appendChild(end)
  }), index: 1 })
  examples.push({ bytes: changed(source, doc => {
    const target = direct(doc.getElementsByTagNameNS(W, "tr")[1]!, "tc")[0]!, paragraph = direct(target, "p")[0]!
    const start = doc.createElementNS(W, "w:commentRangeStart"), end = doc.createElementNS(W, "w:commentRangeEnd")
    start.setAttributeNS(W, "w:id", "50"); end.setAttributeNS(W, "w:id", "50"); paragraph.insertBefore(start, paragraph.firstChild); paragraph.appendChild(end)
  }), index: 1 })
  const sdk = await docxodus(), temp = sdk.openDocxSession(source, {}), prototype = Object.getPrototypeOf(temp) as typeof temp; temp.close()
  const deletion = spyOn(prototype, "deleteTableRow")
  try {
    for (const example of examples) {
      const request = await input(example.bytes, example.index)
      await expect(authorDocxTableRowDeletion(example.bytes, request)).rejects.toThrow("docx_row_deletion:")
    }
    expect(deletion).not.toHaveBeenCalled()
  } finally { deletion.mockRestore() }
})

test("extra, moved or foreign carriers, altered old comments and changed peer content never pass final-artifact proof", async () => {
  const source = fixture(), request = await input(source), authored = await authorDocxTableRowDeletion(source, request)
  const peer = changed(authored.bytes, doc => doc.getElementsByTagNameNS(W, "t")[0]!.appendChild(doc.createTextNode(" changed")))
  const foreign = changed(authored.bytes, doc => {
    const mark = doc.getElementsByTagNameNS(W, "tr")[1]!.getElementsByTagNameNS(W, "del")[0]!; mark.setAttributeNS(W, "w:author", "Another Reviewer")
  })
  const moved = changed(authored.bytes, doc => {
    const paragraph = direct(direct(doc.getElementsByTagNameNS(W, "tr")[1]!, "tc")[0]!, "p")[0]!, wrapper = direct(paragraph, "del")[0]!
    doc.getElementsByTagNameNS(W, "p")[0]!.appendChild(wrapper)
  })
  const duplicate = changed(authored.bytes, doc => {
    const row = doc.getElementsByTagNameNS(W, "tr")[1]!, mark = direct(direct(row, "trPr")[0]!, "del")[0]!
    mark.parentNode!.appendChild(mark.cloneNode(true))
  })
  const comment = changed(authored.bytes, doc => doc.getElementsByTagNameNS(W, "t")[0]!.appendChild(doc.createTextNode(" changed")), "word/comments.xml")
  for (const bytes of [peer, foreign, moved, duplicate, comment]) await expect(resolveDocxTableRowDeletion(source, bytes, request, "accept")).rejects.toThrow("docx_row_deletion:")
  expect(await resolveDocxTableRowDeletion(source, authored.bytes, request, "reject")).toEqual(source)
})

test("the actual SDK mutation and selective decision both roll back when a callback fails after mutating", async () => {
  const source = fixture(), request = await input(source), authored = await authorDocxTableRowDeletion(source, request), sdk = await docxodus(), temp = sdk.openDocxSession(source, {})
  const prototype = Object.getPrototypeOf(temp) as typeof temp; temp.close()
  for (const operation of ["deleteTableRow", "acceptRevision"] as const) {
    const original = prototype[operation], batch = prototype.executeBatch
    let mutated = false, restored = false
    const mutation = spyOn(prototype, operation).mockImplementation(function(this: typeof temp, id: string) {
      const result = original.call(this, id)
      mutated = result.success
      return { ...result, success: false, error: { code: "internal_error", message: "Synthetic failure after native mutation" } }
    })
    const observation = spyOn(prototype, "executeBatch").mockImplementation(function(this: typeof temp, ...args: Parameters<typeof batch>) {
      const before = this.getPackageManifest().rawPackageBytesDigest, version = this.getVersion(), result = batch.apply(this, args), after = this.getPackageManifest().rawPackageBytesDigest
      restored = !result.success && before.algorithm === after.algorithm && before.value === after.value && version === this.getVersion()
      return result
    })
    try {
      await expect(operation === "deleteTableRow" ? authorDocxTableRowDeletion(source, request) : resolveDocxTableRowDeletion(source, authored.bytes, request, "accept")).rejects.toThrow("failed_rolled_back")
      expect(mutated).toBe(true); expect(restored).toBe(true)
    } finally { mutation.mockRestore(); observation.mockRestore() }
  }
})

test("Word row and paragraph identity metadata survives tracked deletion and both decisions", async () => {
  const W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
  const source = changed(fixture(true), doc => {
    doc.documentElement!.setAttributeNS("http://www.w3.org/2000/xmlns/", "xmlns:w14", W14)
    const target = doc.getElementsByTagNameNS(W, "tr")[1]!
    target.setAttributeNS(W14, "w14:paraId", "1234ABCD")
    target.setAttributeNS(W14, "w14:textId", "00ABCDEF")
    const paragraph = direct(direct(target, "tc")[0]!, "p")[0]!
    paragraph.setAttributeNS(W14, "w14:paraId", "ABCDEF12")
    paragraph.setAttributeNS(W14, "w14:textId", "12345678")
    const run = doc.createElementNS(W, "w:r"), properties = doc.createElementNS(W, "w:rPr"), underline = doc.createElementNS(W, "w:u"), text = doc.createElementNS(W, "w:t")
    underline.setAttributeNS(W, "w:val", "single"); properties.appendChild(underline); run.appendChild(properties); text.appendChild(doc.createTextNode(" underline")); run.appendChild(text); paragraph.appendChild(run)
  })
  const target = rows(source)[1]!
  // The previous row guard rejected any non-w attribute, including these IDs.
  expect(Array.from(target.attributes).every(a => a.namespaceURI === W && /^rsid(?:R|RPr|Del|Tr)$/.test(a.localName))).toBe(false)
  const request = await input(source)
  expect((await docxTableRowDeletionTargets(source)).targets.some(t => t.target_handle === request.target_handle)).toBe(true)
  const authored = await authorDocxTableRowDeletion(source, request), tracked = rows(authored.bytes)[1]!
  for (const name of ["paraId", "textId"]) expect(tracked.getAttributeNS(W14, name)).toBe(target.getAttributeNS(W14, name))
  expect(tracked.getElementsByTagNameNS(W, "b").length).toBe(1)
  expect(tracked.getElementsByTagNameNS(W, "i").length).toBe(1)
  expect(tracked.getElementsByTagNameNS(W, "u").length).toBe(1)
  const altered = changed(authored.bytes, doc => doc.getElementsByTagNameNS(W, "tr")[1]!.setAttributeNS(W14, "w14:paraId", "FFFFFFFF"))
  await expect(resolveDocxTableRowDeletion(source, altered, request, "accept")).rejects.toThrow("tracked_inverse_changed_source_row")
  expect(body(await resolveDocxTableRowDeletion(source, authored.bytes, request, "accept")).documentElement!.toString()).toBe(expectedBody(source, 1))
  expect(await resolveDocxTableRowDeletion(source, authored.bytes, request, "reject")).toEqual(source)
  for (const [name, value] of [["paraId", "bad"], ["unknown", "12345678"]]) {
    const invalid = changed(source, doc => doc.getElementsByTagNameNS(W, "tr")[1]!.setAttributeNS(W14, `w14:${name}`, value!))
    const invalidRequest = await input(invalid)
    expect((await docxTableRowDeletionTargets(invalid)).targets.some(t => t.target_handle === invalidRequest.target_handle)).toBe(false)
  }
})
