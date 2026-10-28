import { expect, spyOn, test } from "bun:test"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { docxodus } from "./docxodus"
import { authorDocxTableRowInsertion, docxTableRowInsertionTargets, resolveDocxTableRowInsertion, type DocxTableRowInsertion } from "./docx-table-row-insertion"

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
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
    "word/styles.xml": strToU8(`<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>${kind === "conditional" ? '<w:style w:type="table" w:styleId="Conditional"><w:tblStylePr w:type="firstRow"><w:rPr><w:b/></w:rPr></w:tblStylePr></w:style>' : ""}</w:styles>`),
    "word/settings.xml": strToU8(`<w:settings xmlns:w="${W}"/>`),
    "word/comments.xml": strToU8(`<w:comments xmlns:w="${W}"><w:comment w:id="2" w:author="Commenter"><w:p><w:r><w:t>Keep original comment.</w:t></w:r></w:p></w:comment></w:comments>`),
    "word/document.xml": strToU8(`<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${p("Before")}${body}${p("After")}<w:sectPr/></w:body></w:document>`),
  }, { mtime: new Date(1980, 0, 1) })
}
const parse = (bytes: Uint8Array) => new DOMParser().parseFromString(strFromU8(unzipSync(bytes)["word/document.xml"]!), "application/xml")
function modify(bytes: Uint8Array, change: (doc: Document) => void) {
  const parts = unzipSync(bytes), doc = parse(bytes)
  change(doc)
  parts["word/document.xml"] = strToU8(new XMLSerializer().serializeToString(doc))
  return zipSync(parts)
}
async function input(source: Uint8Array, cells = ["New label", "New obligation"]): Promise<DocxTableRowInsertion> {
  const packet = await docxTableRowInsertionTargets(source)
  return { source_sha256: packet.source_sha256, target_handle: packet.targets[0]!.target_handle, position: "after", cells, author: "Existing Editor" }
}
const direct = (node: Node, name: string) => Array.from(node.childNodes).filter((e): e is Element => e.nodeType === 1 && (e as Element).namespaceURI === W && (e as Element).localName === name)
const rowText = (bytes: Uint8Array) => Array.from(parse(bytes).getElementsByTagNameNS(W, "tr")).map(row => direct(row, "tc").map(cell => Array.from(cell.getElementsByTagNameNS(W, "t")).map(t => t.textContent).join("")))
async function revisions(bytes: Uint8Array) { const s = (await docxodus()).openDocxSession(bytes, { persistAnchorIds: false }); try { return s.listRevisions() } finally { s.close() } }
const durableRecords = (rows: Awaited<ReturnType<typeof revisions>>) => rows.map(({ anchorId: _anchor, affectedAnchors: _affected, ...record }) => record)

test("one reviewed row groups actual carriers and preserves same-author revisions, comments and package parts", async () => {
  const source = fixture(), original = new Uint8Array(source), request = await input(source)
  const authored = await authorDocxTableRowInsertion(source, request)
  expect(authored.revisions.map(r => r.family).sort()).toEqual(["content_insert", "content_insert", "row_insert"])
  expect(rowText(authored.bytes)).toEqual([["Reference", "Ten"], request.cells, ["Prior boldlinkFixed cache", "Right"]])
  const accepted = await resolveDocxTableRowInsertion(source, authored.bytes, request, "accept")
  const rejected = await resolveDocxTableRowInsertion(source, authored.bytes, request, "reject")
  expect(rejected).toEqual(source)
  expect(durableRecords(await revisions(accepted))).toEqual(durableRecords(await revisions(source)))
  expect(rowText(accepted)).toEqual(rowText(authored.bytes))
  const old = unzipSync(source), current = unzipSync(authored.bytes)
  for (const name of Object.keys(old).filter(n => !["word/document.xml", "word/settings.xml"].includes(n))) expect(current[name]).toEqual(old[name])
  expect(unzipSync(accepted)["word/settings.xml"]).toEqual(old["word/settings.xml"])
  expect(source).toEqual(original)
})

test("literal punctuation, emoji and empty cells remain literal words without added dependencies", async () => {
  const source = fixture(), request = await input(source, ["**literal** [x] \\_ <tag> & 😀e\u0301", ""])
  request.position = "before"
  const authored = await authorDocxTableRowInsertion(source, request)
  expect(rowText(authored.bytes)[0]).toEqual(request.cells)
  expect(authored.revisions.filter(r => r.family === "content_insert")).toHaveLength(1)
  expect(await resolveDocxTableRowInsertion(source, authored.bytes, request, "reject")).toEqual(source)
})

test("source identity includes unchanged wording's formatting and exact opaque target", async () => {
  const source = fixture(), request = await input(source)
  const changed = modify(source, doc => { const props = doc.getElementsByTagNameNS(W, "tcPr")[0]!; props.appendChild(doc.createElementNS(W, "w:shd")) })
  await expect(authorDocxTableRowInsertion(changed, request)).rejects.toThrow("stale_source")
  await expect(authorDocxTableRowInsertion(source, { ...request, target_handle: "guessed-row-1" })).rejects.toThrow("target_or_cell_count_changed")
  await expect(authorDocxTableRowInsertion(source, { ...request, cells: ["One"] })).rejects.toThrow("target_or_cell_count_changed")
})

test("supported geometry control is admitted, while merges, nested tables, conditional styles and pending rows are unavailable", async () => {
  expect((await docxTableRowInsertionTargets(fixture())).targets).toHaveLength(2)
  for (const kind of ["merged", "pending", "styled"]) expect((await docxTableRowInsertionTargets(fixture(kind))).targets).toEqual([])
  await expect(docxTableRowInsertionTargets(fixture("conditional"))).rejects.toThrow("conditional_or_numbered_style_unsupported")
  await expect(docxTableRowInsertionTargets(fixture("nested"))).rejects.toThrow("nested_or_wrapped_table_unsupported")
})

test("multi-paragraph reference cells are unavailable instead of advertising joined words", async () => {
  const source = fixture(), original = new Uint8Array(source)
  expect((await docxTableRowInsertionTargets(source)).targets[0]!.cells[0]).toBe("Reference")
  const multiple = modify(source, doc => {
    const cell = doc.getElementsByTagNameNS(W, "tc")[0]!, first = cell.getElementsByTagNameNS(W, "p")[0]!
    first.getElementsByTagNameNS(W, "t")[0]!.textContent = "Pay"
    const next = first.cloneNode(true) as Element
    next.getElementsByTagNameNS(W, "t")[0]!.textContent = "monthly"
    cell.appendChild(next)
  })
  const before = new Uint8Array(multiple)
  expect((await docxTableRowInsertionTargets(multiple)).targets).toEqual([])
  expect(multiple).toEqual(before)
  expect(source).toEqual(original)
})

test("3 by 3 table insertion preserves existing rows at a middle and final boundary", async () => {
  const source = modify(fixture(), doc => {
    const table = doc.getElementsByTagNameNS(W, "tbl")[0]!, grid = direct(table, "tblGrid")[0]!
    grid.appendChild(grid.firstChild!.cloneNode(true))
    for (const [index, row] of direct(table, "tr").entries()) {
      const extra = direct(row, "tc")[0]!.cloneNode(true) as Element
      while (extra.childNodes.length > 1) extra.removeChild(extra.lastChild!)
      const paragraph = doc.createElementNS(W, "w:p"), run = doc.createElementNS(W, "w:r"), text = doc.createElementNS(W, "w:t")
      text.textContent = `Third ${index + 1}`; run.appendChild(text); paragraph.appendChild(run); extra.appendChild(paragraph); row.appendChild(extra)
    }
    const final = direct(table, "tr")[0]!.cloneNode(true) as Element
    for (const [index, cell] of direct(final, "tc").entries()) {
      while (cell.childNodes.length > 1) cell.removeChild(cell.lastChild!)
      const paragraph = doc.createElementNS(W, "w:p"), run = doc.createElementNS(W, "w:r"), text = doc.createElementNS(W, "w:t")
      text.textContent = `Final ${index + 1}`; run.appendChild(text); paragraph.appendChild(run); cell.appendChild(paragraph)
    }
    table.appendChild(final)
  })
  const original = new Uint8Array(source), packet = await docxTableRowInsertionTargets(source), rows = rowText(source)
  expect(packet.targets.map(target => [target.row, target.columns])).toEqual([[1, 3], [2, 3], [3, 3]])
  for (const index of [1, 2]) {
    const request: DocxTableRowInsertion = { source_sha256: packet.source_sha256, target_handle: packet.targets[index]!.target_handle,
      position: "after", cells: ["New one", "New two", "New three"], author: "Existing Editor" }
    const authored = await authorDocxTableRowInsertion(source, request)
    expect(rowText(authored.bytes)).toEqual([...rows.slice(0, index + 1), request.cells, ...rows.slice(index + 1)])
    const accepted = await resolveDocxTableRowInsertion(source, authored.bytes, request, "accept")
    expect(rowText(accepted)).toEqual(rowText(authored.bytes))
    expect(durableRecords(await revisions(accepted))).toEqual(durableRecords(await revisions(source)))
    expect(await resolveDocxTableRowInsertion(source, authored.bytes, request, "reject")).toEqual(source)
  }
  expect(source).toEqual(original)
})

test("identical tables keep distinct exact row targets and preserve the unselected table", async () => {
  const source = modify(fixture(), doc => {
    const table = doc.getElementsByTagNameNS(W, "tbl")[0]!, firstCell = table.getElementsByTagNameNS(W, "tc")[0]!.cloneNode(true)
    for (const row of direct(table, "tr")) for (const cell of direct(row, "tc")) row.replaceChild(firstCell.cloneNode(true), cell)
    table.parentNode!.insertBefore(table.cloneNode(true), table.nextSibling)
  })
  const original = new Uint8Array(source), packet = await docxTableRowInsertionTargets(source)
  const first = packet.targets.find(target => target.table === 1 && target.row === 2)!, second = packet.targets.find(target => target.table === 2 && target.row === 2)!
  expect(packet.targets).toHaveLength(4)
  expect(first.cells).toEqual(second.cells)
  expect(first.target_handle).not.toBe(second.target_handle)
  const request: DocxTableRowInsertion = { source_sha256: packet.source_sha256, target_handle: second.target_handle,
    position: "after", cells: ["Chosen table", "New obligation"], author: "Existing Editor" }
  const authored = await authorDocxTableRowInsertion(source, request), current = parse(authored.bytes)
  const tables = Array.from(current.getElementsByTagNameNS(W, "tbl"))
  expect(tables.map(table => direct(table, "tr").length)).toEqual([2, 3])
  const serialize = (node: Node) => new XMLSerializer().serializeToString(node)
  expect(serialize(tables[0]!)).toEqual(serialize(parse(source).getElementsByTagNameNS(W, "tbl")[0]!))
  expect(rowText(authored.bytes).at(-1)).toEqual(request.cells)
  await expect(resolveDocxTableRowInsertion(source, authored.bytes, { ...request, target_handle: first.target_handle }, "accept")).rejects.toThrow("inserted_row_count_changed")
  expect(await resolveDocxTableRowInsertion(source, authored.bytes, request, "reject")).toEqual(source)
  expect(source).toEqual(original)
})

test("independent resolution refuses carrier collisions, semantic extra nodes and changed unrelated parts", async () => {
  const source = fixture(), request = await input(source), authored = await authorDocxTableRowInsertion(source, request)
  const collision = modify(authored.bytes, doc => { const row = doc.getElementsByTagNameNS(W, "tr")[1]!; row.getElementsByTagNameNS(W, "ins")[0]!.setAttributeNS(W, "w:id", "101") })
  await expect(resolveDocxTableRowInsertion(source, collision, request, "reject")).rejects.toThrow("source_revision_id_collision")
  const extra = modify(authored.bytes, doc => { const p = doc.getElementsByTagNameNS(W, "tr")[1]!.getElementsByTagNameNS(W, "p")[0]!; p.appendChild(doc.createElementNS(W, "w:proofErr")) })
  await expect(resolveDocxTableRowInsertion(source, extra, request, "accept")).rejects.toThrow("unsupported_new_row_content")
  const parts = unzipSync(authored.bytes)
  parts["word/comments.xml"] = strToU8(strFromU8(parts["word/comments.xml"]!).replace("original", "changed"))
  await expect(resolveDocxTableRowInsertion(source, zipSync(parts), request, "accept")).rejects.toThrow("unrelated_part_changed")
  const moved = modify(authored.bytes, doc => {
    const old = Array.from(doc.getElementsByTagNameNS(W, "ins")).find(el => el.getAttributeNS(W, "id") === "101")!
    doc.getElementsByTagNameNS(W, "tr")[2]!.getElementsByTagNameNS(W, "p")[1]!.appendChild(old)
  })
  await expect(resolveDocxTableRowInsertion(source, moved, request, "accept")).rejects.toThrow("original_document_changed")
  expect(await resolveDocxTableRowInsertion(source, authored.bytes, request, "reject")).toEqual(source)
})

test("missing or unsupported manifest digest refuses authoring before any insertion", async () => {
  const source = fixture(), request = await input(source), s = (await docxodus()).openDocxSession(source, {})
  const prototype = Object.getPrototypeOf(s) as typeof s, baseline = s.getPackageManifest(); s.close()
  const manifest = spyOn(prototype, "getPackageManifest"), insert = spyOn(prototype, "insertTableRow")
  try {
    for (const rawPackageBytesDigest of [null, { algorithm: "SHA-1", value: "a".repeat(64) }, { algorithm: "SHA-256", value: "" }]) {
      manifest.mockReturnValue({ ...baseline, rawPackageBytesDigest } as typeof baseline)
      await expect(authorDocxTableRowInsertion(source, request)).rejects.toThrow("package_digest_unavailable")
    }
    expect(insert).not.toHaveBeenCalled()
  } finally { manifest.mockRestore(); insert.mockRestore() }
})

test("a later fill failure rolls back the real SDK batch package and version", async () => {
  const source = fixture(), request = await input(source), dx = await docxodus(), s = dx.openDocxSession(source, {})
  const prototype = Object.getPrototypeOf(s) as typeof s; s.close()
  const fill = prototype.replaceCellContent, batch = prototype.executeBatch
  let calls = 0, rolledBack = false
  const fail = spyOn(prototype, "replaceCellContent").mockImplementation(function(this: typeof s, ...args: Parameters<typeof fill>) {
    if (++calls === 2) return { success: false, created: [], removed: [], modified: [], error: { code: "internal_error", message: "Synthetic later fill failure" } }
    return fill.apply(this, args)
  })
  const observe = spyOn(prototype, "executeBatch").mockImplementation(function(this: typeof s, ...args: Parameters<typeof batch>) {
    const before = this.getPackageManifest().rawPackageBytesDigest, version = this.getVersion(), result = batch.apply(this, args)
    const after = this.getPackageManifest().rawPackageBytesDigest
    rolledBack = !result.success && after.algorithm === before.algorithm && after.value === before.value && this.getVersion() === version
    return result
  })
  try { await expect(authorDocxTableRowInsertion(source, request)).rejects.toThrow("authoring_failed_rolled_back"); expect(calls).toBe(2); expect(rolledBack).toBe(true) }
  finally { fail.mockRestore(); observe.mockRestore() }
})

test("a grouped decision failure rolls back earlier successful content decisions", async () => {
  const source = fixture(), request = await input(source), authored = await authorDocxTableRowInsertion(source, request), dx = await docxodus(), s = dx.openDocxSession(source, {})
  const prototype = Object.getPrototypeOf(s) as typeof s; s.close()
  const reject = prototype.rejectRevision, batch = prototype.executeBatch
  let calls = 0, rolledBack = false
  const fail = spyOn(prototype, "rejectRevision").mockImplementation(function(this: typeof s, ...args: Parameters<typeof reject>) {
    if (++calls === 3) return { success: false, created: [], removed: [], modified: [], error: { code: "revision_not_found", message: "Synthetic disappeared group member" } }
    return reject.apply(this, args)
  })
  const observe = spyOn(prototype, "executeBatch").mockImplementation(function(this: typeof s, ...args: Parameters<typeof batch>) {
    const before = this.getPackageManifest().rawPackageBytesDigest, version = this.getVersion(), result = batch.apply(this, args)
    const after = this.getPackageManifest().rawPackageBytesDigest
    rolledBack = !result.success && after.algorithm === before.algorithm && after.value === before.value && this.getVersion() === version
    return result
  })
  try { await expect(resolveDocxTableRowInsertion(source, authored.bytes, request, "reject")).rejects.toThrow("resolution_failed_rolled_back"); expect(calls).toBe(3); expect(rolledBack).toBe(true) }
  finally { fail.mockRestore(); observe.mockRestore() }
  expect(await resolveDocxTableRowInsertion(source, authored.bytes, request, "reject")).toEqual(source)
})
