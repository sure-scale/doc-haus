import { createHash } from "node:crypto"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { z } from "zod"
import { docxodus } from "./docxodus"
import { docxTableRowInsertionTargets } from "./docx-table-row-insertion"

// Offline proof only. No route, proposal recorder or Word import uses deletion.
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", XML = "http://www.w3.org/XML/1998/namespace"
const BODY = "word/document.xml", SETTINGS = "word/settings.xml"
type Session = ReturnType<Awaited<ReturnType<typeof docxodus>>["openDocxSession"]>
type Revision = ReturnType<Session["listRevisions"]>[number]
type Target = Awaited<ReturnType<typeof docxTableRowInsertionTargets>>["targets"][number]
type Package = ReturnType<typeof unpack>
const requestSchema = z.object({ source_sha256: z.string().regex(/^[a-f0-9]{64}$/), target_handle: z.string().min(1), author: z.string().trim().min(1).max(100) }).strict()
export type DocxTableRowDeletion = z.infer<typeof requestSchema>
const hash = (value: Uint8Array | string) => createHash("sha256").update(value).digest("hex")
function require(value: unknown, reason: string): asserts value { if (!value) throw new Error(`docx_row_deletion:${reason}`) }
const all = (node: Document | Element, name: string) => Array.from(node.getElementsByTagNameNS(W, name))
const kids = (node: Node) => Array.from(node.childNodes).filter((n): n is Element => n.nodeType === 1)
const children = (node: Node, name: string) => kids(node).filter(n => n.namespaceURI === W && n.localName === name)
const xml = (node: Node) => new XMLSerializer().serializeToString(node)
function parse(value: string) {
  require(!/<!DOCTYPE|<!ENTITY/i.test(value), "unsupported_xml")
  const fail = () => { throw new Error("docx_row_deletion:invalid_xml") }
  return new DOMParser({ errorHandler: { warning: fail, error: fail, fatalError: fail } }).parseFromString(value, "application/xml")
}
function unpack(bytes: Uint8Array) {
  const parts = unzipSync(bytes)
  require(parts[BODY] && parts[SETTINGS], "existing_main_and_settings_required")
  const docs = Object.fromEntries(Object.entries(parts).filter(([name]) => /\.(xml|rels)$/.test(name)).map(([name, value]) => [name, parse(strFromU8(value))]))
  return { parts, docs, body: docs[BODY]!, settings: docs[SETTINGS]! }
}
function canonical(node: Node): string {
  if (node.nodeType === 9) return JSON.stringify(Array.from(node.childNodes).filter(n => !(n.nodeType === 7 && n.nodeName === "xml")).map(canonical))
  if (node.nodeType !== 1) return JSON.stringify([node.nodeType, node.nodeName, node.nodeValue])
  const el = node as Element
  return JSON.stringify([el.namespaceURI, el.nodeName, Array.from(el.attributes).map(a => [a.namespaceURI, a.name, a.value]).sort(), Array.from(el.childNodes).map(canonical)])
}
const same = (a: Node, b: Node) => canonical(a) === canonical(b)
function only(node: Element, names: string[]) {
  require(Array.from(node.childNodes).every(n => n.nodeType === 1 && (n as Element).namespaceURI === W && names.includes((n as Element).localName)), "unsupported_target_content")
}
function rowAt(doc: Document, view: Target) {
  const body = children(doc.documentElement!, "body")
  require(body.length === 1, "invalid_body")
  const tables = children(body[0]!, "tbl"), table = tables[view.table - 1], row = table && children(table, "tr")[view.row - 1]
  require(table && row, "target_row_unavailable")
  return { table, row }
}
// Structural node extents include zero-text markers. A range may encompass the
// row with BOTH endpoints outside it, so checking descendants alone is unsafe.
function storyRanges(doc: Document) {
  const extents = new Map<Node, [number, number]>()
  let cursor = 0
  const visit = (node: Node) => {
    const start = cursor++
    // xmldom leaves have no childNodes collection. They still occupy structural
    // extent; every element (including unknown markers) is visited normally.
    if (node.nodeType === 9 || node.nodeType === 1) for (const child of Array.from(node.childNodes)) visit(child)
    extents.set(node, [start, cursor++])
  }
  visit(doc)
  require(!all(doc, "fldChar").length && !all(doc, "instrText").length && !all(doc, "delInstrText").length, "complex_story_field_unsupported")
  const families = [["bookmarkStart", "bookmarkEnd"], ["commentRangeStart", "commentRangeEnd"], ["permStart", "permEnd"]] as const
  const known = new Set<string>(families.flat()), intervals: [number, number][] = []
  require(!all(doc, "*").some(el => /Range(?:Start|End)$/.test(el.localName) && !known.has(el.localName)), "unsupported_story_range")
  for (const [startName, endName] of families) {
    const starts = all(doc, startName), ends = all(doc, endName)
    const id = (el: Element) => { const value = el.getAttributeNS(W, "id"); require(typeof value === "string" && /^[+-]?\d+$/.test(value), "invalid_story_range"); return BigInt(value).toString() }
    const startIds = starts.map(id), endIds = ends.map(id)
    require(new Set(startIds).size === starts.length && new Set(endIds).size === ends.length && starts.length === ends.length, "ambiguous_story_range")
    for (const [index, start] of starts.entries()) {
      const at = endIds.indexOf(startIds[index]!)
      require(at >= 0, "ambiguous_story_range")
      const a = extents.get(start)![0], b = extents.get(ends[at]!)![1]
      require(a < b, "reversed_story_range")
      intervals.push([a, b])
    }
  }
  return { extents, intervals }
}
function targetContent(pkg: Package, view: Target, ranges: ReturnType<typeof storyRanges>) {
  const { table, row } = rowAt(pkg.body, view)
  require(children(table, "tr").length > 1 && !all(row, "tblHeader").length, "sole_or_header_row_unsupported")
  require(Array.from(row.attributes).every(a => a.namespaceURI === W && /^rsid(?:R|RPr|Del|Tr)$/.test(a.localName)), "unsupported_row_attributes")
  const [start, end] = ranges.extents.get(row)!
  require(ranges.intervals.every(([a, b]) => b < start || a > end), "annotation_range_overlaps_target")
  require(!Object.values(pkg.docs).some(doc => ["vanish", "webHidden", "specVanish"].some(name => all(doc, name).length)), "hidden_content_unsupported")
  const rowPr = children(row, "trPr")
  require(rowPr.length <= 1 && rowPr.every(p => !p.attributes.length), "unsupported_row_properties")
  // Native Reject removes an empty trPr after stripping its deletion mark. An
  // originally empty container therefore cannot round-trip exactly. Do not
  // normalize it away or manufacture a property to retain it.
  require(!rowPr.length || rowPr[0]!.childNodes.length > 0, "empty_source_row_properties_unsupported")
  if (rowPr[0]) only(rowPr[0], ["cantSplit", "trHeight", "jc", "tblCellSpacing"])
  for (const prop of rowPr.flatMap(kids)) require(!prop.childNodes.length && Array.from(prop.attributes).every(a => a.namespaceURI === W), "unsupported_row_properties")
  const texts = children(row, "tc").map(cell => {
    require(!cell.attributes.length, "unsupported_cell_attributes")
    const cellPr = children(cell, "tcPr")
    require(cellPr.length <= 1 && cellPr.every(p => !p.attributes.length), "unsupported_cell_properties")
    if (cellPr[0]) only(cellPr[0], ["tcW", "shd", "tcBorders", "tcMar", "vAlign", "textDirection", "noWrap", "tcFitText"])
    for (const prop of cellPr.flatMap(kids)) {
      if (["tcBorders", "tcMar"].includes(prop.localName)) only(prop, ["top", "bottom", "left", "right", "start", "end", "insideH", "insideV", "tl2br", "tr2bl"])
      else require(!prop.childNodes.length, "unsupported_cell_properties")
      for (const node of [prop, ...kids(prop)]) require(Array.from(node.attributes).every(a => a.namespaceURI === W) && (node === prop || !node.childNodes.length), "unsupported_cell_properties")
    }
    const paragraphs = children(cell, "p")
    require(paragraphs.length === 1, "multi_paragraph_cell_unsupported")
    const p = paragraphs[0]!
    require(Array.from(p.attributes).every(a => a.namespaceURI === W && /^rsid(?:R|RPr|RDefault|Del)$/.test(a.localName) || a.namespaceURI === "http://schemas.microsoft.com/office/word/2010/wordml" && ["paraId", "textId"].includes(a.localName)), "unsupported_paragraph_attributes")
    only(p, ["pPr", "r"])
    const pPr = children(p, "pPr")
    require(pPr.length <= 1 && (!pPr.length || !pPr[0]!.attributes.length), "unsupported_paragraph_properties")
    if (pPr[0]) only(pPr[0], ["spacing", "ind", "jc", "keepNext", "keepLines", "widowControl", "pageBreakBefore", "contextualSpacing"])
    for (const prop of pPr.flatMap(kids)) require(!prop.childNodes.length && Array.from(prop.attributes).every(a => a.namespaceURI === W), "unsupported_paragraph_properties")
    return children(p, "r").map(run => {
      require(!run.attributes.length, "unsupported_literal_run")
      only(run, ["rPr", "t"])
      const properties = children(run, "rPr")
      require(properties.length <= 1 && (!properties.length || !properties[0]!.attributes.length), "unsupported_run_properties")
      if (properties[0]) only(properties[0], ["b", "bCs", "i", "iCs", "u", "strike", "dstrike", "caps", "smallCaps", "color", "sz", "szCs", "rFonts", "lang", "highlight", "vertAlign", "kern", "spacing", "position"])
      for (const prop of properties.flatMap(kids)) require(!prop.childNodes.length && Array.from(prop.attributes).every(a => a.namespaceURI === W), "unsupported_run_properties")
      require(children(run, "t").length > 0, "unsupported_literal_run")
      return children(run, "t").map(t => {
        require(Array.from(t.childNodes).every(n => n.nodeType === 3) && Array.from(t.attributes).every(a => a.namespaceURI === XML && a.localName === "space" && ["preserve", "default"].includes(a.value)), "unsupported_literal_text")
        return t.textContent ?? ""
      }).join("")
    }).join("")
  })
  require(texts.length === view.columns && texts.some(t => t.length) && texts.join("").length <= 24000 && texts.every(t => t.length <= 4000 && t.isWellFormed() && !/[\u0000-\u001f\u007f]/.test(t)) && JSON.stringify(texts) === JSON.stringify(view.cells), "incomplete_visible_cells")
  return row
}
/** Offered coordinates are display only; the opaque handle binds the source and
 * exact canonical SDK table/row/cell identities already checked by insertion. */
export async function docxTableRowDeletionTargets(bytes: Uint8Array) {
  const source = new Uint8Array(bytes), packet = await docxTableRowInsertionTargets(source), pkg = unpack(source), ranges = storyRanges(pkg.body)
  return { source_sha256: packet.source_sha256, targets: packet.targets.filter(view => {
    try { targetContent(pkg, view, ranges); return true } catch (error) {
      if (error instanceof Error && error.message.startsWith("docx_row_deletion:")) return false
      throw error
    }
  }) }
}
function request(value: DocxTableRowDeletion, source: Uint8Array) {
  const input = requestSchema.parse(value)
  require(input.source_sha256 === hash(source), "stale_source")
  return input
}
async function open(bytes: Uint8Array) { return (await docxodus()).openDocxSession(new Uint8Array(bytes), { persistAnchorIds: false, smartQuotes: false }) }
async function locate(source: Uint8Array, pkg: Package, session: Session, input: DocxTableRowDeletion) {
  const packet = await docxTableRowDeletionTargets(source), view = packet.targets.find(t => t.target_handle === input.target_handle)
  require(view, "target_not_eligible")
  const tableIds = Object.entries(session.project().anchorIndex).filter(([, a]) => a.kind === "tbl" && a.partUri === "/word/document.xml")
  const id = tableIds[view.table - 1]?.[0], result = id && session.getTableMetadata(id), metadata = result && result.metadata, row = metadata?.rows[view.row - 1], cell = row?.cells[0]
  require(result && result.success && id && row && cell && hash(JSON.stringify([input.source_sha256, id, row.anchor.id, cell.anchor.id])) === input.target_handle, "canonical_target_changed")
  return { view, sourceRow: targetContent(pkg, view, storyRanges(pkg.body)), cell: cell.anchor.id, paragraphIds: row.cells.flatMap(c => c.paragraphAnchors.map(p => p.id)), rowId: row.anchor.id }
}
function settingsBase(doc: Document) {
  const copy = doc.cloneNode(true) as Document, marks = all(copy, "trackRevisions")
  require(marks.length <= 1 && marks.every(mark => mark.parentNode === copy.documentElement && !mark.childNodes.length && Array.from(mark.attributes).every(a => a.namespaceURI === W && a.localName === "val" && ["1", "true", "on", "0", "false", "off"].includes(a.value))), "unsupported_settings")
  marks.forEach(mark => mark.parentNode!.removeChild(mark)); return copy
}
function deliver(source: Package, bytes: Uint8Array, tracked: boolean) {
  const output = unpack(bytes)
  require(JSON.stringify(Object.keys(output.parts).sort()) === JSON.stringify(Object.keys(source.parts).sort()), "package_inventory_changed")
  for (const name of Object.keys(source.parts)) if (name !== BODY) require(name === SETTINGS
    ? same(settingsBase(source.settings), settingsBase(output.settings))
    : source.docs[name] ? same(source.docs[name]!, output.docs[name]!) : hash(source.parts[name]!) === hash(output.parts[name]!), `unrelated_part_changed:${name}`)
  const parts: Record<string, Uint8Array> = { ...source.parts, [BODY]: output.parts[BODY]! }
  if (tracked) {
    const mark = all(output.settings, "trackRevisions")
    require(mark.length === 1 && !["0", "false", "off"].includes(mark[0]!.getAttributeNS(W, "val") ?? ""), "tracking_not_enabled")
    parts[SETTINGS] = output.parts[SETTINGS]!
  }
  return zipSync(parts, { mtime: new Date(1980, 0, 1) })
}
function digest(session: Session) {
  const value = session.getPackageManifest().rawPackageBytesDigest
  require(value?.algorithm === "SHA-256" && typeof value.value === "string" && /^[a-f0-9]{64}$/.test(value.value), "package_digest_unavailable")
  return value.value
}
// Keep native identity AND durable raw carrier extents; projected anchors can
// regenerate on stripped saves and are not original-source identity.
function signature(revisions: Revision[], pkg: Package, body = pkg.body) {
  return JSON.stringify(revisions.toSorted((a, b) => a.id.localeCompare(b.id)).map(({ anchorId: _anchor, affectedAnchors: _affected, ...revision }) => ({ revision, extents: revision.constituentKeys.map(key => {
    const match = /^kind=([^:]+):(.+):([^:]+):([^:]+)$/.exec(key), part = revision.partUri === "/word/document.xml" ? body : pkg.docs[revision.partUri.replace(/^\//, "")]
    require(match && part, "old_revision_extent_unavailable")
    const carriers = Array.from(part.getElementsByTagNameNS(match[2]!, match[3]!)).filter(el => el.getAttributeNS(W, "id") === match[4])
    require(carriers.length === 1, "old_revision_extent_ambiguous")
    const carrier = carriers[0]!, path: number[] = []
    let node: Node = carrier
    while (node.parentNode && node.parentNode.nodeType !== 9) { path.unshift(Array.from(node.parentNode.childNodes).indexOf(node as ChildNode)); node = node.parentNode }
    require(node === part.documentElement, "old_revision_extent_unavailable")
    return { key, path, subtree: canonical(carrier) }
  }) })))
}
type Carrier = { key: string; id: string; element: Element; kind: "RowMark" | "ParaMark" | "Content"; paragraph?: number }
function inverse(row: Element, original: Element, input: DocxTableRowDeletion) {
  const found: Carrier[] = [], copy = row.cloneNode(true) as Element
  const paragraphs = children(row, "tc").flatMap(c => children(c, "p")), originalParagraphs = children(original, "tc").flatMap(c => children(c, "p"))
  const add = (element: Element, kind: Carrier["kind"], paragraph?: number) => {
    const id = element.getAttributeNS(W, "id")
    require(typeof id === "string" && /^(0|[1-9]\d*)$/.test(id) && element.getAttributeNS(W, "author") === input.author && element.getAttributeNS(W, "date") && Array.from(element.attributes).every(a => a.namespaceURI === W && ["id", "author", "date"].includes(a.localName)), "invalid_deletion_carrier")
    if (kind !== "Content") require(!element.childNodes.length, "invalid_deletion_mark")
    else require(kids(element).length === 1 && children(element, "r").length === 1 && element.childNodes.length === 1, "invalid_deletion_wrapper")
    found.push({ key: `kind=${kind}:${W}:del:${id}`, id, element, kind, paragraph })
  }
  const rowMark = children(children(row, "trPr")[0] ?? row, "del")
  require(rowMark.length === 1 && rowMark[0]!.parentNode !== row, "missing_row_deletion")
  add(rowMark[0]!, "RowMark")
  for (const [index, p] of paragraphs.entries()) {
    const props = children(p, "pPr"), runProps = props[0] && children(props[0], "rPr"), mark = runProps?.[0] && children(runProps[0], "del")
    require(mark?.length === 1, "missing_paragraph_deletion")
    add(mark[0]!, "ParaMark", index)
    const content = children(p, "del")
    content.forEach(el => add(el, "Content", index))
    require(content.length === children(originalParagraphs[index]!, "r").length, "incomplete_deleted_runs")
  }
  require(new Set(found.map(c => c.id)).size === found.length && all(row, "del").length === found.length, "unclaimed_deletion_carrier")
  require(new Set(found.map(c => c.element.getAttributeNS(W, "date"))).size === 1, "deletion_stamp_changed")
  for (const wrapper of all(copy, "del")) {
    if (wrapper.parentNode?.nodeType === 1 && (wrapper.parentNode as Element).localName === "p") {
      const run = children(wrapper, "r")[0]!
      for (const t of all(run, "delText")) {
        const originalText = copy.ownerDocument!.createElementNS(W, "w:t")
        for (const attr of Array.from(t.attributes)) originalText.setAttributeNS(attr.namespaceURI, attr.name, attr.value)
        for (const node of Array.from(t.childNodes)) originalText.appendChild(node.cloneNode(true))
        t.parentNode!.replaceChild(originalText, t)
      }
      wrapper.parentNode!.replaceChild(run, wrapper)
    } else wrapper.parentNode!.removeChild(wrapper)
  }
  // Delete only property containers absent in the source. Existing empty or
  // formatted containers are significant and must survive the inverse.
  const originalRowPr = children(original, "trPr"), copyRowPr = children(copy, "trPr")
  if (!originalRowPr.length && copyRowPr.length === 1 && !copyRowPr[0]!.childNodes.length && !copyRowPr[0]!.attributes.length) copy.removeChild(copyRowPr[0]!)
  const restoredParagraphs = children(copy, "tc").flatMap(c => children(c, "p"))
  for (const [index, p] of restoredParagraphs.entries()) {
    const originalPr = children(originalParagraphs[index]!, "pPr"), props = children(p, "pPr"), runProps = props[0] && children(props[0], "rPr")
    if (runProps?.length === 1 && !children(originalPr[0] ?? originalParagraphs[index]!, "rPr").length && !runProps[0]!.childNodes.length && !runProps[0]!.attributes.length) props[0]!.removeChild(runProps[0]!)
    if (!originalPr.length && props.length === 1 && !props[0]!.childNodes.length && !props[0]!.attributes.length) p.removeChild(props[0]!)
  }
  require(same(copy, original), "tracked_inverse_changed_source_row")
  return found
}
async function inspect(sourceBytes: Uint8Array, bytes: Uint8Array, input: DocxTableRowDeletion) {
  const source = unpack(sourceBytes), output = unpack(bytes), baseline = await open(sourceBytes), session = await open(bytes)
  try {
    const located = await locate(sourceBytes, source, baseline, input), before = baseline.listRevisions(), row = rowAt(output.body, located.view).row
    deliver(source, bytes, true)
    const physical = inverse(row, located.sourceRow, input), restored = output.body.cloneNode(true) as Document
    const replace = rowAt(restored, located.view).row
    replace.parentNode!.replaceChild(restored.importNode(located.sourceRow, true), replace)
    require(same(restored, source.body), "unrelated_main_content_changed")
    const oldIds = new Set(before.map(r => r.id)), current = session.listRevisions(), added = current.filter(r => !oldIds.has(r.id))
    require(new Set(current.map(r => r.id)).size === current.length && signature(current.filter(r => oldIds.has(r.id)), output, restored) === signature(before, source), "old_revision_changed")
    const ids = Object.entries(session.project().anchorIndex).filter(([, a]) => a.kind === "tbl" && a.partUri === "/word/document.xml"), result = session.getTableMetadata(ids[located.view.table - 1]![0]), meta = result.metadata?.rows[located.view.row - 1]
    require(result.success && meta && meta.cells.length === located.view.columns, "deleted_row_extent_unavailable")
    const paragraphIds = meta.cells.flatMap(c => c.paragraphAnchors.map(p => p.id)), claimed = new Set<string>(), sourceRevisionKeys = new Set(before.flatMap(r => r.constituentKeys))
    for (const carrier of physical) require(!sourceRevisionKeys.has(carrier.key), "source_revision_id_collision")
    for (const revision of added) {
      require(revision.family === "row_delete" && revision.type === "delete" && revision.resolutionStatus === "supported" && revision.scope === "body" && revision.partUri === "/word/document.xml" && revision.author === input.author, "unexpected_native_revision")
      const owned = physical.filter(c => revision.constituentKeys.includes(c.key))
      require(owned.length === physical.length && owned.length === revision.constituentKeys.length && new Set(revision.constituentKeys).size === owned.length && owned.every(c => !claimed.has(c.key)) && JSON.stringify(revision.constituentIds.toSorted()) === JSON.stringify(owned.map(c => c.id).toSorted()), "native_membership_changed")
      require(revision.date === physical[0]!.element.getAttributeNS(W, "date") && JSON.stringify(revision.affectedAnchors.map(a => a.id).sort()) === JSON.stringify([meta.anchor.id, ...paragraphIds].sort()) && revision.anchorId === meta.anchor.id, "native_extent_changed")
      owned.forEach(c => claimed.add(c.key))
    }
    require(added.length === 1 && claimed.size === physical.length, "unclaimed_native_carrier")
    return { source, output, session, located, before, revisions: added }
  } catch (error) { session.close(); throw error } finally { baseline.close() }
}
/** Detached summaries from this verified artifact; no live session, DOM or
 * caller-provided revision ID grants authority to the review wrapper. */
export async function verifyDocxTableRowDeletion(sourceBytes: Uint8Array, authored: Uint8Array, value: DocxTableRowDeletion) {
  const source = new Uint8Array(sourceBytes), input = request(value, source), checked = await inspect(source, new Uint8Array(authored), input)
  try {
    const row = rowAt(checked.output.body, checked.located.view).row, paragraphs = all(checked.output.body, "p")
    return { revisions: structuredClone(checked.revisions), paragraphs: children(row, "tc").map((cell, index) => {
      const paragraph = children(cell, "p")[0]!, xml_paragraph_index = paragraphs.indexOf(paragraph)
      require(xml_paragraph_index >= 0, "deleted_paragraph_unavailable")
      return { cell: index, xml_paragraph_index, text: checked.located.view.cells[index]! }
    }) }
  } finally { checked.session.close() }
}
/** Reconstruct complete ownership from THIS artifact, never caller-supplied IDs. */
export async function resolveDocxTableRowDeletion(sourceBytes: Uint8Array, authored: Uint8Array, value: DocxTableRowDeletion, decision: "accept" | "reject") {
  const source = new Uint8Array(sourceBytes), input = request(value, source)
  require(decision === "accept" || decision === "reject", "invalid_decision")
  const checked = await inspect(source, new Uint8Array(authored), input)
  try {
    const { session } = checked, before = digest(session), version = session.getVersion(), member = checked.revisions[0]!
    const result = session.executeBatch([{ tool: "table", action: decision, mutation: () => {
      const found = session.listRevisions().filter(r => r.id === member.id && JSON.stringify(r.constituentKeys) === JSON.stringify(member.constituentKeys))
      require(found.length === 1, "native_group_changed")
      return decision === "accept" ? session.acceptRevision(found[0]!.id) : session.rejectRevision(found[0]!.id)
    } }], "atomic")
    if (!result.success) {
      require(session.getVersion() === version && digest(session) === before, "failed_resolution_not_rolled_back")
      throw new Error("docx_row_deletion:resolution_failed_rolled_back")
    }
    const bytes = deliver(checked.source, session.save(), false), output = unpack(bytes), intended = checked.source.body.cloneNode(true) as Document
    if (decision === "accept") { const row = rowAt(intended, checked.located.view).row; row.parentNode!.removeChild(row) }
    require(same(output.body, intended), "resolution_changed_intended_main")
    require(signature(session.listRevisions(), output) === signature(checked.before, checked.source, intended), "resolution_changed_old_revision")
    const reopened = await open(bytes)
    try { require(signature(reopened.listRevisions(), output) === signature(checked.before, checked.source, intended), "reopened_revision_changed") } finally { reopened.close() }
    return decision === "reject" ? new Uint8Array(source) : bytes
  } finally { checked.session.close() }
}
/** Pure disposable authoring: no bytes escape until both selective decisions
 * prove exact rejection and only the explicitly selected row's removal. */
export async function authorDocxTableRowDeletion(sourceBytes: Uint8Array, value: DocxTableRowDeletion) {
  const source = new Uint8Array(sourceBytes), input = request(value, source), pkg = unpack(source), session = await open(source)
  let bytes: Uint8Array
  try {
    const located = await locate(source, pkg, session, input)
    session.setRevisionAuthor(input.author); session.setTrackedChanges((await docxodus()).TrackedChangeMode.RenderInline)
    const before = digest(session), version = session.getVersion(), result = session.executeBatch([{ tool: "table", action: "delete_row", mutation: () => session.deleteTableRow(located.cell) }], "atomic")
    if (!result.success) {
      require(session.getVersion() === version && digest(session) === before, "failed_authoring_not_rolled_back")
      throw new Error("docx_row_deletion:authoring_failed_rolled_back")
    }
    bytes = deliver(pkg, session.save(), true)
  } finally { session.close() }
  const checked = await inspect(source, bytes, input)
  let revisions: { id: string; family: string }[]
  try { revisions = checked.revisions.map(({ id, family }) => ({ id, family })) } finally { checked.session.close() }
  await resolveDocxTableRowDeletion(source, bytes, input, "reject")
  await resolveDocxTableRowDeletion(source, bytes, input, "accept")
  return { bytes, revisions }
}
