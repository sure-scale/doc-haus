import { createHash } from "node:crypto"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { z } from "zod"
import { docxodus } from "./docxodus"

// Source-bound authoring proof; callers own review and persistence. Never import its output into live Word.
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const PT = "http://powertools.codeplex.com/2011"
const BODY = "word/document.xml", SETTINGS = "word/settings.xml"
const hash = (value: Uint8Array | string) => createHash("sha256").update(value).digest("hex")
type Session = ReturnType<Awaited<ReturnType<typeof docxodus>>["openDocxSession"]>
type Revision = ReturnType<Session["listRevisions"]>[number]
type Metadata = NonNullable<ReturnType<Session["getTableMetadata"]>["metadata"]>
type Package = ReturnType<typeof unpack>
const requestSchema = z.object({ source_sha256: z.string().regex(/^[a-f0-9]{64}$/), target_handle: z.string().min(1),
  position: z.enum(["before", "after"]), cells: z.array(z.string().max(4000)).min(1).max(32),
  author: z.string().trim().min(1).max(100) }).strict()
export type DocxTableRowInsertion = z.infer<typeof requestSchema>
type Target = { target_handle: string; table: number; row: number; cells: string[]; columns: number }
type Located = { view: Target; table: Element; metadata: Metadata; row: number; cell: string }

function require(value: unknown, reason: string): asserts value { if (!value) throw new Error(`docx_row_insertion:${reason}`) }
const all = (node: Document | Element, name: string) => Array.from(node.getElementsByTagNameNS(W, name))
const kids = (node: Node) => Array.from(node.childNodes).filter((n): n is Element => n.nodeType === 1)
const children = (node: Node, name: string) => kids(node).filter(n => n.namespaceURI === W && n.localName === name)
function parse(xml: string) {
  require(!/<!DOCTYPE|<!ENTITY/i.test(xml), "unsupported_xml")
  const fail = () => { throw new Error("docx_row_insertion:invalid_xml") }
  return new DOMParser({ errorHandler: { warning: fail, error: fail, fatalError: fail } }).parseFromString(xml, "application/xml")
}
function unpack(bytes: Uint8Array) {
  const parts = unzipSync(bytes)
  require(parts[BODY] && parts[SETTINGS], "existing_main_and_settings_required")
  const docs = Object.fromEntries(Object.entries(parts).filter(([name]) => /\.(xml|rels)$/.test(name)).map(([name, bytes]) => [name, parse(strFromU8(bytes))]))
  return { parts, docs, body: docs[BODY]!, settings: docs[SETTINGS]! }
}
// Preserve every namespace binding, attribute, text, PI and comment. Only XML
// declarations/attribute order/empty-element serialization are non-semantic.
function canonical(node: Node): string {
  if (node.nodeType === 9) return JSON.stringify(Array.from(node.childNodes).filter(n => !(n.nodeType === 7 && n.nodeName === "xml")).map(canonical))
  if (node.nodeType !== 1) return JSON.stringify([node.nodeType, node.nodeName, node.nodeValue])
  const el = node as Element
  return JSON.stringify([el.namespaceURI, el.nodeName, Array.from(el.attributes).map(a => [a.namespaceURI, a.name, a.value]).sort(), Array.from(el.childNodes).map(canonical)])
}
function same(a: Node, b: Node) { return canonical(a) === canonical(b) }
function serialized(node: Node) { return new XMLSerializer().serializeToString(node) }
function standalone(node: Element, sdk = false) {
  const copy = parse(serialized(node)).documentElement!
  // raw.getXml pretty-prints fragments. This mapping-only comparison removes
  // its indentation and generated IDs; the delivered full-package proof below
  // compares original text and namespace bindings without this normalization.
  for (const el of [copy, ...Array.from(copy.getElementsByTagName("*"))]) {
    if (sdk) for (const attr of Array.from(el.attributes)) if (attr.namespaceURI === PT || attr.value === PT && attr.namespaceURI === "http://www.w3.org/2000/xmlns/") el.removeAttributeNode(attr)
    if (el.namespaceURI !== W || !["t", "delText", "instrText"].includes(el.localName))
      for (const child of Array.from(el.childNodes)) if (child.nodeType === 3 && !child.nodeValue?.trim()) el.removeChild(child)
  }
  return copy
}
function only(node: Element, names: string[]) {
  require(Array.from(node.childNodes).every(n => n.nodeType === 1 && (n as Element).namespaceURI === W && names.includes((n as Element).localName)), "unsupported_new_row_content")
}
function plain(node: Element) { return all(node, "t").map(t => t.textContent ?? "").join("") }
function packageDigest(session: Session) {
  const digest = session.getPackageManifest().rawPackageBytesDigest
  require(digest && digest.algorithm === "SHA-256" && typeof digest.value === "string" && /^[a-f0-9]{64}$/.test(digest.value), "package_digest_unavailable")
  return `${digest.algorithm}:${digest.value}`
}
function revisionSignature(rows: Revision[], pkg: Package, body = pkg.body) {
  return JSON.stringify(rows.toSorted((a, b) => a.id.localeCompare(b.id)).map(revision => {
    // Projector anchors may regenerate on a stripped save. Replace them with
    // exact source-carrier locations; retain every other registry field.
    const { anchorId: _anchor, affectedAnchors: _affected, ...identity } = revision
    const part = revision.partUri === "/word/document.xml" ? body : pkg.docs[revision.partUri.replace(/^\//, "")]
    require(part && revision.constituentKeys.length > 0 && new Set(revision.constituentKeys).size === revision.constituentKeys.length, "existing_revision_carrier_unavailable")
    const extents = revision.constituentKeys.map(key => {
      const match = /^kind=([^:]+):(.+):([^:]+):([^:]+)$/.exec(key)
      require(match, "existing_revision_carrier_unavailable")
      const [, , namespace, name, id] = match
      const found = Array.from(part.getElementsByTagNameNS(namespace!, name!)).filter(el => el.getAttributeNS(W, "id") === id)
      require(found.length === 1, "existing_revision_carrier_ambiguous")
      const carrier = found[0]!, path: number[] = []
      let node: Node = carrier
      while (node.parentNode && node.parentNode.nodeType !== 9) {
        path.unshift(Array.from(node.parentNode.childNodes).indexOf(node as ChildNode))
        node = node.parentNode
      }
      require(node === part.documentElement, "existing_revision_extent_unavailable")
      return { key, path, subtree: canonical(carrier) }
    })
    return { identity, extents }
  }))
}
function settingBase(doc: Document) {
  const copy = doc.cloneNode(true) as Document
  require(all(copy, "trackRevisions").length === children(copy.documentElement!, "trackRevisions").length, "unsupported_settings")
  require(all(copy, "trackRevisions").length <= 1, "unsupported_settings")
  for (const el of all(copy, "trackRevisions")) el.parentNode!.removeChild(el)
  return copy
}
// Validate the SDK package first, then retain the original bytes of untouched
// parts. No dependency addition, relationship change or namespace change passes.
function deliver(source: Package, saved: Uint8Array, tracked: boolean) {
  const candidate = unpack(saved)
  require(JSON.stringify(Object.keys(source.parts).sort()) === JSON.stringify(Object.keys(candidate.parts).sort()), "package_inventory_changed")
  for (const name of Object.keys(source.parts)) {
    if (name === BODY) continue
    if (name === SETTINGS) {
      require(same(settingBase(source.settings), settingBase(candidate.settings)), "settings_changed")
      continue
    }
    require(source.docs[name] ? same(source.docs[name]!, candidate.docs[name]!) : hash(source.parts[name]!) === hash(candidate.parts[name]!), `unrelated_part_changed:${name}`)
  }
  const parts: Record<string, Uint8Array> = { ...source.parts, [BODY]: candidate.parts[BODY]! }
  if (tracked) {
    const marks = all(candidate.settings, "trackRevisions")
    require(marks.length === 1 && !marks[0]!.childNodes.length && Array.from(marks[0]!.attributes).every(a => a.namespaceURI === W && a.localName === "val" && ["1", "true", "on"].includes(a.value)), "tracking_not_enabled")
    // Keep the SDK's schema position for the one admitted setting change.
    parts[SETTINGS] = candidate.parts[SETTINGS]!
  }
  return zipSync(parts, { mtime: new Date(1980, 0, 1) })
}
// Only the styles inherited by this table and newly created default paragraphs
// determine row semantics. Unused definitions remain protected package data.
function safeRowStyles(styles: Document | undefined, table: Element) {
  const definitions = styles ? children(styles.documentElement!, "style") : []
  const truthy = (value: string | null) => value !== null && ["1", "true", "on"].includes(value)
  const defaults = (type: string) => definitions.filter(s => s.getAttributeNS(W, "type") === type && truthy(s.getAttributeNS(W, "default")))
  const tableDefaults = defaults("table"), paragraphDefaults = defaults("paragraph")
  if (tableDefaults.length > 1 || paragraphDefaults.length > 1 ||
    [...tableDefaults, ...paragraphDefaults].some(s => !s.getAttributeNS(W, "styleId"))) return false
  if (styles && all(styles, "docDefaults").some(d => all(d, "numPr").length)) return false
  const tblPr = children(table, "tblPr")
  if (tblPr.length > 1) return false
  const refs = tblPr.length ? children(tblPr[0]!, "tblStyle") : []
  if (refs.length > 1) return false
  const explicit = refs[0]?.getAttributeNS(W, "val")
  if (refs.length && !explicit) return false
  const chain = (type: string, id: string | null | undefined) => {
    const seen = new Set<string>()
    while (id) {
      if (seen.has(id)) return false
      seen.add(id)
      const matches = definitions.filter(s => s.getAttributeNS(W, "styleId") === id)
      if (matches.length !== 1 || matches[0]!.getAttributeNS(W, "type") !== type) return false
      const style = matches[0]!
      if (all(style, "numPr").length || type === "table" && all(style, "tblStylePr").length) return false
      const bases = children(style, "basedOn")
      if (bases.length > 1 || bases.length && !bases[0]!.getAttributeNS(W, "val")) return false
      id = bases[0]?.getAttributeNS(W, "val")
    }
    return true
  }
  return chain("table", explicit ?? tableDefaults[0]?.getAttributeNS(W, "styleId")) &&
    chain("paragraph", paragraphDefaults[0]?.getAttributeNS(W, "styleId"))
}
function locate(source: Package, session: Session, digest: string): Located[] {
  require(!Object.values(source.docs).some(doc => Array.from(doc.getElementsByTagName("*")).some(el => Array.from(el.attributes).some(a => a.namespaceURI === PT || a.value === PT))), "persisted_sdk_identity_unsupported")
  require(!all(source.settings, "documentProtection").length, "protected_document")
  const styles = source.docs["word/styles.xml"]
  const root = source.body.documentElement!, body = children(root, "body")
  require(body.length === 1, "invalid_body")
  const tables = children(body[0]!, "tbl")
  const projected = Object.entries(session.project().anchorIndex).filter(([, a]) => a.kind === "tbl" && a.partUri === "/word/document.xml")
  // Nested tables are not aligned by a guessed preorder in this first slice.
  require(projected.length === tables.length && all(source.body, "tbl").length === tables.length, "nested_or_wrapped_table_unsupported")
  const targets: Located[] = []
  for (const [index, table] of tables.entries()) {
    const id = projected[index]![0], result = session.getTableMetadata(id)
    if (!result.success || !result.metadata) continue
    const metadata = result.metadata
    require(same(standalone(table), standalone(parse(session.raw.getXml(id)).documentElement!, true)), "table_identity_changed")
    const rows = children(table, "tr"), grid = children(table, "tblGrid")
    if (!safeRowStyles(styles, table)) continue
    if (kids(table).some(n => n.namespaceURI !== W || !["tblPr", "tblGrid", "tr"].includes(n.localName)) || grid.length !== 1 ||
      !rows.length || rows.length > 1000 || metadata.rows.length !== rows.length || !metadata.columns.length || metadata.columns.length > 32 ||
      metadata.columns.some(c => c.isVirtual || c.widthTwips <= 0) || children(grid[0]!, "gridCol").length !== metadata.columns.length ||
      ["gridSpan", "vMerge", "hMerge", "gridBefore", "gridAfter", "wBefore", "wAfter", "cnfStyle", "numPr", "sdt", "customXml", "altChunk"].some(n => all(table, n).length) ||
      ["tblPr", "tblGrid", "trPr", "tcPr"].some(n => all(table, n).some(p => kids(p).some(e => /Change$/.test(e.localName) || ["ins", "del", "cellIns", "cellDel", "cellMerge"].includes(e.localName))))) continue
    const valid = rows.every((row, r) => {
      const cells = children(row, "tc"), meta = metadata.rows[r]!
      return kids(row).every(n => n.namespaceURI === W && ["trPr", "tc"].includes(n.localName)) && cells.length === metadata.columns.length &&
        meta.rowIndex === r && meta.gridBefore === 0 && meta.gridAfter === 0 && meta.cells.length === cells.length && meta.cells.every((c, col) =>
          c.rowIndex === r && c.columnIndex === col && c.rowSpan === 1 && c.columnSpan === 1 && c.verticalMerge === "none" && c.tableAnchorId === id && c.rowAnchorId === meta.anchor.id &&
          // This first slice exposes one paragraph per cell. Do not flatten
          // multiple paragraphs into misleading reference text such as Paymonthly.
          children(cells[col]!, "p").length === 1 && c.paragraphAnchors.length === 1 &&
          kids(cells[col]!).every(n => n.namespaceURI === W && ["tcPr", "p"].includes(n.localName)) &&
          same(standalone(cells[col]!), standalone(parse(session.raw.getXml(c.anchor.id)).documentElement!, true)))
    })
    if (!valid) continue
    for (const [row, entry] of metadata.rows.entries()) targets.push({ table, metadata, row, cell: entry.cells[0]!.anchor.id,
      view: { target_handle: hash(JSON.stringify([digest, id, entry.anchor.id, entry.cells[0]!.anchor.id])), table: index + 1, row: row + 1,
        cells: children(rows[row]!, "tc").map(plain), columns: metadata.columns.length } })
  }
  return targets
}
function request(value: DocxTableRowInsertion, source: Uint8Array) {
  const parsed = requestSchema.parse(value)
  require(parsed.source_sha256 === hash(source), "stale_source")
  require(parsed.cells.some(c => c.length) && parsed.cells.join("").length <= 24000 && parsed.cells.every(c => c === c.trim() && !/[\u0000-\u001f\u007f]/.test(c) && c.isWellFormed()), "unsupported_literal_cells")
  return parsed
}
async function open(source: Uint8Array) {
  const dx = await docxodus()
  return dx.openDocxSession(new Uint8Array(source), { persistAnchorIds: false, smartQuotes: false })
}
export async function docxTableRowInsertionTargets(bytes: Uint8Array) {
  const source = new Uint8Array(bytes), pkg = unpack(source), session = await open(source)
  try { return { source_sha256: hash(source), targets: locate(pkg, session, hash(source)).map(t => t.view) } }
  finally { session.close() }
}
function target(source: Package, session: Session, input: DocxTableRowInsertion) {
  const found = locate(source, session, input.source_sha256).find(t => t.view.target_handle === input.target_handle)
  require(found && input.cells.length === found.view.columns, "target_or_cell_count_changed")
  require(session.listRevisions().every(r => r.resolutionStatus === "supported"), "unsupported_existing_revision")
  return found
}
function literal(text: string) { return text.replace(/[\\`*_{}\[\]()#+.!<>|~\-]/g, "\\$&") }
function write(session: Session, cell: string, input: DocxTableRowInsertion) {
  let created: string[] = []
  // Materialize the checkpoint before measuring raw package rollback: the
  // SDK's first save may add XML declarations to loaded logical XML parts.
  session.save()
  const before = packageDigest(session), version = session.getVersion()
  const result = session.executeBatch([{ tool: "table", action: "insert_row", mutation: () => {
    const row = session.insertTableRow(cell, input.position)
    if (!row.success) return row
    created = row.created.map(c => c.id)
    require(created.length === input.cells.length && new Set(created).size === created.length && row.created.every(c => c.kind === "tc"), "created_cell_identity_changed")
    return [row, ...created.flatMap((id, i) => input.cells[i] ? [session.replaceCellContent(id, literal(input.cells[i]!))] : [])]
  } }], "atomic")
  if (!result.success) {
    require(session.getVersion() === version && packageDigest(session) === before, "failed_authoring_not_rolled_back")
    throw new Error("docx_row_insertion:authoring_failed_rolled_back")
  }
  return created
}
function rowAndRemainder(source: Package, output: Package, located: Located, input: DocxTableRowInsertion) {
  const clone = output.body.cloneNode(true) as Document
  const table = children(children(clone.documentElement!, "body")[0]!, "tbl")[located.view.table - 1]
  require(table, "inserted_table_missing")
  const rows = children(table, "tr"), originalRows = children(located.table, "tr")
  const at = located.row + Number(input.position === "after")
  require(rows.length === originalRows.length + 1, "inserted_row_count_changed")
  const row = rows[at]!
  table.removeChild(row)
  require(same(source.body, clone), "original_document_changed")
  const cells = children(row, "tc"), reference = children(originalRows[located.row]!, "tc")
  require(cells.length === input.cells.length, "inserted_cell_count_changed")
  for (const [i, cell] of cells.entries()) {
    const before = children(reference[i]!, "tcPr"), after = children(cell, "tcPr")
    require(before.length === after.length && (!before.length || same(before[0]!, after[0]!)), "inserted_cell_properties_changed")
    require(plain(cell) === input.cells[i], "inserted_cell_text_changed")
  }
  return { row, at, remainder: clone }
}
type Carrier = { key: string; id: string; element: Element; kind: "RowMark" | "ParaMark" | "Content"; paragraph?: number }
function carriers(row: Element, input: DocxTableRowInsertion): Carrier[] {
  require(!row.attributes.length, "unexpected_new_row_attributes")
  only(row, ["trPr", "tc"])
  const rowPr = children(row, "trPr")
  require(rowPr.length === 1 && !rowPr[0]!.attributes.length, "invalid_row_revision")
  only(rowPr[0]!, ["ins"])
  const marks = children(rowPr[0]!, "ins")
  require(marks.length === 1, "invalid_row_revision")
  const found: Carrier[] = []
  const add = (el: Element, kind: Carrier["kind"], paragraph?: number) => {
    const id = el.getAttributeNS(W, "id")
    require(typeof id === "string" && /^(0|[1-9]\d*)$/.test(id) && el.getAttributeNS(W, "author") === input.author && !!el.getAttributeNS(W, "date") &&
      Array.from(el.attributes).every(a => a.namespaceURI === W && ["id", "author", "date"].includes(a.localName)), "invalid_revision_carrier")
    found.push({ key: `kind=${kind}:${W}:ins:${id}`, id, element: el, kind, paragraph })
  }
  require(!marks[0]!.childNodes.length, "invalid_row_revision")
  add(marks[0]!, "RowMark")
  for (const [index, cell] of children(row, "tc").entries()) {
    require(!cell.attributes.length, "unexpected_new_cell_attributes")
    only(cell, ["tcPr", "p"])
    const paragraphs = children(cell, "p")
    require(paragraphs.length === 1 && !paragraphs[0]!.attributes.length, "unexpected_new_paragraph")
    const p = paragraphs[0]!
    only(p, ["pPr", "ins"])
    const pPr = children(p, "pPr")
    require(pPr.length === 1 && !pPr[0]!.attributes.length, "invalid_paragraph_mark")
    only(pPr[0]!, ["rPr"])
    const rPr = children(pPr[0]!, "rPr")
    require(rPr.length === 1 && !rPr[0]!.attributes.length, "invalid_paragraph_mark")
    only(rPr[0]!, ["ins"])
    const mark = children(rPr[0]!, "ins")
    require(mark.length === 1 && !mark[0]!.childNodes.length, "invalid_paragraph_mark")
    add(mark[0]!, "ParaMark", index)
    for (const content of children(p, "ins")) {
      add(content, "Content", index)
      only(content, ["r"])
      for (const run of children(content, "r")) {
        require(!run.attributes.length, "unexpected_literal_run")
        only(run, ["t"])
        require(children(run, "t").length === 1, "unexpected_literal_run")
        const text = children(run, "t")[0]!
        require(!kids(text).length && Array.from(text.childNodes).every(n => n.nodeType === 3) &&
          Array.from(text.attributes).every(a => a.namespaceURI === "http://www.w3.org/XML/1998/namespace" && a.localName === "space" && a.value === "preserve"), "unexpected_literal_text")
      }
    }
  }
  require(new Set(found.map(c => c.id)).size === found.length, "duplicate_new_carrier_id")
  return found
}
function membership(source: Package, output: Package, row: Element, remainder: Document, at: number, located: Located, before: Revision[], session: Session, input: DocxTableRowInsertion) {
  const physical = carriers(row, input), sourceIds = new Set(all(source.body, "*").flatMap(el => {
    if (!el.hasAttributeNS(W, "id")) return []
    const id = el.getAttributeNS(W, "id")
    require(typeof id === "string", "source_revision_id_unavailable")
    return [id.replace(/^\+?0+(?=\d)/, "")]
  }))
  require(physical.every(c => !sourceIds.has(c.id)), "source_revision_id_collision")
  const current = session.listRevisions(), old = new Set(before.map(r => r.id)), added = current.filter(r => !old.has(r.id))
  const oldNow = current.filter(r => old.has(r.id))
  require(new Set(current.map(r => r.id)).size === current.length && revisionSignature(oldNow, output, remainder) === revisionSignature(before, source), "existing_revision_changed")
  const tableIds = Object.entries(session.project().anchorIndex).filter(([, a]) => a.kind === "tbl" && a.partUri === "/word/document.xml")
  const result = session.getTableMetadata(tableIds[located.view.table - 1]![0]), meta = result.metadata?.rows[at]
  require(result.success && meta && meta.cells.length === input.cells.length, "new_row_identity_unavailable")
  const paragraphIds = meta.cells.map(c => { require(c.paragraphAnchors.length === 1, "new_paragraph_identity_unavailable"); return c.paragraphAnchors[0]!.id })
  const claimed = new Set<string>()
  for (const revision of added) {
    require(revision.resolutionStatus === "supported" && revision.type === "insert" && revision.partUri === "/word/document.xml" && revision.scope === "body" && revision.author === input.author &&
      ["row_insert", "content_insert"].includes(revision.family), "unexpected_revision")
    const owned = physical.filter(c => revision.constituentKeys.includes(c.key))
    require(owned.length > 0 && owned.length === revision.constituentKeys.length && new Set(revision.constituentKeys).size === owned.length &&
      JSON.stringify([...new Set(owned.map(c => c.id))].sort()) === JSON.stringify([...revision.constituentIds].sort()) &&
      owned.every(c => !claimed.has(c.key)), "revision_membership_changed")
    owned.forEach(c => claimed.add(c.key))
    if (revision.family === "row_insert") {
      require(owned.length === input.cells.length + 1 && owned.every(c => c.kind !== "Content") && owned.some(c => c.kind === "RowMark") && revision.anchorId === meta.anchor.id &&
        JSON.stringify(revision.affectedAnchors.map(a => a.id).sort()) === JSON.stringify([meta.anchor.id, ...paragraphIds].sort()), "row_revision_extent_changed")
    } else {
      const indices = new Set(owned.map(c => c.paragraph))
      require(owned.every(c => c.kind === "Content") && indices.size === 1, "content_revision_extent_changed")
      const index = owned[0]!.paragraph!
      require(revision.anchorId === paragraphIds[index] && revision.affectedAnchors.length === 1 && revision.affectedAnchors[0]!.id === paragraphIds[index] &&
        revision.text === owned.map(c => plain(c.element)).join(""), "content_revision_extent_changed")
    }
  }
  require(claimed.size === physical.length && added.filter(r => r.family === "row_insert").length === 1, "unclaimed_revision_carrier")
  return added
}
async function inspect(sourceBytes: Uint8Array, bytes: Uint8Array, input: DocxTableRowInsertion) {
  const source = unpack(sourceBytes), output = unpack(bytes), baseline = await open(sourceBytes), session = await open(bytes)
  try {
    const located = target(source, baseline, input), before = baseline.listRevisions()
    const final = deliver(source, bytes, true)
    // Validate the actual delivered bytes, not just the SDK's pre-restoration package.
    require(hash(unpack(final).parts[BODY]!) === hash(output.parts[BODY]!), "delivered_main_changed")
    const { row, at, remainder } = rowAndRemainder(source, output, located, input)
    const revisions = membership(source, output, row, remainder, at, located, before, session, input)
    return { source, output, session, located, before, revisions }
  } catch (error) { session.close(); throw error }
  finally { baseline.close() }
}
/** Read-only proof of actual carriers and result paragraphs in these bytes. */
export async function verifyDocxTableRowInsertion(sourceBytes: Uint8Array, authored: Uint8Array, value: DocxTableRowInsertion) {
  const source = new Uint8Array(sourceBytes), input = request(value, source), checked = await inspect(source, new Uint8Array(authored), input)
  try {
    const table = children(children(checked.output.body.documentElement!, "body")[0]!, "tbl")[checked.located.view.table - 1]!
    const row = children(table, "tr")[checked.located.row + Number(input.position === "after")]!
    const paragraphs = all(checked.output.body, "p")
    return { revisions: checked.revisions, paragraphs: children(row, "tc").map((cell, index) => {
      const paragraph = children(cell, "p")[0]!, xml_paragraph_index = paragraphs.indexOf(paragraph)
      require(xml_paragraph_index >= 0, "new_paragraph_identity_unavailable")
      return { cell: index, xml_paragraph_index, text: input.cells[index]! }
    }) }
  } finally { checked.session.close() }
}
/** Rebuild membership from carriers and the registry; caller-supplied IDs are never trusted. */
export async function resolveDocxTableRowInsertion(sourceBytes: Uint8Array, authored: Uint8Array, value: DocxTableRowInsertion, decision: "accept" | "reject") {
  const source = new Uint8Array(sourceBytes), input = request(value, source)
  require(decision === "accept" || decision === "reject", "invalid_decision")
  const checked = await inspect(source, new Uint8Array(authored), input)
  try {
    const { session } = checked, digest = packageDigest(session), version = session.getVersion()
    const order = checked.revisions.toSorted((a, b) => Number(a.family === "row_insert") - Number(b.family === "row_insert"))
    const result = session.executeBatch(order.map(r => ({ tool: "table", action: decision, mutation: () => decision === "accept" ? session.acceptRevision(r.id) : session.rejectRevision(r.id) })), "atomic")
    if (!result.success) {
      require(session.getVersion() === version && packageDigest(session) === digest, "failed_resolution_not_rolled_back")
      throw new Error("docx_row_insertion:resolution_failed_rolled_back")
    }
    const remaining = session.listRevisions()
    const bytes = deliver(checked.source, session.save(), false), saved = unpack(bytes)
    const remainder = decision === "reject" ? saved.body : rowAndRemainder(checked.source, saved, checked.located, input).remainder
    require(revisionSignature(remaining, saved, remainder) === revisionSignature(checked.before, checked.source), "resolution_changed_existing_revisions")
    if (decision === "reject") {
      require(same(saved.body, checked.source.body), "rejection_did_not_restore_source")
      return new Uint8Array(source)
    }
    const direct = await open(source)
    try {
      const reference = target(checked.source, direct, input)
      write(direct, reference.cell, input)
      const intended = deliver(checked.source, direct.save(), false)
      require(same(saved.body, unpack(intended).body), "acceptance_did_not_match_direct_insertion")
    } finally { direct.close() }
    const reopened = await open(bytes)
    try { require(revisionSignature(reopened.listRevisions(), saved, remainder) === revisionSignature(checked.before, checked.source), "saved_resolution_changed_existing_revisions") }
    finally { reopened.close() }
    return bytes
  } finally { checked.session.close() }
}
/** Author one source-bound row, proving both grouped decisions before returning bytes. */
export async function authorDocxTableRowInsertion(sourceBytes: Uint8Array, value: DocxTableRowInsertion) {
  const source = new Uint8Array(sourceBytes), input = request(value, source), pkg = unpack(source), session = await open(source)
  let bytes: Uint8Array
  try {
    const located = target(pkg, session, input)
    session.setRevisionAuthor(input.author)
    session.setTrackedChanges((await docxodus()).TrackedChangeMode.RenderInline)
    write(session, located.cell, input)
    bytes = deliver(pkg, session.save(), true)
  } finally { session.close() }
  const checked = await inspect(source, bytes, input)
  let revisions: { id: string; family: string }[]
  try { revisions = checked.revisions.map(({ id, family }) => ({ id, family })) }
  finally { checked.session.close() }
  await resolveDocxTableRowInsertion(source, bytes, input, "reject")
  await resolveDocxTableRowInsertion(source, bytes, input, "accept")
  return { bytes, revisions }
}
