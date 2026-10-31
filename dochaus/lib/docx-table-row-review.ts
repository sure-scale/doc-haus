import path from "node:path"
import { createHash } from "node:crypto"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { docxodus } from "./docxodus"
import { exportDocxComments, readDocxCommentParagraphPositions, readDocxComments } from "./docx-comments"
import { authorDocxTableRowInsertion, resolveDocxTableRowInsertion, verifyDocxTableRowInsertion, type DocxTableRowInsertion } from "./docx-table-row-insertion"

// Offline artifact only: a reason belongs to this one verified new row. This
// does not grant publication authority or make a row one native Word revision.
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", R = "http://schemas.openxmlformats.org/package/2006/relationships"
const CT = "http://schemas.openxmlformats.org/package/2006/content-types", XML = "http://www.w3.org/XML/1998/namespace"
const BODY = "word/document.xml", SETTINGS = "word/settings.xml", RELS = "word/_rels/document.xml.rels", TYPES = "[Content_Types].xml"
const COMMENT_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
const COMMENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
type Session = ReturnType<Awaited<ReturnType<typeof docxodus>>["openDocxSession"]>
type Revision = Awaited<ReturnType<typeof verifyDocxTableRowInsertion>>["revisions"][number]
const hash = (v: Uint8Array | string) => createHash("sha256").update(v).digest("hex")
function require(value: unknown, code: string): asserts value { if (!value) throw new Error(`docx_row_review:${code}`) }
function parse(text: string) {
  require(!/<!DOCTYPE|<!ENTITY/i.test(text), "unsupported_xml")
  const fail = () => { throw new Error("docx_row_review:invalid_xml") }
  return new DOMParser({ errorHandler: { warning: fail, error: fail, fatalError: fail } }).parseFromString(text, "application/xml")
}
const xml = (doc: Node) => new XMLSerializer().serializeToString(doc)
const all = (node: Document | Element, ns: string, name: string) => Array.from(node.getElementsByTagNameNS(ns, name))
const kids = (node: Node) => Array.from(node.childNodes).filter((n): n is Element => n.nodeType === 1)
function canonical(node: Node): string {
  if (node.nodeType === 9) return JSON.stringify(Array.from(node.childNodes).filter(n => !(n.nodeType === 7 && n.nodeName === "xml")).map(canonical))
  if (node.nodeType !== 1) return JSON.stringify([node.nodeType, node.nodeName, node.nodeValue])
  const el = node as Element
  return JSON.stringify([el.namespaceURI, el.nodeName, Array.from(el.attributes).map(a => [a.namespaceURI, a.name, a.value]).sort(), Array.from(el.childNodes).map(canonical)])
}
const same = (a: Node, b: Node) => canonical(a) === canonical(b)
function unpack(bytes: Uint8Array) {
  const parts = unzipSync(bytes), docs = Object.fromEntries(Object.entries(parts).filter(([name]) => /\.(xml|rels)$/.test(name)).map(([name, bytes]) => [name, parse(strFromU8(bytes))]))
  require(docs[BODY] && docs[TYPES] && docs[SETTINGS], "invalid_package")
  return { parts, docs }
}
type Package = ReturnType<typeof unpack>
function reasonText(value: string) {
  require(typeof value === "string" && value.trim() && value.length <= 4000 && value.isWellFormed() && !/[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/.test(value), "invalid_reason")
  return value
}
function requestCopy(value: DocxTableRowInsertion) {
  return { ...value, cells: [...value.cells], author: value.author.trim() }
}
function info(pkg: Package) {
  const rels = pkg.docs[RELS] ?? parse(`<Relationships xmlns="${R}"/>`)
  const links = all(rels, R, "Relationship").filter(el => el.getAttribute("Type") === COMMENT_REL)
  require(links.length <= 1 && links.every(el => !el.getAttribute("TargetMode") || el.getAttribute("TargetMode") === "Internal"), "invalid_comment_relationship")
  const target = links[0]?.getAttribute("Target")
  require(!links.length || target, "invalid_comment_relationship")
  const part = target ? path.posix.resolve("/word", target).slice(1) : "word/comments.xml"
  require(links.length ? pkg.docs[part] : !pkg.parts[part], "invalid_comment_part")
  const comments = links.length ? pkg.docs[part]! : parse(`<w:comments xmlns:w="${W}"/>`)
  require(comments.documentElement?.namespaceURI === W && comments.documentElement.localName === "comments", "invalid_comment_part")
  const overrides = all(pkg.docs[TYPES]!, CT, "Override").filter(el => el.getAttribute("PartName") === `/${part}`)
  require(links.length ? overrides.length === 1 && overrides[0]!.getAttribute("ContentType") === COMMENT_TYPE : !overrides.length, "invalid_comment_content_type")
  const ids = all(comments, W, "comment").map(el => el.getAttributeNS(W, "id"))
  require(ids.length === new Set(ids.map(id => numericId(id))).size, "ambiguous_comment_id")
  return { rels, links, comments, part }
}
function numericId(id: string | null) { require(typeof id === "string" && /^[+-]?\d+$/.test(id), "invalid_comment_id"); return BigInt(id).toString() }
function create(doc: Document, ns: string, name: string, attrs: Record<string, string> = {}) {
  const el = doc.createElementNS(ns, name)
  for (const [key, value] of Object.entries(attrs)) el.setAttributeNS(key.startsWith("w:") ? W : key.startsWith("xml:") ? XML : null, key, value)
  return el
}
function expectedComment(doc: Document, id: string, author: string, reason: string) {
  const comment = create(doc, W, "w:comment", { "w:id": id, "w:author": author })
  for (const line of reason.split(/\r\n|\r|\n/)) {
    const p = create(doc, W, "w:p"), r = create(doc, W, "w:r"), t = create(doc, W, "w:t", { "xml:space": "preserve" })
    t.appendChild(doc.createTextNode(line)); r.appendChild(t); p.appendChild(r); comment.appendChild(p)
  }
  return comment
}
// Independently check every annotation/dependency delta BEFORE restoring any
// source part. With `present=false`, native RemoveComment has removed the exact
// owned entry/triple but deliberately retains its now-empty part/dependencies.
function peel(source: Package, bytes: Uint8Array, author: string, reason: string, removedId?: string) {
  const output = unpack(bytes), before = info(source), body = output.docs[BODY]!
  const rels = output.docs[RELS], comments = output.docs[before.part]
  require(rels && comments, "missing_review_comment")
  const oldIds = new Set(all(before.comments, W, "comment").map(el => numericId(el.getAttributeNS(W, "id"))))
  const added = all(comments, W, "comment").filter(el => !oldIds.has(numericId(el.getAttributeNS(W, "id"))))
  require(removedId ? added.length === 0 : added.length === 1, "unexpected_review_comments")
  const id = removedId ?? added[0]!.getAttributeNS(W, "id")
  require(typeof id === "string" && /^(0|[1-9]\d*)$/.test(id), "invalid_owned_comment_id")
  for (const doc of Object.values(source.docs)) for (const name of ["comment", "commentRangeStart", "commentRangeEnd", "commentReference"])
    require(all(doc, W, name).every(el => numericId(el.getAttributeNS(W, "id")) !== id), "owned_comment_id_collision")
  const expectedComments = before.comments.cloneNode(true) as Document
  if (!removedId) expectedComments.documentElement!.appendChild(expectedComment(expectedComments, id, author, reason))
  require(same(comments, parse(xml(expectedComments))), "comment_content_or_prior_comments_changed")
  const expectedRels = before.rels.cloneNode(true) as Document, expectedTypes = source.docs[TYPES]!.cloneNode(true) as Document
  if (!before.links.length) {
    const candidates = all(rels, R, "Relationship").filter(el => el.getAttribute("Type") === COMMENT_REL)
    require(candidates.length === 1, "unexpected_comment_relationships")
    const relId = candidates[0]!.getAttribute("Id")
    require(relId && !all(before.rels, R, "Relationship").some(el => el.getAttribute("Id") === relId), "comment_relationship_collision")
    expectedRels.documentElement!.appendChild(create(expectedRels, R, "Relationship", { Id: relId, Type: COMMENT_REL, Target: "comments.xml" }))
    expectedTypes.documentElement!.appendChild(create(expectedTypes, CT, "Override", { PartName: `/${before.part}`, ContentType: COMMENT_TYPE }))
  }
  require(same(rels, parse(xml(expectedRels))) && same(output.docs[TYPES]!, parse(xml(expectedTypes))), "comment_dependencies_changed")
  const names = new Set([...Object.keys(source.parts), before.part, RELS])
  require(JSON.stringify(Object.keys(output.parts).sort()) === JSON.stringify([...names].sort()), "comment_package_inventory_changed")
  for (const [name, part] of Object.entries(source.parts)) if (![BODY, SETTINGS, before.part, RELS, TYPES].includes(name))
    require(source.docs[name] ? same(source.docs[name]!, output.docs[name]!) : hash(part) === hash(output.parts[name]!), `unrelated_part_changed:${name}`)
  const markers = Object.entries(output.docs).flatMap(([part, doc]) => ["commentRangeStart", "commentRangeEnd", "commentReference"].flatMap(name =>
    all(doc, W, name).filter(el => numericId(el.getAttributeNS(W, "id")) === id).map(el => ({ part, name, el }))))
  let paragraphIndex: number | undefined
  if (removedId) require(markers.length === 0, "owned_comment_markers_remain")
  else {
    require(markers.length === 3 && markers.every(mark => mark.part === BODY), "owned_comment_marker_extent_changed")
    const get = (name: string) => { const found = markers.filter(mark => mark.name === name); require(found.length === 1, "owned_comment_marker_extent_changed"); return found[0]!.el }
    const start = get("commentRangeStart"), end = get("commentRangeEnd"), reference = get("commentReference"), paragraph = start.parentNode as Element, run = reference.parentNode as Element
    require(paragraph?.namespaceURI === W && paragraph.localName === "p" && end.parentNode === paragraph && run.parentNode === paragraph && run.namespaceURI === W && run.localName === "r" && !run.attributes.length && run.childNodes.length === 1, "owned_comment_marker_extent_changed")
    for (const marker of [start, end, reference]) require(!marker.childNodes.length && marker.attributes.length === 1 && marker.attributes[0]!.namespaceURI === W && marker.attributes[0]!.localName === "id" && marker.getAttributeNS(W, "id") === id, "owned_comment_marker_changed")
    const nodes = Array.from(paragraph.childNodes), offset = kids(paragraph)[0]?.localName === "pPr" ? 1 : 0
    require(nodes[offset] === start && nodes.at(-2) === end && nodes.at(-1) === run, "owned_comment_range_changed")
    paragraphIndex = all(body, W, "p").indexOf(paragraph)
    require(paragraphIndex >= 0, "owned_comment_paragraph_unavailable")
    paragraph.removeChild(start); paragraph.removeChild(end); paragraph.removeChild(run)
  }
  const core: Record<string, Uint8Array> = { ...output.parts, [BODY]: strToU8(xml(body)) }
  // SDK Save may add an XML declaration or reserialize empty tags. Only after
  // exact XML-tree/binary equality above, retain the original untouched bytes.
  for (const [name, part] of Object.entries(source.parts)) if (![BODY, SETTINGS, before.part, RELS, TYPES].includes(name)) core[name] = part
  for (const name of [before.part, RELS, TYPES]) {
    if (source.parts[name]) core[name] = source.parts[name]!
    else delete core[name]
  }
  return { core: zipSync(core, { mtime: new Date(1980, 0, 1) }), commentId: id, paragraphIndex, commentPart: before.part }
}
function anchor(bytes: Uint8Array, paragraph: { xml_paragraph_index: number; text: string }) {
  const found = readDocxCommentParagraphPositions(bytes).filter(row => row.xml_paragraph_indexes.length === 1 && row.xml_paragraph_indexes[0] === paragraph.xml_paragraph_index && row.paragraph_text === paragraph.text)
  require(found.length === 1 && paragraph.text.length, "result_paragraph_unavailable")
  return { paragraph_index: found[0]!.paragraph_index, paragraph_text: found[0]!.paragraph_text }
}
function stable(record: Revision, keepId = true) {
  const { anchorId: _anchor, affectedAnchors: _affected, id, ...rest } = record
  return JSON.stringify({ ...(keepId ? { id } : {}), ...rest })
}
async function readRegistry(bytes: Uint8Array) {
  const s = (await docxodus()).openDocxSession(bytes, { persistAnchorIds: false })
  try { return s.listRevisions() } finally { s.close() }
}
async function inspect(source: Uint8Array, bytes: Uint8Array, input: DocxTableRowInsertion, reason: string) {
  const peeled = peel(unpack(source), bytes, input.author, reason), verified = await verifyDocxTableRowInsertion(source, peeled.core, input)
  require(verified.revisions.every(member => member.author === input.author), "review_author_changed")
  const paragraph = verified.paragraphs.find(p => p.text.length > 0)
  require(paragraph && paragraph.xml_paragraph_index === peeled.paragraphIndex, "reason_not_on_first_nonempty_result_cell")
  const [coreRegistry, finalRegistry] = await Promise.all([readRegistry(peeled.core), readRegistry(bytes)])
  require(JSON.stringify(coreRegistry.map(r => stable(r)).sort()) === JSON.stringify(finalRegistry.map(r => stable(r)).sort()), "comment_changed_revision_membership")
  return { ...peeled, verified, paragraph }
}
function digest(session: Session) {
  const d = session.getPackageManifest().rawPackageBytesDigest
  require(d?.algorithm === "SHA-256" && typeof d.value === "string" && /^[a-f0-9]{64}$/.test(d.value), "package_digest_unavailable")
  return d.value
}
function trackingBase(doc: Document) {
  const copy = doc.cloneNode(true) as Document, marks = all(copy, W, "trackRevisions")
  require(marks.length <= 1 && marks.every(mark => mark.parentNode === copy.documentElement && !mark.childNodes.length && Array.from(mark.attributes).every(a => a.namespaceURI === W && a.localName === "val" && ["1", "true", "on"].includes(a.value))), "unexpected_tracking_setting")
  marks.forEach(mark => mark.parentNode!.removeChild(mark)); return copy
}
function comparePackages(actual: Uint8Array, expected: Uint8Array) {
  const a = unpack(actual), b = unpack(expected)
  require(JSON.stringify(Object.keys(a.parts).sort()) === JSON.stringify(Object.keys(b.parts).sort()), "resolved_package_inventory_changed")
  for (const name of Object.keys(a.parts)) require(a.docs[name] ? same(a.docs[name]!, b.docs[name]!) : hash(a.parts[name]!) === hash(b.parts[name]!), `resolved_package_changed:${name}`)
}
// Reuse only the audited annotation/dependency proof. Deleted-row review has a
// distinct explicit anchor contract and does not relax insertion's mapping.
export const validateDocxRowReviewReason = reasonText
export const assertDocxRowReviewPackage = comparePackages
// This proves annotation data only. The caller must independently establish
// that this exact physical paragraph is its verified formatting target.
export function peelAcceptedParagraphReviewComment(source: Uint8Array, bytes: Uint8Array, author: string, reason: string,
  anchor: { kind: "accepted"; xml_paragraph_index: number; text: string }) {
  require(anchor.kind === "accepted" && Number.isSafeInteger(anchor.xml_paragraph_index) && anchor.xml_paragraph_index >= 0 && typeof anchor.text === "string" && anchor.text.length > 0, "invalid_accepted_anchor")
  const result = peel(unpack(source), bytes, author, reason)
  require(result.paragraphIndex === anchor.xml_paragraph_index, "accepted_comment_extent_changed")
  const positions = readDocxCommentParagraphPositions(bytes).filter(row => row.xml_paragraph_indexes.length === 1 && row.xml_paragraph_indexes[0] === anchor.xml_paragraph_index && row.paragraph_text === anchor.text)
  require(positions.length === 1, "accepted_comment_text_changed")
  const comments = readDocxComments(bytes).filter(comment => comment.comment_id === result.commentId)
  require(comments.length === 1 && comments[0]!.markers.length === 3 && comments[0]!.markers.every(marker => marker.xml_paragraph_index === anchor.xml_paragraph_index && marker.paragraph_index === positions[0]!.paragraph_index && marker.accepted_paragraph_text === anchor.text && marker.deleted_paragraph_text === ""), "accepted_comment_text_changed")
  return result
}
export function peelDeletedRowReviewComment(source: Uint8Array, bytes: Uint8Array, author: string, reason: string,
  anchor: { kind: "deleted"; xml_paragraph_index: number; text: string }) {
  require(anchor.kind === "deleted" && Number.isSafeInteger(anchor.xml_paragraph_index) && anchor.xml_paragraph_index >= 0 && typeof anchor.text === "string" && anchor.text.length > 0, "invalid_deleted_anchor")
  const result = peel(unpack(source), bytes, author, reason)
  require(result.paragraphIndex === anchor.xml_paragraph_index, "deleted_comment_extent_changed")
  const comment = readDocxComments(bytes).filter(comment => comment.comment_id === result.commentId)
  require(comment.length === 1 && comment[0]!.markers.length === 3 && comment[0]!.markers.every(marker => marker.xml_paragraph_index === anchor.xml_paragraph_index && !marker.accepted_paragraph_text.trim() && marker.deleted_paragraph_text === anchor.text), "deleted_comment_text_changed")
  return result
}
export function peelRemovedRowReviewComment(source: Uint8Array, bytes: Uint8Array, author: string, reason: string, ownedId: string) {
  return peel(unpack(source), bytes, author, reason, ownedId)
}
/** Review this one row and its reason as an application-controlled atomic group.
 * Reject explicitly removes this reason; native Word's own row rejection may
 * instead collapse its comment onto a surviving paragraph. */
export async function resolveDocxTableRowReview(sourceBytes: Uint8Array, artifact: Uint8Array, value: DocxTableRowInsertion, storedReason: string, decision: "accept" | "reject") {
  require(decision === "accept" || decision === "reject", "invalid_decision")
  const source = new Uint8Array(sourceBytes), bytes = new Uint8Array(artifact), input = requestCopy(value), reason = reasonText(storedReason)
  const checked = await inspect(source, bytes, input, reason), expectedCore = await resolveDocxTableRowInsertion(source, checked.core, input, decision)
  let expectedBytes = source
  if (decision === "accept") {
    const expected = exportDocxComments(expectedCore, [{ id: "row-review", text: reason, author: input.author, anchor: anchor(expectedCore, checked.paragraph) }])
    require(expected.unresolved.length === 0 && expected.attached.length === 1, "accepted_comment_unresolved")
    expectedBytes = expected.bytes
  }
  const session = (await docxodus()).openDocxSession(bytes, { persistAnchorIds: false })
  try {
    const before = digest(session), version = session.getVersion()
    const order = checked.verified.revisions.toSorted((a, b) => Number(a.family === "row_insert") - Number(b.family === "row_insert"))
    const steps = order.map(member => ({ tool: "table", action: decision, mutation: () => {
      const found = session.listRevisions().filter(row => row.partUri === member.partUri && JSON.stringify([...row.constituentKeys].sort()) === JSON.stringify([...member.constituentKeys].sort()))
      require(found.length === 1 && found[0]!.family === member.family && found[0]!.author === member.author && found[0]!.date === member.date && found[0]!.dateUtc === member.dateUtc && found[0]!.resolutionStatus === "supported", "review_group_member_changed")
      return decision === "accept" ? session.acceptRevision(found[0]!.id) : session.rejectRevision(found[0]!.id)
    } }))
    if (decision === "reject") steps.unshift({ tool: "comment", action: "reject", mutation: () => {
      const found = session.listComments().filter(comment => String(comment.id) === checked.commentId)
      require(found.length === 1, "owned_comment_unavailable")
      return session.removeComment(found[0]!.anchorId)
    } })
    const result = session.executeBatch(steps, "atomic")
    if (!result.success) {
      require(version === session.getVersion() && before === digest(session), "review_resolution_not_rolled_back")
      throw new Error("docx_row_review:review_resolution_failed_rolled_back")
    }
    const saved = session.save(), parts = unpack(saved), original = unpack(source)
    require(same(trackingBase(parts.docs[SETTINGS]!), trackingBase(unpack(checked.core).docs[SETTINGS]!)), "resolved_settings_changed")
    parts.parts[SETTINGS] = original.parts[SETTINGS]!
    const normalized = zipSync(parts.parts, { mtime: new Date(1980, 0, 1) })
    if (decision === "reject") {
      const removed = peel(original, normalized, input.author, reason, checked.commentId)
      comparePackages(removed.core, source)
      return new Uint8Array(source)
    }
    const accepted = peel(original, normalized, input.author, reason)
    require(accepted.commentId === checked.commentId && accepted.paragraphIndex === checked.paragraph.xml_paragraph_index, "accepted_comment_extent_changed")
    comparePackages(accepted.core, expectedCore); comparePackages(normalized, expectedBytes)
    require(JSON.stringify((await readRegistry(normalized)).map(r => stable(r)).sort()) === JSON.stringify((await readRegistry(source)).map(r => stable(r)).sort()), "resolved_prior_revisions_changed")
    // Restore exact bytes only for parts already proved unchanged; keep the
    // actual native-resolved main/comment payloads in the returned artifact.
    for (const name of Object.keys(original.parts)) if (![BODY, accepted.commentPart].includes(name) &&
      (original.docs[name] ? same(parts.docs[name]!, original.docs[name]!) : hash(parts.parts[name]!) === hash(original.parts[name]!))) parts.parts[name] = original.parts[name]!
    const final = zipSync(parts.parts, { mtime: new Date(1980, 0, 1) })
    comparePackages(final, expectedBytes)
    return final
  } finally { session.close() }
}
/** Add a stored reason only after core authoring succeeds, then prove both
 * grouped decisions against the final exported, commented package. */
export async function authorDocxTableRowReview(sourceBytes: Uint8Array, value: DocxTableRowInsertion, storedReason: string) {
  const source = new Uint8Array(sourceBytes), input = requestCopy(value), reason = reasonText(storedReason)
  const core = await authorDocxTableRowInsertion(source, input), verified = await verifyDocxTableRowInsertion(source, core.bytes, input), paragraph = verified.paragraphs.find(p => p.text.length > 0)
  require(paragraph, "result_paragraph_unavailable")
  const output = exportDocxComments(core.bytes, [{ id: "row-review", text: reason, author: input.author, anchor: anchor(core.bytes, paragraph) }])
  require(output.unresolved.length === 0 && output.attached.length === 1, "reason_comment_unresolved")
  const checked = await inspect(source, output.bytes, input, reason)
  await resolveDocxTableRowReview(source, output.bytes, input, reason, "reject")
  await resolveDocxTableRowReview(source, output.bytes, input, reason, "accept")
  return { bytes: output.bytes, revisions: checked.verified.revisions.map(({ id, family }) => ({ id, family })), commentId: checked.commentId }
}
