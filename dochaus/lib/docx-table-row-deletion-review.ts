import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, unzipSync, zipSync } from "fflate"
import { docxodus } from "./docxodus"
import { exportDocxComments } from "./docx-comments"
import { authorDocxTableRowDeletion, resolveDocxTableRowDeletion, verifyDocxTableRowDeletion, type DocxTableRowDeletion } from "./docx-table-row-deletion"
import { assertDocxRowReviewPackage, peelDeletedRowReviewComment, peelRemovedRowReviewComment, validateDocxRowReviewReason } from "./docx-table-row-review"

// Offline application-controlled review only. The exported explanation remains
// a native comment; this does not establish Microsoft Word decision semantics.
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", SETTINGS = "word/settings.xml"
type Session = ReturnType<Awaited<ReturnType<typeof docxodus>>["openDocxSession"]>
type Revision = ReturnType<Session["listRevisions"]>[number]
function require(value: unknown, reason: string): asserts value { if (!value) throw new Error(`docx_row_deletion_review:${reason}`) }
function normalized(input: DocxTableRowDeletion) { return { ...input, author: input.author.trim() } }
function native(record: Revision) {
  const { id: _publicId, anchorId: _anchor, affectedAnchors: _affected, ...rest } = record
  return JSON.stringify(rest)
}
function identity(record: Revision) {
  const { anchorId: _anchor, affectedAnchors: _affected, ...rest } = record
  return JSON.stringify(rest)
}
function extent(record: Revision) { return JSON.stringify({ anchor: record.anchorId, affected: record.affectedAnchors }) }
async function registry(bytes: Uint8Array) {
  const session = (await docxodus()).openDocxSession(bytes, { persistAnchorIds: false })
  try { return session.listRevisions() } finally { session.close() }
}
function member(records: Revision[], wanted: Revision) {
  const found = records.filter(record => record.partUri === wanted.partUri && JSON.stringify(record.constituentKeys.toSorted()) === JSON.stringify(wanted.constituentKeys.toSorted()))
  require(found.length === 1 && native(found[0]!) === native(wanted), "native_member_changed")
  return found[0]!
}
async function inspect(source: Uint8Array, bytes: Uint8Array, input: DocxTableRowDeletion, reason: string) {
  // Core verification derives the target, so an alleged comment anchor supplied
  // by a caller cannot choose a paragraph. Peel first checks only exact added
  // annotation data; the full core proof then establishes its authority.
  const unanchored = peelDeletedCandidate(source, bytes, input.author, reason)
  const verified = await verifyDocxTableRowDeletion(source, unanchored.core, input), paragraph = verified.paragraphs.find(p => p.text.length > 0)
  require(paragraph && verified.revisions.every(r => r.author === input.author), "deleted_paragraph_unavailable")
  const checked = peelDeletedRowReviewComment(source, bytes, input.author, reason, { kind: "deleted", ...paragraph })
  const [core, final] = await Promise.all([registry(checked.core), registry(bytes)])
  require(JSON.stringify(core.map(identity).sort()) === JSON.stringify(final.map(identity).sort()), "comment_changed_native_membership")
  return { ...checked, verified, paragraph }
}
// Derive the sole owned comment's physical location from read-only native XML,
// before trusting it: the anchored adapter validates its exact deleted text and
// the core proof below validates that it is the reviewed row's paragraph.
function peelDeletedCandidate(source: Uint8Array, bytes: Uint8Array, author: string, reason: string) {
  const output = unzipSync(bytes), sourceParts = unzipSync(source)
  const parse = (part: Uint8Array) => {
    const text = strFromU8(part), fail = () => { throw new Error("docx_row_deletion_review:invalid_xml") }
    require(!/<!DOCTYPE|<!ENTITY/i.test(text), "invalid_xml")
    return new DOMParser({ errorHandler: { warning: fail, error: fail, fatalError: fail } }).parseFromString(text, "application/xml")
  }
  require(output["word/document.xml"] && sourceParts["word/document.xml"], "invalid_package")
  const doc = parse(output["word/document.xml"]!), old = parse(sourceParts["word/document.xml"]!), paragraphs = Array.from(doc.getElementsByTagNameNS(W, "p"))
  const oldIds = new Set(Array.from(old.getElementsByTagNameNS(W, "commentRangeStart")).map(marker => marker.getAttributeNS(W, "id")))
  const added = Array.from(doc.getElementsByTagNameNS(W, "commentRangeStart")).filter(marker => !oldIds.has(marker.getAttributeNS(W, "id")))
  require(added.length === 1, "owned_comment_unavailable")
  const paragraph = added[0]!.parentNode as Element, index = paragraphs.indexOf(paragraph)
  require(index >= 0, "deleted_comment_extent_changed")
  const texts = Array.from(paragraph.getElementsByTagNameNS(W, "delText")).map(text => text.textContent ?? "").join("")
  return peelDeletedRowReviewComment(source, bytes, author, reason, { kind: "deleted", xml_paragraph_index: index, text: texts })
}
function digest(session: Session) {
  const value = session.getPackageManifest().rawPackageBytesDigest
  require(value?.algorithm === "SHA-256" && typeof value.value === "string" && /^[a-f0-9]{64}$/.test(value.value), "package_digest_unavailable")
  return value.value
}
function restoreSettings(source: Uint8Array, bytes: Uint8Array, admitted: Uint8Array) {
  const original = unzipSync(source), output = unzipSync(bytes), before = unzipSync(admitted)
  const base = (value: Uint8Array) => {
    const fail = () => { throw new Error("docx_row_deletion_review:invalid_settings") }
    const text = strFromU8(value)
    require(!/<!DOCTYPE|<!ENTITY/i.test(text), "invalid_settings")
    const doc = new DOMParser({ errorHandler: { warning: fail, error: fail, fatalError: fail } }).parseFromString(text, "application/xml"), marks = Array.from(doc.getElementsByTagNameNS(W, "trackRevisions"))
    require(marks.length <= 1 && marks.every(mark => mark.parentNode === doc.documentElement && !mark.childNodes.length && Array.from(mark.attributes).every(a => a.namespaceURI === W && a.localName === "val" && ["1", "true", "on"].includes(a.value))), "settings_changed")
    marks.forEach(mark => mark.parentNode!.removeChild(mark))
    return new XMLSerializer().serializeToString(doc)
  }
  // The shared strict comparator retains all namespaces/attributes/text/PI and
  // ignores only harmless XML serialization, not arbitrary settings changes.
  const a = { ...before, [SETTINGS]: new TextEncoder().encode(base(before[SETTINGS]!)) }, b = { ...before, [SETTINGS]: new TextEncoder().encode(base(output[SETTINGS]!)) }
  assertDocxRowReviewPackage(zipSync(a), zipSync(b))
  output[SETTINGS] = original[SETTINGS]!
  return zipSync(output, { mtime: new Date(1980, 0, 1) })
}
/** Both application decisions remove only the owned explanation first. Actual
 * external native Accept may instead collapse it onto a surviving paragraph. */
export async function resolveDocxTableRowDeletionReview(sourceBytes: Uint8Array, artifact: Uint8Array, value: DocxTableRowDeletion, storedReason: string, decision: "accept" | "reject") {
  require(decision === "accept" || decision === "reject", "invalid_decision")
  const source = new Uint8Array(sourceBytes), bytes = new Uint8Array(artifact), input = normalized(value), reason = validateDocxRowReviewReason(storedReason)
  const checked = await inspect(source, bytes, input, reason), expected = await resolveDocxTableRowDeletion(source, checked.core, input, decision)
  const session = (await docxodus()).openDocxSession(bytes, { persistAnchorIds: false })
  try {
    const owned = checked.verified.revisions.map(wanted => member(session.listRevisions(), wanted)), before = digest(session), version = session.getVersion()
    const result = session.executeBatch([{ tool: "comment", action: "remove_reason", mutation: () => {
      const comments = session.listComments().filter(comment => String(comment.id) === checked.commentId)
      require(comments.length === 1, "owned_comment_unavailable")
      return session.removeComment(comments[0]!.anchorId)
    } }, ...owned.map(wanted => ({ tool: "table", action: decision, mutation: () => {
      const found = member(session.listRevisions(), wanted)
      require(extent(found) === extent(wanted), "native_extent_changed_after_comment_removal")
      return decision === "accept" ? session.acceptRevision(found.id) : session.rejectRevision(found.id)
    } }))], "atomic")
    if (!result.success) {
      require(session.getVersion() === version && digest(session) === before, "failed_review_not_rolled_back")
      throw new Error("docx_row_deletion_review:review_failed_rolled_back")
    }
    const saved = restoreSettings(source, session.save(), checked.core), removed = peelRemovedRowReviewComment(source, saved, input.author, reason, checked.commentId)
    assertDocxRowReviewPackage(removed.core, expected)
    const [actual, intended] = await Promise.all([registry(removed.core), registry(expected)])
    require(JSON.stringify(actual.map(identity).sort()) === JSON.stringify(intended.map(identity).sort()), "final_native_membership_changed")
    // Canonical acceptance is exactly the comment-free core's proven result.
    // The actual final commented package was resolved and verified above.
    return expected
  } finally { session.close() }
}
export async function authorDocxTableRowDeletionReview(sourceBytes: Uint8Array, value: DocxTableRowDeletion, storedReason: string) {
  const source = new Uint8Array(sourceBytes), input = normalized(value), reason = validateDocxRowReviewReason(storedReason)
  const core = await authorDocxTableRowDeletion(source, input), verified = await verifyDocxTableRowDeletion(source, core.bytes, input), paragraph = verified.paragraphs.find(p => p.text.length > 0)
  require(paragraph, "deleted_paragraph_unavailable")
  const output = exportDocxComments(core.bytes, [{ id: "row-deletion-review", text: reason, author: input.author,
    anchor: { deleted_paragraph_text: paragraph.text, paragraph_index: paragraph.xml_paragraph_index } }])
  require(output.unresolved.length === 0 && output.attached.length === 1 && output.attached[0]!.anchor_kind === "deleted", "deleted_reason_unresolved")
  const checked = await inspect(source, output.bytes, input, reason)
  await resolveDocxTableRowDeletionReview(source, output.bytes, input, reason, "reject")
  await resolveDocxTableRowDeletionReview(source, output.bytes, input, reason, "accept")
  return { bytes: output.bytes, revisions: checked.verified.revisions.map(({ id, family }) => ({ id, family })), commentId: checked.commentId }
}
