import { DOMParser } from "@xmldom/xmldom"
import { strFromU8, unzipSync, zipSync } from "fflate"
import { docxodus } from "./docxodus"
import { exportDocxComments, readDocxCommentParagraphPositions, readDocxComments } from "./docx-comments"
import { assertExactSpacingPackage, authorExactParagraphSpacing, resolveExactParagraphSpacing, verifyExactParagraphSpacing, type ExactSpacingRequest } from "./docx-exact-spacing"
import { assertDocxRowReviewPackage, peelAcceptedParagraphReviewComment, peelRemovedRowReviewComment, validateDocxRowReviewReason } from "./docx-table-row-review"

// Offline artifact review only. Application Reject removes its explanation;
// this does not establish Microsoft Word's own comment/decision behavior.
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", BODY = "word/document.xml"
const RELS = "word/_rels/document.xml.rels", TYPES = "[Content_Types].xml"
type Session = ReturnType<Awaited<ReturnType<typeof docxodus>>["openDocxSession"]>
type Revision = ReturnType<Session["listRevisions"]>[number]
function require(value: unknown, reason: string): asserts value { if (!value) throw new Error(`docx_exact_spacing_review:${reason}`) }
function copy(value: ExactSpacingRequest): ExactSpacingRequest {
  require(typeof value.author === "string" && value.author.trim().length > 0 && value.author.trim().length <= 100 && value.author.isWellFormed() && !/[\u0000-\u001f\u007f]/.test(value.author), "invalid_author")
  return { ...value, author: value.author.trim(), tableContext: value.tableContext === null ? null : JSON.parse(JSON.stringify(value.tableContext)) }
}
function parts(bytes: Uint8Array) {
  assertExactSpacingPackage(bytes)
  return unzipSync(bytes)
}
const parse = (bytes: Uint8Array) => new DOMParser().parseFromString(strFromU8(bytes), "application/xml")
const all = (node: Document | Element, name: string) => Array.from(node.getElementsByTagNameNS(W, name))
const numericId = (value: string | null) => { require(typeof value === "string" && /^[+-]?\d+$/.test(value), "invalid_comment_id"); return BigInt(value).toString() }
function identity(record: Revision) {
  const { anchorId: _anchor, affectedAnchors: _affected, ...rest } = record
  return JSON.stringify(rest)
}
function native(record: Revision) {
  const { id: _id, anchorId: _anchor, affectedAnchors: _affected, ...rest } = record
  return JSON.stringify(rest)
}
async function registry(bytes: Uint8Array) {
  assertExactSpacingPackage(bytes)
  const session = (await docxodus()).openDocxSession(bytes, { persistAnchorIds: false })
  try { return session.listRevisions() } finally { session.close() }
}
function paragraph(core: Uint8Array, physicalId: string, text: string) {
  const body = parse(parts(core)[BODY]!), marks = all(body, "pPrChange").filter(mark => mark.getAttributeNS(W, "id") === physicalId)
  require(marks.length === 1, "formatting_owner_unavailable")
  const pp = marks[0]!.parentNode as Element, owner = pp.parentNode as Element
  require(pp.namespaceURI === W && pp.localName === "pPr" && owner.namespaceURI === W && owner.localName === "p", "formatting_owner_unavailable")
  const index = all(body, "p").indexOf(owner), positions = readDocxCommentParagraphPositions(core).filter(row => row.xml_paragraph_indexes.length === 1 && row.xml_paragraph_indexes[0] === index && row.paragraph_text === text)
  require(index >= 0 && text.length > 0 && positions.length === 1, "exact_paragraph_unavailable")
  return { xml_paragraph_index: index, text, anchor: { paragraph_index: positions[0]!.paragraph_index, paragraph_text: text } }
}
async function inspect(source: Uint8Array, bytes: Uint8Array, input: ExactSpacingRequest, reason: string) {
  // Validate the raw inputs before any reader/peeler can rebuild their ZIPs.
  parts(source); parts(bytes)
  const oldIds = new Set(readDocxComments(source).map(comment => numericId(comment.comment_id)))
  const added = readDocxComments(bytes).filter(comment => !oldIds.has(numericId(comment.comment_id)))
  require(added.length === 1 && added[0]!.markers.length === 3, "owned_comment_unavailable")
  const index = added[0]!.markers[0]!.xml_paragraph_index
  require(index !== null && Number.isSafeInteger(index) && index >= 0, "exact_paragraph_unavailable")
  // The claimed comment position is only a hypothesis until the peeled core
  // independently proves the formatting carrier's physical owner below.
  const checked = peelAcceptedParagraphReviewComment(source, bytes, input.author, reason, { kind: "accepted", xml_paragraph_index: index, text: input.paragraphText })
  assertExactSpacingPackage(checked.core)
  const proof = await verifyExactParagraphSpacing(source, checked.core, input), selected = paragraph(checked.core, proof.physicalRevisionId, input.paragraphText)
  require(selected.xml_paragraph_index === checked.paragraphIndex, "reason_not_on_formatting_target")
  const [core, final] = await Promise.all([registry(checked.core), registry(bytes)])
  require(JSON.stringify(core.map(identity).sort()) === JSON.stringify(final.map(identity).sort()), "comment_changed_native_membership")
  const key = `kind=PropsChange:${W}:pPrChange:${proof.physicalRevisionId}`
  const owned = final.filter(record => record.partUri === "/word/document.xml" && record.constituentKeys.length === 1 && record.constituentKeys[0] === key)
  require(owned.length === 1, "native_member_unavailable")
  return { ...checked, proof, selected, owned: owned[0]! }
}
function member(session: Session, wanted: Revision, physicalId: string) {
  const found = session.listRevisions().filter(record => record.partUri === wanted.partUri && JSON.stringify(record.constituentKeys.toSorted()) === JSON.stringify(wanted.constituentKeys.toSorted()))
  require(found.length === 1 && native(found[0]!) === native(wanted), "native_member_changed")
  const row = found[0]!
  require(typeof row.anchorId === "string", "native_extent_changed")
  const anchor = session.project().anchorIndex[row.anchorId]
  require(anchor && anchor.partUri === "/word/document.xml" && anchor.scope === "body" && (anchor.kind === "p" || anchor.kind === "tc") && row.affectedAnchors.length === 1 && row.affectedAnchors[0]!.id === row.anchorId, "native_extent_changed")
  const xml = new DOMParser().parseFromString(session.raw.getXml(row.anchorId), "application/xml")
  const marks = all(xml, "pPrChange").filter(mark => mark.getAttributeNS(W, "id") === physicalId)
  require(marks.length === 1 && marks[0]!.getAttributeNS(W, "author") === wanted.author && marks[0]!.getAttributeNS(W, "date") === wanted.date, "native_extent_changed")
  return row
}
function digest(session: Session) {
  const value = session.getPackageManifest().rawPackageBytesDigest
  require(value?.algorithm === "SHA-256" && typeof value.value === "string" && /^[a-f0-9]{64}$/.test(value.value), "package_digest_unavailable")
  return value.value
}
function restoreUnchanged(source: Uint8Array, bytes: Uint8Array, commentPart: string) {
  const original = parts(source), output = parts(bytes)
  // These parts were already proved unchanged by the annotation/core checks.
  // New comment relationships/types stay; old ones can retain source bytes.
  for (const [name, value] of Object.entries(original)) if (name !== BODY && name !== commentPart && (![RELS, TYPES].includes(name) || original[commentPart])) output[name] = value
  const restored = zipSync(output, { mtime: new Date(1980, 0, 1) })
  assertExactSpacingPackage(restored)
  assertDocxRowReviewPackage(restored, bytes)
  return restored
}
export async function verifyExactParagraphSpacingReview(sourceInput: Uint8Array, artifactInput: Uint8Array, value: ExactSpacingRequest, storedReason: string) {
  const source = new Uint8Array(sourceInput), bytes = new Uint8Array(artifactInput), input = copy(value), reason = validateDocxRowReviewReason(storedReason)
  const checked = await inspect(source, bytes, input, reason)
  return { commentId: checked.commentId, physicalRevisionId: checked.proof.physicalRevisionId, beforeDirect: checked.proof.beforeDirect, afterDirect: checked.proof.afterDirect }
}
export async function resolveExactParagraphSpacingReview(sourceInput: Uint8Array, artifactInput: Uint8Array, value: ExactSpacingRequest, storedReason: string, decision: "accept" | "reject") {
  require(decision === "accept" || decision === "reject", "invalid_decision")
  const source = new Uint8Array(sourceInput), bytes = new Uint8Array(artifactInput), input = copy(value), reason = validateDocxRowReviewReason(storedReason)
  const checked = await inspect(source, bytes, input, reason), expectedCore = await resolveExactParagraphSpacing(source, checked.core, input, decision)
  assertExactSpacingPackage(expectedCore)
  let expected = source
  if (decision === "accept") {
    const output = exportDocxComments(expectedCore, [{ id: "spacing-review", text: reason, author: input.author, anchor: checked.selected.anchor }])
    require(output.unresolved.length === 0 && output.attached.length === 1 && output.attached[0]!.comment_id === checked.commentId, "accepted_reason_unresolved")
    expected = output.bytes
  }
  assertExactSpacingPackage(expected)
  const session = (await docxodus()).openDocxSession(bytes, { persistAnchorIds: false })
  try {
    member(session, checked.owned, checked.proof.physicalRevisionId)
    session.listComments(); session.listAnnotations(); session.save()
    const before = digest(session), version = session.getVersion()
    const resolve = { tool: "paragraph", action: decision, mutation: () => {
      const found = member(session, checked.owned, checked.proof.physicalRevisionId)
      return decision === "accept" ? session.acceptRevision(found.id) : session.rejectRevision(found.id)
    } }
    const steps = decision === "accept" ? [resolve] : [{ tool: "comment", action: "remove_reason", mutation: () => {
      const comments = session.listComments().filter(comment => String(comment.id) === checked.commentId)
      require(comments.length === 1, "owned_comment_unavailable")
      return session.removeComment(comments[0]!.anchorId)
    } }, resolve]
    const result = session.executeBatch(steps, "atomic")
    if (!result.success) {
      require(session.getVersion() === version && digest(session) === before, "failed_review_not_rolled_back")
      throw new Error("docx_exact_spacing_review:review_failed_rolled_back")
    }
    const saved = session.save()
    assertExactSpacingPackage(saved)
    const peeled = decision === "accept"
      ? peelAcceptedParagraphReviewComment(source, saved, input.author, reason, { kind: "accepted", ...checked.selected })
      : peelRemovedRowReviewComment(source, saved, input.author, reason, checked.commentId)
    assertExactSpacingPackage(peeled.core)
    require(peeled.commentId === checked.commentId, "owned_comment_changed")
    assertDocxRowReviewPackage(peeled.core, expectedCore)
    if (decision === "accept") assertDocxRowReviewPackage(saved, expected)
    const [actual, intended] = await Promise.all([registry(peeled.core), registry(expectedCore)])
    require(JSON.stringify(actual.map(identity).sort()) === JSON.stringify(intended.map(identity).sort()), "final_native_membership_changed")
    if (decision === "reject") return source
    return restoreUnchanged(source, saved, checked.commentPart)
  } finally { session.close() }
}
export async function authorExactParagraphSpacingReview(sourceInput: Uint8Array, value: ExactSpacingRequest, storedReason: string) {
  const source = new Uint8Array(sourceInput), input = copy(value), reason = validateDocxRowReviewReason(storedReason)
  assertExactSpacingPackage(source)
  const core = await authorExactParagraphSpacing(source, input), proof = await verifyExactParagraphSpacing(source, core.trackedBytes, input)
  const selected = paragraph(core.trackedBytes, proof.physicalRevisionId, input.paragraphText)
  const output = exportDocxComments(core.trackedBytes, [{ id: "spacing-review", text: reason, author: input.author, anchor: selected.anchor }])
  require(output.unresolved.length === 0 && output.attached.length === 1, "reason_unresolved")
  const checked = await inspect(source, output.bytes, input, reason), bytes = restoreUnchanged(source, output.bytes, checked.commentPart)
  const verified = await verifyExactParagraphSpacingReview(source, bytes, input, reason)
  await resolveExactParagraphSpacingReview(source, bytes, input, reason, "accept")
  await resolveExactParagraphSpacingReview(source, bytes, input, reason, "reject")
  return { bytes, ...verified }
}
