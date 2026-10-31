import { authorWordChatTarget } from "../../../dochaus/lib/docx-flat-opc"
import type { WordChatPreparedAmendment, WordChatProposal, WordChatSnapshot, WordChatTableRowInsertionProposal, WordChatTableRowDeletionProposal } from "../../../dochaus/lib/word-chat-amendment-contract"
import { targetWordChatAmendment, validWordChatFormatting, validWordChatInlineFormatting, validWordChatParagraphInsertion, validWordChatParagraphDeletion, validWordChatListInsertion, wordChatListCaptureCurrent, validWordChatTableRowInsertion, validWordChatTableRowDeletion } from "../../../dochaus/lib/word-chat-amendment-contract"
import { authorDocxTableRowInsertion, docxTableRowInsertionTargets, type DocxTableRowInsertion } from "../../../dochaus/lib/docx-table-row-insertion"
import { authorDocxTableRowDeletion, docxTableRowDeletionTargets, type DocxTableRowDeletion } from "../../../dochaus/lib/docx-table-row-deletion"
import { parseZip } from "../../../dochaus/lib/zip"

export { MAX_WORD_AMENDMENT_OOXML_BYTES, WordAmendmentError, WordFormattingError } from "../../../dochaus/lib/docx-flat-opc"

/** Pure preparation against the reviewed target. Inline properties become a
 * native Word command; wording and paragraph properties use shared authoring.
 * The route must
 * resolve its stored proposal and current editing lease before calling this.
 * Neither this function nor the shared Docxodus writer publishes a version. */
export async function authorWordChatAmendment(ooxml: string, proposal: WordChatProposal, context?: WordChatSnapshot): Promise<WordChatPreparedAmendment> {
  if (proposal.paragraphAction === "append_list_item") {
    const pinnedContext = context && { ...context, selection:proposal.target }
    const canonical = pinnedContext && targetWordChatAmendment(pinnedContext, proposal.before)
    if (!context || !canonical || !proposal.id || proposal.before !== proposal.target.text || !wordChatListCaptureCurrent(context, proposal.list) ||
      JSON.stringify(context.list) !== JSON.stringify(proposal.list) ||
      canonical.paragraphs[0]?.id !== proposal.target.paragraphs[0]?.id || canonical.startOffset !== proposal.target.startOffset || canonical.endOffset !== proposal.target.endOffset ||
      !validWordChatListInsertion(pinnedContext!, canonical, proposal.after, proposal.paragraphAction, proposal.formatting, proposal.inlineFormatting)) throw new Error("word_proposal_changed")
    return {kind:"list_insert", proposalId:proposal.id, target:structuredClone(proposal.target), text:proposal.after.slice(proposal.before.length + 1), list:structuredClone(proposal.list)}
  }
  if (proposal.list !== undefined) throw new Error("word_proposal_changed")
  if (proposal.paragraphAction !== undefined) {
    if (!proposal.id || proposal.before !== proposal.target.text) throw new Error("word_proposal_changed")
    const target = proposal.target, snapshot = {text:target.paragraphs.map(p => p.text).join("\r"), paragraphs:target.paragraphs,
      selection:target, capturedAt:target.capturedAt, saved:target.saved, truncated:false}
    const canonical = targetWordChatAmendment(snapshot,target.text)
    if (!canonical || canonical.startOffset !== target.startOffset || canonical.endOffset !== target.endOffset ||
      !validWordChatParagraphDeletion(snapshot,canonical,proposal.after,proposal.paragraphAction,proposal.formatting,proposal.inlineFormatting)) throw new Error("word_proposal_changed")
    return {kind:"paragraph_delete",proposalId:proposal.id,target:structuredClone(target)}
  }
  if (proposal.inlineFormatting === undefined && proposal.id && proposal.before === proposal.target.text) {
    const target = proposal.target, snapshot = {text:target.paragraphs.map(p => p.text).join("\r"), paragraphs:target.paragraphs,
      selection:target, capturedAt:target.capturedAt, saved:target.saved, truncated:false}
    const canonical = targetWordChatAmendment(snapshot,target.text)
    if (canonical && canonical.startOffset === target.startOffset && canonical.endOffset === target.endOffset &&
      validWordChatParagraphInsertion(snapshot,canonical,proposal.after,proposal.formatting))
      return {kind:"paragraph_insert",proposalId:proposal.id,target:structuredClone(target),text:proposal.after.slice(proposal.before.length + 1), ...(proposal.formatting === undefined ? {} : {formatting:structuredClone(proposal.formatting)})}
  }
  if (proposal.inlineFormatting !== undefined) {
    if (!proposal.id || proposal.before !== proposal.target.text || proposal.after !== proposal.before || proposal.formatting !== undefined)
      throw new Error("word_proposal_changed")
    const target = proposal.target, snapshot = {text:target.paragraphs.map(p => p.text).join("\r"), paragraphs:target.paragraphs,
      selection:target, capturedAt:target.capturedAt, saved:target.saved, truncated:false}
    const canonical = targetWordChatAmendment(snapshot, target.text)
    if (!canonical || !validWordChatInlineFormatting(snapshot, canonical, proposal.inlineFormatting)) throw new Error("word_proposal_changed")
    return { kind:"inline_format", proposalId:proposal.id, target:structuredClone(target), inlineFormatting:{...proposal.inlineFormatting} }
  }
  if (!proposal.id || proposal.before !== proposal.target.text || (proposal.formatting !== undefined
    ? !validWordChatFormatting(proposal.target, proposal.formatting, proposal.after)
    : proposal.before === proposal.after))
    throw new Error("word_proposal_changed")
  const authored = await authorWordChatTarget(ooxml, proposal.target, proposal.after, proposal.formatting)
  return { ...authored, proposalId: proposal.id }
}

const MAX_ROW_SOURCE_BYTES = 16 * 1024 * 1024
const MAX_ROW_EXPANDED_BYTES = 64 * 1024 * 1024
const MAX_ROW_PART_BYTES = 16 * 1024 * 1024

/** Bound decompression before the existing writer opens the package. Filtered
 * ZIP entries are refused here because the writer also reads those entries. */
function checkedWordRowPackage(source: Uint8Array) {
  if (!source.length || source.length > MAX_ROW_SOURCE_BYTES) throw new Error("word_row_source_size")
  const zip = parseZip(source), view = new DataView(source.buffer, source.byteOffset, source.byteLength)
  let end = -1
  for (let at = source.length - 22; at >= Math.max(0, source.length - 22 - 0xffff); at--) {
    if (view.getUint32(at, true) === 0x06054b50) { end = at; break }
  }
  if (end < 0 || end + 22 + view.getUint16(end + 20, true) !== source.length ||
    end >= 20 && view.getUint32(end - 20, true) === 0x07064b50 ||
    view.getUint16(end + 4, true) !== 0 || view.getUint16(end + 6, true) !== 0 ||
    view.getUint16(end + 8, true) !== view.getUint16(end + 10, true) ||
    view.getUint16(end + 10, true) !== zip.entries.length || zip.entries.length > 4096 ||
    !Number.isSafeInteger(zip.uncompressed) || zip.uncompressed > MAX_ROW_EXPANDED_BYTES ||
    zip.entries.some(entry => !Number.isSafeInteger(entry.size) || entry.size > MAX_ROW_PART_BYTES || entry.encrypted || entry.unsafePath) ||
    new Set(zip.entries.map(entry => entry.filename)).size !== zip.entries.length) throw new Error("word_row_source_archive")
  for (const entry of zip.entries) entry.content()
}

export type WordChatAuthoredTableRow = {
  kind: "table_row_insert"
  proposalId: string
  sourceSha256: string
  table: WordChatTableRowInsertionProposal["table"]
  cells: string[]
  author: "doc.haus"
  writerRequest: DocxTableRowInsertion
  /** Verification oracle only. Never import these bytes into a live Word document. */
  authoredBytes: Uint8Array
  revisions: { id: string; family: string }[]
}

/** Pure authoring against the exact current Word export. The future route must
 * provide server-owned reviewed context and proposal, and recheck its lease. */
export async function authorWordChatTableRow(snapshot: WordChatSnapshot, proposal: WordChatTableRowInsertionProposal, currentBytes: Uint8Array): Promise<WordChatAuthoredTableRow> {
  if (!validWordChatTableRowInsertion(snapshot, proposal)) throw new Error("word_row_proposal_changed")
  if (!(currentBytes instanceof Uint8Array) || !currentBytes.length || currentBytes.length > MAX_ROW_SOURCE_BYTES) throw new Error("word_row_source_size")
  const captured = {tableIndex:proposal.table.tableIndex,rows:proposal.table.rows.map(row => row.map(cell => ({...cell})))}
  const cells = [...proposal.cells], proposalId = proposal.id
  const source = new Uint8Array(currentBytes)
  checkedWordRowPackage(source)
  const packet = await docxTableRowInsertionTargets(source)
  const rows = captured.rows, table = packet.targets.filter(target => target.table === captured.tableIndex + 1)
  if (table.length !== rows.length || rows.some((cells, index) => {
    const matches = table.filter(target => target.row === index + 1)
    return matches.length !== 1 || matches[0]!.columns !== cells.length || matches[0]!.cells.length !== cells.length ||
      matches[0]!.cells.some((text, col) => text !== cells[col]!.text)
  })) throw new Error("word_row_source_changed")
  const target = table.find(item => item.row === rows.length)!
  const writerRequest: DocxTableRowInsertion = {source_sha256:packet.source_sha256,target_handle:target.target_handle,position:"after",cells:[...cells],author:"doc.haus"}
  const authored = await authorDocxTableRowInsertion(source, writerRequest)
  return {kind:"table_row_insert", proposalId, sourceSha256:packet.source_sha256,
    table:captured, cells,
    author:"doc.haus", writerRequest, authoredBytes:authored.bytes, revisions:authored.revisions}
}

/** Prepare a source-bound deletion oracle, retaining the entire captured table. */
export async function authorWordChatTableRowDeletion(snapshot: WordChatSnapshot, proposal: WordChatTableRowDeletionProposal, currentBytes: Uint8Array) {
  if (!validWordChatTableRowDeletion(snapshot, proposal)) throw new Error("word_row_proposal_changed")
  if (!(currentBytes instanceof Uint8Array) || !currentBytes.length || currentBytes.length > MAX_ROW_SOURCE_BYTES) throw new Error("word_row_source_size")
  const source = new Uint8Array(currentBytes), table = structuredClone(proposal.table), cells = [...proposal.cells]
  checkedWordRowPackage(source)
  const discovered = await docxTableRowInsertionTargets(source), rows = discovered.targets.filter(t => t.table === table.tableIndex + 1)
  if (rows.length !== table.rows.length || table.rows.some((row, index) => {
    const found = rows.filter(t => t.row === index + 1)
    return found.length !== 1 || found[0]!.columns !== row.length || found[0]!.cells.length !== row.length || found[0]!.cells.some((value, col) => value !== row[col]!.text)
  })) throw new Error("word_row_source_changed")
  const eligible = await docxTableRowDeletionTargets(source), target = eligible.targets.filter(t => t.table === table.tableIndex + 1 && t.row === proposal.rowIndex + 1)
  if (eligible.source_sha256 !== discovered.source_sha256 || target.length !== 1) throw new Error("word_row_target_unsupported")
  const writerRequest: DocxTableRowDeletion = {source_sha256:eligible.source_sha256,target_handle:target[0]!.target_handle,author:"doc.haus"}
  const authored = await authorDocxTableRowDeletion(source, writerRequest)
  return {kind:"table_row_delete" as const, proposalId:proposal.id, sourceSha256:eligible.source_sha256, table, rowIndex:proposal.rowIndex, cells, author:"doc.haus" as const, writerRequest, authoredBytes:authored.bytes, revisions:authored.revisions}
}
