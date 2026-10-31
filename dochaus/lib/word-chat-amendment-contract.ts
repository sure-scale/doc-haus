/** Current unsaved Word text is request context, never a saved source version. */
export type WordChatParagraph = { id: string; text: string }
export type WordChatTarget = {
  id: string
  text: string
  paragraphs: WordChatParagraph[]
  startOffset: number
  endOffset: number
  capturedAt: string
  saved: boolean
}
export type WordChatSnapshot = {
  text: string
  paragraphs: WordChatParagraph[]
  saved: boolean
  capturedAt: string
  truncated: boolean
  selection?: WordChatTarget
  /** Host-captured candidate geometry; the native executor must revalidate it. */
  table?: WordChatTableCapture
  list?: WordChatListCapture
}
/** Host-captured automatic list; native identity is never model-authored. */
type WordChatListMember = { id: string; text: string; level: number; label: string }
export type WordChatListCapture =
  | { id: number; format: "decimal"; members: WordChatListMember[] }
  | { id: number; format: "native"; members: Array<WordChatListMember & { siblingIndex: number }> }
export function wordChatListCaptureCurrent(snapshot: WordChatSnapshot, value: unknown): value is WordChatListCapture {
  if (!value || typeof value !== "object" || Array.isArray(value)) return false
  const list = value as WordChatListCapture
  if (Object.keys(list).some(k => !["id", "format", "members"].includes(k)) || !Number.isSafeInteger(list.id) || list.id < 0 || !(["decimal", "native"] as unknown[]).includes(list.format) ||
    !Array.isArray(list.members) || !list.members.length || list.members.length > 1000) return false
  const ids = new Set<string>()
  let totalText = 0
  for (const [index, member] of list.members.entries()) {
    if (!member || typeof member !== "object" || Array.isArray(member) || Object.keys(member).some(k => !(list.format === "native" ? ["id", "text", "level", "label", "siblingIndex"] : ["id", "text", "level", "label"]).includes(k)) ||
      typeof member.id !== "string" || !member.id || ids.has(member.id) || typeof member.text !== "string" || member.text.length > 24_000 ||
      (list.format === "decimal" ? member.level !== 0 || member.label !== `${index + 1}.` :
        !Number.isSafeInteger(member.level) || member.level < 0 || member.level > 8 || typeof member.label !== "string" || !member.label.trim() || member.label.length > 256 ||
        !("siblingIndex" in member) || typeof member.siblingIndex !== "number" || !Number.isSafeInteger(member.siblingIndex) || member.siblingIndex < 0)) return false
    ids.add(member.id)
    totalText += member.text.length
    if (totalText > 24_000) return false
  }
  const at = snapshot.paragraphs.findIndex(p => p.id === list.members[0]!.id)
  return at >= 0 && snapshot.paragraphs.filter(p => ids.has(p.id)).length === list.members.length &&
    list.members.every((m, i) => snapshot.paragraphs[at + i]?.id === m.id && snapshot.paragraphs[at + i]?.text === m.text)
}
/** One paragraph per cell in a simple, unmerged, top-level table. The table
 * index and paragraph IDs come from Word, never from the model. */
export type WordChatTableCapture = {
  tableIndex: number
  rows: WordChatParagraph[][]
}
export type WordChatTableRowInsertionProposal = {
  kind: "table_row_insert"
  id: string
  documentId: string
  explanation: string
  openItems: string[]
  createdAt: string
  position: "end"
  table: WordChatTableCapture
  cells: string[]
  /** Reviewed character formatting per cell, aligned with cells. Absent for a
   * plain row; every flag is explicit so the executor sets each segment. */
  runs?: WordChatCellRun[][]
}
export type WordChatCellRun = { text: string; bold: boolean; italic: boolean; underline: boolean }
export type WordChatTableRowDeletionProposal = {
  kind: "table_row_delete"
  id: string
  documentId: string
  explanation: string
  openItems: string[]
  createdAt: string
  table: WordChatTableCapture
  rowIndex: number
  cells: string[]
}
export type WordChatTableRowProposal = WordChatTableRowInsertionProposal | WordChatTableRowDeletionProposal
/** Additive operation type: existing paragraph proposal consumers remain unchanged. */
export type WordChatReviewedOperation =
  | { kind: "paragraph"; proposal: WordChatProposal }
  | { kind: "table_row_insert"; proposal: WordChatTableRowInsertionProposal }
  | { kind: "table_row_delete"; proposal: WordChatTableRowDeletionProposal }

/** Validate the capture's identity and ordering against the same snapshot.
 * This does not prove table geometry or grant permission to write in Word. */
export function wordChatTableCaptureCurrent(snapshot: WordChatSnapshot, value: unknown): value is WordChatTableCapture {
  if (!value || typeof value !== "object" || Array.isArray(value)) return false
  const table = value as WordChatTableCapture
  if (Object.keys(table).some(key => !["tableIndex", "rows"].includes(key)) ||
    !Number.isSafeInteger(table.tableIndex) || table.tableIndex < 0 || !Array.isArray(table.rows) ||
    !table.rows.length || table.rows.length > 1000 || !Array.isArray(table.rows[0]) ||
    !table.rows[0].length || table.rows[0].length > 32) return false
  const width = table.rows[0].length
  const cells: WordChatParagraph[] = []
  for (const row of table.rows) {
    if (!Array.isArray(row) || row.length !== width) return false
    for (const cell of row) {
      if (!cell || typeof cell !== "object" || Array.isArray(cell) || Object.keys(cell).some(key => !["id", "text"].includes(key)) ||
        typeof cell.id !== "string" || !cell.id || typeof cell.text !== "string") return false
      cells.push(cell)
    }
  }
  const ids = new Set(cells.map(cell => cell.id))
  if (ids.size !== cells.length) return false
  const at = snapshot.paragraphs.findIndex(paragraph => paragraph.id === cells[0]!.id)
  return at >= 0 && snapshot.paragraphs.filter(paragraph => ids.has(paragraph.id)).length === cells.length &&
    cells.every((cell, index) => snapshot.paragraphs[at + index]?.id === cell.id && snapshot.paragraphs[at + index]?.text === cell.text)
}

/** Only literal cells for an end insertion are admitted by this seam. */
export function validWordChatTableRowInsertion(snapshot: WordChatSnapshot, value: unknown): value is WordChatTableRowInsertionProposal {
  if (!value || typeof value !== "object" || Array.isArray(value)) return false
  const row = value as WordChatTableRowInsertionProposal
  if (Object.keys(row).some(key => !["kind", "id", "documentId", "explanation", "openItems", "createdAt", "position", "table", "cells", "runs"].includes(key)) ||
    row.kind !== "table_row_insert" || row.position !== "end" || typeof row.id !== "string" || !row.id ||
    typeof row.documentId !== "string" || !row.documentId || typeof row.explanation !== "string" ||
    typeof row.createdAt !== "string" || !Array.isArray(row.openItems) || row.openItems.some(item => typeof item !== "string") ||
    !wordChatTableCaptureCurrent(snapshot, row.table) || !wordChatTableCaptureCurrent(snapshot, snapshot.table) ||
    row.table.tableIndex !== snapshot.table.tableIndex || row.table.rows.length !== snapshot.table.rows.length ||
    row.table.rows.some((cells, i) => cells.length !== snapshot.table!.rows[i]!.length ||
      cells.some((cell, j) => cell.id !== snapshot.table!.rows[i]![j]!.id || cell.text !== snapshot.table!.rows[i]![j]!.text)) || !Array.isArray(row.cells) ||
    row.cells.length !== row.table.rows[0]!.length) return false
  return row.cells.some(cell => typeof cell === "string" && cell.length > 0) && row.cells.every(cell =>
    typeof cell === "string" && cell.length <= 4000 && cell === cell.trim() && wellFormedWordChatCell(cell) && !/[\u0000-\u001f\u007f]/.test(cell)) && row.cells.join("").length <= 24_000 &&
    (row.runs === undefined || Array.isArray(row.runs) && canonicalWordChatCellRuns(row.cells, row.runs) === row.runs)
}
const runFlags = ["bold", "italic", "underline"] as const
/** Canonical cell runs: one entry per cell, non-empty texts that concatenate
 * to the cell, adjacent runs that differ, and at least one formatted span.
 * Returns the same array when already canonical, a normalized copy for loose
 * model output (missing flags are off; empty texts dropped; equal neighbours
 * merged), undefined for a plain row, or null when the runs do not match. */
export function canonicalWordChatCellRuns(cells: string[], value: unknown): WordChatCellRun[][] | undefined | null {
  if (value === undefined) return undefined
  if (!Array.isArray(value) || value.length !== cells.length) return null
  let canonical = true, formatted = false, count = 0
  const result: WordChatCellRun[][] = []
  for (const [index, entry] of value.entries()) {
    if (!Array.isArray(entry) || entry.length > 64) return null
    const merged: WordChatCellRun[] = []
    for (const run of entry) {
      if (!run || typeof run !== "object" || Array.isArray(run) || Object.keys(run).some(key => key !== "text" && !(runFlags as readonly string[]).includes(key)) ||
        typeof run.text !== "string" || !wellFormedWordChatCell(run.text) || runFlags.some(flag => run[flag] !== undefined && typeof run[flag] !== "boolean")) return null
      if (runFlags.some(flag => run[flag] === undefined)) canonical = false
      const next: WordChatCellRun = { text: run.text, bold: run.bold === true, italic: run.italic === true, underline: run.underline === true }
      if (!next.text) { canonical = false; continue }
      const last = merged.at(-1)
      if (last && runFlags.every(flag => last[flag] === next[flag])) { canonical = false; last.text += next.text; continue }
      merged.push(next)
      formatted ||= runFlags.some(flag => next[flag])
    }
    if (merged.map(run => run.text).join("") !== cells[index]) return null
    count += merged.length
    result.push(merged)
  }
  if (count > 256) return null
  if (!formatted) return undefined
  return canonical ? value as WordChatCellRun[][] : result
}
/** A deletion binds the exact existing row and same-read capture. */
export function validWordChatTableRowDeletion(snapshot: WordChatSnapshot, value: unknown): value is WordChatTableRowDeletionProposal {
  if (!value || typeof value !== "object" || Array.isArray(value)) return false
  const row = value as WordChatTableRowDeletionProposal
  if (Object.keys(row).some(key => !["kind", "id", "documentId", "explanation", "openItems", "createdAt", "table", "rowIndex", "cells"].includes(key)) ||
    row.kind !== "table_row_delete" || typeof row.id !== "string" || !row.id || typeof row.documentId !== "string" || !row.documentId ||
    typeof row.explanation !== "string" || typeof row.createdAt !== "string" || !Array.isArray(row.openItems) || row.openItems.some(item => typeof item !== "string") ||
    !wordChatTableCaptureCurrent(snapshot, row.table) || !wordChatTableCaptureCurrent(snapshot, snapshot.table) ||
    row.table.tableIndex !== snapshot.table.tableIndex || row.table.rows.length !== snapshot.table.rows.length ||
    row.table.rows.some((cells, i) => cells.length !== snapshot.table!.rows[i]!.length || cells.some((cell, j) => cell.id !== snapshot.table!.rows[i]![j]!.id || cell.text !== snapshot.table!.rows[i]![j]!.text)) || row.table.rows.length < 2 || !Number.isSafeInteger(row.rowIndex) || row.rowIndex < 0 || row.rowIndex >= row.table.rows.length ||
    !Array.isArray(row.cells) || row.cells.length !== row.table.rows[row.rowIndex]!.length) return false
  return row.cells.every((cell, i) => typeof cell === "string" && cell === row.table.rows[row.rowIndex]![i]!.text && cell.length <= 4000 && wellFormedWordChatCell(cell) && !/[\u0000-\u001f\u007f]/.test(cell)) && row.cells.some(cell => cell.length > 0) && row.cells.join("").length <= 24_000
}
function wellFormedWordChatCell(value: string): boolean {
  for (let i = 0; i < value.length; i++) {
    const code = value.charCodeAt(i)
    if (code >= 0xd800 && code <= 0xdbff) {
      const next = value.charCodeAt(++i)
      if (!(next >= 0xdc00 && next <= 0xdfff)) return false
    } else if (code >= 0xdc00 && code <= 0xdfff) return false
  }
  return true
}
export type WordChatProposal = {
  id: string
  documentId: string
  before: string
  after: string
  explanation: string
  openItems: string[]
  target: WordChatTarget
  createdAt: string
  /** Explicit structural action, separate from ordinary wording replacement. */
  paragraphAction?: "delete" | "append_list_item" | "insert_list_item_before" | "merge" | "split"
  /** Captured by the host and attached by the server for list insertion only. */
  list?: WordChatListCapture
  /** Optional reviewed paragraph properties, indexed against the resulting text. */
  formatting?: WordChatParagraphFormat[]
  /** Optional reviewed character properties on the exact captured span. */
  inlineFormatting?: WordChatInlineFormat
}
export type WordChatInlineFormat = { bold?: boolean; italic?: boolean; underline?: boolean }
/** A reviewed property command for the live Word executor. The resulting
 * native document, rather than a simulated DOCX, is verified and synchronized. */
export type WordChatInlineFormattingPlan = {
  kind: "inline_format"
  proposalId: string
  target: WordChatTarget
  inlineFormatting: WordChatInlineFormat
}
export type WordChatParagraphInsertionPlan = {
  kind: "paragraph_insert"
  proposalId: string
  target: WordChatTarget
  text: string
  formatting?: WordChatParagraphFormat[]
}
export type WordChatParagraphSplitPlan = { kind: "paragraph_split"; proposalId: string; target: WordChatTarget }
export type WordChatParagraphMergePlan = { kind: "paragraph_merge"; proposalId: string; target: WordChatTarget }
export type WordChatParagraphDeletionPlan = { kind: "paragraph_delete"; proposalId: string; target: WordChatTarget }
export type WordChatListInsertionPlan = { kind: "list_insert"; proposalId: string; target: WordChatTarget; text: string; list: WordChatListCapture; position?: "end" | "before_final" }
export type WordChatPreparedAmendment = WordChatListInsertionPlan | WordChatAuthoredAmendment | WordChatInlineFormattingPlan | WordChatParagraphInsertionPlan | WordChatParagraphDeletionPlan | WordChatParagraphMergePlan | WordChatParagraphSplitPlan
export type WordChatParagraphFormat = {
  /** Zero-based index within the resulting complete paragraphs. */
  paragraph: number
  style?: "Title" | "Heading1" | "Heading2" | "Heading3" | "Normal"
  /** Paragraph spacing in points, not Word twips. */
  spacingBefore?: number
  spacingAfter?: number
  /** A multiple of single line spacing. */
  lineSpacing?: number
}
export type WordChatAuthoredAmendment = {
  proposalId: string
  ooxml: string
  paragraphs: string[]
  revisions: { author: string; type: "Added" | "Deleted" | "Formatted"; text: string }[]
  author: "doc.haus"
  formatting?: WordChatParagraphFormat[]
  wording?: { ooxml: string; unformattedOoxml?: string; revisions: WordChatAuthoredAmendment["revisions"] }
}

/** Bind one literal inserted paragraph to a uniquely resolved complete source
 * paragraph; its only optional property is one allowed style. */
export function validWordChatParagraphInsertion(snapshot: WordChatSnapshot, target: WordChatTarget, replacement: unknown, formatting?: unknown, inlineFormatting?: unknown): boolean {
  if (!snapshot || !target || !Array.isArray(snapshot.paragraphs) || !Array.isArray(target.paragraphs) || target.paragraphs.length !== 1 ||
    typeof target.text !== "string" || !target.text.trim() || !wellFormedWordChatCell(target.text) ||
    target.startOffset !== 0 || target.endOffset !== target.paragraphs[0]?.text.length || target.text !== target.paragraphs[0]?.text ||
    !wordChatTargetCurrent(snapshot,target) || snapshot.selection && !wordChatTargetCurrent(snapshot,snapshot.selection) ||
    inlineFormatting !== undefined || typeof replacement !== "string" || !replacement.startsWith(target.text + "\r")) return false
  if (formatting !== undefined && (!Array.isArray(formatting) || formatting.length !== 1 || !formatting[0] || typeof formatting[0] !== "object" || Array.isArray(formatting[0]) ||
    (Object.keys(formatting[0]).length !== 2 || Object.keys(formatting[0]).some(key => !["paragraph", "style"].includes(key))) || formatting[0].paragraph !== 1 || !["Title", "Heading1", "Heading2", "Heading3", "Normal"].includes(formatting[0].style))) return false
  const inserted = replacement.slice(target.text.length + 1)
  if (!inserted.trim() || inserted.length > 24_000 || !wellFormedWordChatCell(inserted) || /[\u0000-\u001f\u007f]/.test(inserted)) return false
  const resolved = targetWordChatAmendment(snapshot,target.text)
  return !!resolved && resolved.id === target.id && resolved.capturedAt === target.capturedAt && resolved.saved === target.saved &&
    resolved.startOffset === target.startOffset && resolved.endOffset === target.endOffset && resolved.paragraphs.length === 1 &&
    resolved.paragraphs[0]!.id === target.paragraphs[0]!.id && resolved.paragraphs[0]!.text === target.paragraphs[0]!.text
}

/** An explicit list action is bound to the last captured complete member. */
export function validWordChatListInsertion(snapshot: WordChatSnapshot, target: WordChatTarget, replacement: unknown, action: unknown, formatting?: unknown, inlineFormatting?: unknown): boolean {
  const list = snapshot?.list
  if ((action !== "append_list_item" && action !== "insert_list_item_before") || formatting !== undefined || inlineFormatting !== undefined ||
    !wordChatListCaptureCurrent(snapshot, list) || list.members.at(-1)!.id !== target?.paragraphs?.[0]?.id ||
    list.members.at(-1)!.text !== target?.text || typeof replacement !== "string") return false
  if (action === "append_list_item") return validWordChatParagraphInsertion(snapshot, target, replacement)
  const suffix = `\r${target.text}`
  if (!replacement.endsWith(suffix)) return false
  const inserted = replacement.slice(0, -suffix.length)
  return validWordChatParagraphInsertion(snapshot, target, `${target.text}\r${inserted}`)
}

/** An explicit action distinguishes removing the paragraph mark from clearing text. */
export function validWordChatParagraphDeletion(snapshot: WordChatSnapshot, target: WordChatTarget, replacement: unknown, action: unknown, formatting?: unknown, inlineFormatting?: unknown): boolean {
  if (!snapshot || !target || !Array.isArray(snapshot.paragraphs) || !Array.isArray(target.paragraphs) || target.paragraphs.length !== 1 ||
    typeof target.text !== "string" || !target.text.trim() || !wellFormedWordChatCell(target.text) ||
    target.startOffset !== 0 || target.endOffset !== target.paragraphs[0]?.text.length || target.text !== target.paragraphs[0]?.text ||
    !wordChatTargetCurrent(snapshot,target) || snapshot.selection && !wordChatTargetCurrent(snapshot,snapshot.selection) ||
    replacement !== "" || action !== "delete" || formatting !== undefined || inlineFormatting !== undefined) return false
  const resolved = targetWordChatAmendment(snapshot,target.text)
  return !!resolved && resolved.id === target.id && resolved.capturedAt === target.capturedAt && resolved.saved === target.saved &&
    resolved.startOffset === target.startOffset && resolved.endOffset === target.endOffset && resolved.paragraphs.length === 1 &&
    resolved.paragraphs[0]!.id === target.paragraphs[0]!.id && resolved.paragraphs[0]!.text === target.paragraphs[0]!.text
}

/** Merge only two complete adjacent source paragraphs with one literal space. */
export function validWordChatParagraphMerge(snapshot: WordChatSnapshot, target: WordChatTarget, replacement: unknown, action: unknown, formatting?: unknown, inlineFormatting?: unknown): boolean {
  if (!snapshot || !target || !Array.isArray(target.paragraphs) || target.paragraphs.length !== 2 || action !== "merge" || formatting !== undefined || inlineFormatting !== undefined ||
    !wordChatTargetCurrent(snapshot, target) || snapshot.selection && !wordChatTargetCurrent(snapshot, snapshot.selection)) return false
  const [left,right] = target.paragraphs
  if (!left || !right || left.id === right.id || target.startOffset !== 0 || target.endOffset !== right.text.length ||
    [left.text,right.text].some(text => !text.trim() || !wellFormedWordChatCell(text) || /[\u0000-\u001f\u007f]/.test(text)) ||
    target.text !== left.text + "\r" + right.text || replacement !== left.text + " " + right.text ||
    snapshot.list?.members.some(member => member.id === left.id || member.id === right.id)) return false
  const resolved = targetWordChatAmendment(snapshot,target.text)
  return !!resolved && resolved.id === target.id && resolved.capturedAt === target.capturedAt && resolved.saved === target.saved &&
    resolved.startOffset === target.startOffset && resolved.endOffset === target.endOffset && resolved.paragraphs.length === 2 &&
    resolved.paragraphs.every((p,i) => p.id === target.paragraphs[i]!.id && p.text === target.paragraphs[i]!.text)
}
/** Replace exactly one existing space with a paragraph boundary. */
export function validWordChatParagraphSplit(snapshot: WordChatSnapshot, target: WordChatTarget, replacement: unknown, action: unknown, formatting?: unknown, inlineFormatting?: unknown): boolean {
  if (!snapshot || !target || !Array.isArray(target.paragraphs) || target.paragraphs.length !== 1 || action !== "split" || formatting !== undefined || inlineFormatting !== undefined ||
    typeof replacement !== "string" || !wordChatTargetCurrent(snapshot,target) || snapshot.selection && !wordChatTargetCurrent(snapshot,snapshot.selection)) return false
  const paragraph = target.paragraphs[0]!, pieces = replacement.split("\r")
  if (target.startOffset !== 0 || target.endOffset !== paragraph.text.length || target.text !== paragraph.text || pieces.length !== 2 ||
    pieces.some(piece => !piece.trim() || !wellFormedWordChatCell(piece) || /[\u0000-\u001f\u007f]/.test(piece)) ||
    !wellFormedWordChatCell(target.text) || /[\u0000-\u001f\u007f]/.test(target.text) || target.text !== pieces[0] + " " + pieces[1] ||
    snapshot.list?.members.some(member => member.id === paragraph.id)) return false
  const resolved = targetWordChatAmendment(snapshot,target.text)
  return !!resolved && resolved.id === target.id && resolved.capturedAt === target.capturedAt && resolved.saved === target.saved &&
    resolved.startOffset === target.startOffset && resolved.endOffset === target.endOffset && resolved.paragraphs.length === 1 &&
    resolved.paragraphs[0]!.id === paragraph.id && resolved.paragraphs[0]!.text === paragraph.text
}
/** Bind inline properties to the unique resolved native paragraph and UTF-16
 * offsets. The model supplies wording, never an identity or offset. */
export function validWordChatInlineFormatting(snapshot: WordChatSnapshot, target: WordChatTarget, value: unknown,
  replacement: unknown = target?.text, paragraphFormatting?: unknown): value is WordChatInlineFormat {
  if (!snapshot || !target || typeof target.id !== "string" || !target.id || typeof target.text !== "string" ||
    typeof target.capturedAt !== "string" || typeof target.saved !== "boolean" || !Array.isArray(snapshot.paragraphs) ||
    snapshot.paragraphs.some(paragraph => !paragraph || typeof paragraph.id !== "string" || typeof paragraph.text !== "string") ||
    !Array.isArray(target.paragraphs) || target.paragraphs.length !== 1 ||
    snapshot.selection && !wordChatTargetCurrent(snapshot, snapshot.selection) ||
    !wordChatTargetCurrent(snapshot, target) || typeof replacement !== "string" || replacement !== target.text ||
    !wellFormedWordChatCell(target.text) || paragraphFormatting !== undefined ||
    !value || typeof value !== "object" || Array.isArray(value)) return false
  const resolved = targetWordChatAmendment(snapshot, target.text)
  if (!resolved || resolved.id !== target.id || resolved.text !== target.text || resolved.startOffset !== target.startOffset ||
    resolved.endOffset !== target.endOffset || resolved.capturedAt !== target.capturedAt || resolved.saved !== target.saved ||
    resolved.paragraphs.length !== 1 || resolved.paragraphs[0]!.id !== target.paragraphs[0]!.id || resolved.paragraphs[0]!.text !== target.paragraphs[0]!.text) return false
  const format = value as Record<string, unknown>, keys = Object.keys(format)
  return keys.length > 0 && keys.every(key => ["bold", "italic", "underline"].includes(key) && typeof format[key] === "boolean")
}

/** Paragraph formatting cannot borrow a partial selection's authority to
 * change the surrounding paragraph, or carry arbitrary model-authored XML. */
export function validWordChatFormatting(target: WordChatTarget, value: unknown, replacement = target.text): value is WordChatParagraphFormat[] {
  if (!target?.paragraphs?.length || target.paragraphs.length > 64 || target.startOffset !== 0 ||
    target.endOffset !== target.paragraphs.at(-1)!.text.length ||
    target.text !== target.paragraphs.map(p => p.text).join("\r") ||
    /\n/.test(replacement) || replacement.split("\r").length > 64 ||
    !Array.isArray(value) || !value.length || value.length > 64) return false
  const seen = new Set<number>()
  const fields = new Set(["paragraph", "style", "spacingBefore", "spacingAfter", "lineSpacing"])
  return value.every(op => {
    if (!op || typeof op !== "object" || Array.isArray(op) || Object.keys(op).some(key => !fields.has(key)) ||
      !Number.isSafeInteger(op.paragraph) || op.paragraph < 0 || op.paragraph >= replacement.split("\r").length || seen.has(op.paragraph)) return false
    seen.add(op.paragraph)
    if (op.style !== undefined && !["Title", "Heading1", "Heading2", "Heading3", "Normal"].includes(op.style)) return false
    for (const field of ["spacingBefore", "spacingAfter"] as const) if (op[field] !== undefined &&
      (typeof op[field] !== "number" || !Number.isFinite(op[field]) || op[field] < 0 || op[field] > 48 || !Number.isInteger(op[field] * 20))) return false
    if (op.lineSpacing !== undefined && (typeof op.lineSpacing !== "number" || !Number.isFinite(op.lineSpacing) ||
      op.lineSpacing < 1 || op.lineSpacing > 2 || !Number.isInteger(op.lineSpacing * 240))) return false
    return [...fields].some(field => field !== "paragraph" && op[field] !== undefined)
  })
}

/** JSON/model paragraph separators use LF on some providers. Word Range.text
 * uses CR; soft line breaks remain their distinct native U+000B character. */
export const wordChatParagraphBreaks = (text: string) => text.replace(/\r\n|\n/g, "\r")

export type WordChatApplyPlan = { supported: true; paragraphs: string[];
  structure?: { kind: "insert" | "delete"; retained: { before: number; after: number }[] } } | { supported: false; reason: string }

/** The scoped Word import ends before the existing final paragraph mark. Keep
 * that paragraph intact for structural changes; otherwise rejecting a revision
 * could join the clause to unrelated text after the approved passage. */
export function wordChatApplyPlan(target: WordChatTarget, replacement: string, formatting?: WordChatParagraphFormat[]): WordChatApplyPlan {
  if (formatting !== undefined && !validWordChatFormatting(target, formatting, replacement))
    return { supported: false, reason: "Select complete paragraphs and prepare valid paragraph formatting." }
  const unsupported = { supported: false as const, reason: "This change needs paragraph restructuring. Refine it into smaller changes, or make it directly in Word." }
  if (/\n/.test(replacement) || !target.paragraphs.length) return unsupported
  const before = target.paragraphs.map(p => p.text)
  const paragraphs = (before[0]!.slice(0, target.startOffset) + replacement + before.at(-1)!.slice(target.endOffset)).split("\r")
  if (paragraphs.length > 64) return { supported: false, reason: "Refine this change to use at most 64 paragraphs." }
  if (paragraphs.length === before.length) return { supported: true, paragraphs }
  if (before.at(-1) !== paragraphs.at(-1) || before.some(p => !p.length) || paragraphs.some(p => !p.length || /[\t\v]/.test(p))) return unsupported
  const insert = paragraphs.length > before.length, shorter = insert ? before : paragraphs, longer = insert ? paragraphs : before
  if (formatting && !insert) return unsupported
  // Count exact ordered alignments, pinning the unchanged final paragraph. A
  // repeated paragraph must not silently become the one that is removed.
  const paths = Array.from({ length: shorter.length }, (_, i) => Array<number>(longer.length).fill(i === shorter.length - 1 ? 1 : 0))
  for (let i = shorter.length - 2; i >= 0; i--) for (let j = longer.length - 2; j >= 0; j--)
    paths[i]![j] = Math.min(2, paths[i]![j + 1]! + (shorter[i] === longer[j] ? paths[i + 1]![j + 1]! : 0))
  if (paths[0]![0] !== 1) {
    if (!insert || paths[0]![0] > 1 || target.startOffset !== 0 || target.endOffset !== before.at(-1)!.length) return unsupported
    // Infer rewritten retained paragraphs only when a unique ordered mapping
    // has strong wording overlap. Ambiguous mappings must be reviewed smaller.
    const score = (a: string, b: string) => {
      const words = (s: string) => new Set(s.toLowerCase().match(/[\p{L}\p{N}]+/gu) ?? [])
      const x = words(a), y = words(b), common = [...x].filter(w => y.has(w)).length
      return common >= 2 && common / Math.max(x.size, y.size) >= 0.5 ? common / Math.max(x.size, y.size) : -Infinity
    }
    const rows = before.length - 1, cols = paragraphs.length - 1
    const dp = Array.from({ length: rows + 1 }, () => Array<{ score: number; count: number }>(cols + 1))
    for (let j = 0; j <= cols; j++) dp[rows]![j] = { score: 0, count: 1 }
    for (let i = rows - 1; i >= 0; i--) for (let j = cols; j >= 0; j--) {
      const skip = j < cols ? dp[i]![j + 1]! : { score: -Infinity, count: 0 }
      const next = j < cols ? dp[i + 1]![j + 1]! : { score: -Infinity, count: 0 }
      const takeScore = j < cols ? score(before[i]!, paragraphs[j]!) + next.score : -Infinity
      const best = Math.max(skip.score, takeScore)
      dp[i]![j] = { score: best, count: best === -Infinity ? 0 : Math.min(2, (skip.score === best ? skip.count : 0) + (takeScore === best ? next.count : 0)) }
    }
    if (dp[0]![0]!.count !== 1) return unsupported
    const retained: { before: number; after: number }[] = []
    let j = 0
    for (let i = 0; i < rows; i++) {
      while (j < cols && score(before[i]!, paragraphs[j]!) + dp[i + 1]![j + 1]!.score !== dp[i]![j]!.score) j++
      retained.push({ before: i, after: j++ })
    }
    retained.push({ before: rows, after: cols })
    return { supported: true, paragraphs, structure: { kind: "insert", retained } }
  }
  const retained: { before: number; after: number }[] = []
  let at = 0
  for (let i = 0; i < shorter.length; i++) {
    if (i === shorter.length - 1) at = longer.length - 1
    else while (at < longer.length - 1 && !(shorter[i] === longer[at] && paths[i + 1]![at + 1])) at++
    retained.push(insert ? { before: i, after: at } : { before: at, after: i }); at++
  }
  return { supported: true, paragraphs, structure: { kind: insert ? "insert" : "delete", retained } }
}

/** A retained selection must still name the same adjacent native paragraphs,
 * not merely words that happen to survive elsewhere in the document. */
export function wordChatTargetCurrent(snapshot: WordChatSnapshot, target: WordChatTarget): boolean {
  if (!target || typeof target.text !== "string" || !target.text.trim() || target.text.length > 24_000 ||
    !Array.isArray(target.paragraphs) || !target.paragraphs.length || target.paragraphs.length > 64 ||
    target.paragraphs.some(paragraph => !paragraph || typeof paragraph.id !== "string" || !paragraph.id || typeof paragraph.text !== "string") ||
    !Number.isSafeInteger(target.startOffset) || !Number.isSafeInteger(target.endOffset)) return false
  const first = target.paragraphs[0]!
  const at = snapshot.paragraphs.findIndex(paragraph => paragraph.id === first.id)
  if (at < 0 || snapshot.paragraphs.some((paragraph, index) => index !== at && paragraph.id === first.id)) return false
  const pieces: string[] = []
  for (const [index, paragraph] of target.paragraphs.entries()) {
    const current = snapshot.paragraphs[at + index]
    if (!current || current.id !== paragraph.id || current.text !== paragraph.text) return false
    const start = index === 0 ? target.startOffset : 0
    const end = index === target.paragraphs.length - 1 ? target.endOffset : paragraph.text.length
    if (start < 0 || end < start || end > paragraph.text.length) return false
    pieces.push(paragraph.text.slice(start, end))
  }
  return pieces.join("\r") === wordChatParagraphBreaks(target.text)
}

/** The model supplies words, never native IDs or offsets. Resolve one exact
 * occurrence inside the explicitly selected passage, otherwise the open buffer.
 * The original paragraph sequence and selected boundaries remain authoritative. */
export function targetWordChatAmendment(snapshot: WordChatSnapshot, before: string): WordChatTarget | undefined {
  const quote = wordChatParagraphBreaks(before)
  if (!quote.trim() || quote.length > 24_000) return
  const selected = snapshot.selection
  const paragraphs = selected ? selected.paragraphs : snapshot.paragraphs
  if (!paragraphs.length || paragraphs.some(paragraph => !paragraph.id)) return
  const pieces: { text: string; start: number; offset: number }[] = []
  let length = 0
  for (const [index, paragraph] of paragraphs.entries()) {
    const start = selected && index === 0 ? selected.startOffset : 0
    const end = selected && index === paragraphs.length - 1 ? selected.endOffset : paragraph.text.length
    if (!Number.isSafeInteger(start) || !Number.isSafeInteger(end) || start < 0 || end < start || end > paragraph.text.length) return
    const text = paragraph.text.slice(start, end)
    pieces.push({text,start,offset:length})
    length += text.length + 1
  }
  const passage = pieces.map(piece => piece.text).join("\r")
  if (selected && passage !== wordChatParagraphBreaks(selected.text)) return
  const at = passage.indexOf(quote)
  if (at < 0 || passage.indexOf(quote, at + 1) >= 0) return
  const locate = (offset: number) => {
    const index = pieces.findIndex(piece => piece.offset <= offset && offset <= piece.offset + piece.text.length)
    const piece = pieces[index]
    return piece ? {index,offset:piece.start + offset - piece.offset} : undefined
  }
  const start = locate(at), end = locate(at + quote.length)
  if (!start || !end || end.index - start.index >= 64) return
  return { id: `${selected?.id ?? snapshot.capturedAt}:${paragraphs[start.index]!.id}:${start.offset}`, text: quote,
    paragraphs: paragraphs.slice(start.index,end.index + 1).map(paragraph => ({...paragraph})), startOffset: start.offset, endOffset: end.offset,
    capturedAt: snapshot.capturedAt, saved: snapshot.saved }
}
