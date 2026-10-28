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
}
export type WordChatAuthoredAmendment = {
  proposalId: string
  ooxml: string
  paragraphs: string[]
  revisions: { author: string; type: "Added" | "Deleted"; text: string }[]
  author: "doc.haus"
}

/** JSON/model paragraph separators use LF on some providers. Word Range.text
 * uses CR; soft line breaks remain their distinct native U+000B character. */
export const wordChatParagraphBreaks = (text: string) => text.replace(/\r\n|\n/g, "\r")

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
