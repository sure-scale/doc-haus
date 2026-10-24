// Word list numbering prints a label before a list paragraph ("2.1", "(a)",
// "Article 1"). The label is not part of the paragraph's text, which quotes,
// find text and amendments are matched against and edit. A reader shown the
// paragraph with its label may copy the label too; these helpers take it back
// out, only where it is that paragraph's own label.

/** A paragraph as a reader sees it: its list label, then its text. */
export function shownText(row: { text: string; label?: string }) {
  return row.label ? `${row.label} ${row.text}` : row.text
}

/** `text` without a leading `label` and the space or tab printed after it.
 * Unchanged unless it opens with exactly that label followed by more text. */
export function withoutListLabel(text: string, label: string | undefined) {
  const start = text.length - text.trimStart().length
  if (!label || !text.startsWith(label, start)) return text
  const gap = /^[ \t\u00a0]+(?=\S)/.exec(text.slice(start + label.length))
  return gap ? text.slice(start + label.length + gap[0].length) : text
}

/** Whether `quote` is in the paragraph as a reader sees it: in its text, or as
 * its own label followed by the opening of its text. `normalize` applies to
 * both sides. */
export function quotedInParagraph(row: { text: string; label?: string }, quote: string, normalize = (text: string) => text) {
  const text = normalize(row.text)
  if (text.includes(normalize(quote))) return true
  const rest = withoutListLabel(quote, row.label)
  return rest !== quote && text.startsWith(normalize(rest))
}

/** A number opening `text` that is written like `label` but is not it ("3.2"
 * on a paragraph Word numbers 2.2, "Article 2" on Article 1): typed numbering,
 * which would print beside the label. */
export function retypedListLabel(text: string, label: string | undefined) {
  if (!label) return
  const shape = label.trim().replace(/\d+|\p{L}+|\s+|[^\d\p{L}\s]/gu, part =>
    /^\d/.test(part) ? "\\d{1,3}"
      : /^\s/.test(part) ? "\\s+"
        : /^\p{L}/u.test(part) ? part.length <= 2 || /^(?:[ivxlcdm]+|[IVXLCDM]+)$/.test(part) ? "\\p{L}{1,4}" : part
          : part.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"))
  const typed = new RegExp(`^\\s*(${shape})[ \\t\\u00a0]+(?=\\S)`, "iu").exec(text)?.[1]
  return typed && typed !== label.trim() ? typed : undefined
}

