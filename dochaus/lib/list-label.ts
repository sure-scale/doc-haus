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
 * which would print beside the label. Not when `current`, the paragraph's
 * present text, opens with the same number ("1.5 times the Fees" under 2.1):
 * that number is its text. */
export function retypedListLabel(text: string, label: string | undefined, current?: string) {
  if (!label) return
  const shape = label.trim().replace(/\d+|\p{L}+|\s+|[^\d\p{L}\s]/gu, part =>
    /^\d/.test(part) ? "\\d{1,3}"
      : /^\s/.test(part) ? "\\s+"
        : /^\p{L}/u.test(part) ? part.length <= 2 || /^(?:[ivxlcdm]+|[IVXLCDM]+)$/.test(part) ? "\\p{L}{1,4}" : part
          : part.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"))
  const opening = (value: string) => new RegExp(`^\\s*(${shape})[ \\t\\u00a0]+(?=\\S)`, "iu").exec(value)?.[1]
  const typed = opening(text)
  return typed && typed !== label.trim() && (current === undefined || opening(current) !== typed) ? typed : undefined
}

/** Find text and replacement for an edit tool, with the paragraphs `find`
 * names. `find` is kept when `search` finds it. Otherwise, when it opens with
 * a paragraph's own list label and the rest opens that paragraph's shown text,
 * the edit is to those paragraphs: both texts lose the label. A tool that
 * records the edit as text its accept finds again (`recorded`) gets `error`
 * instead when the rest is also in other paragraphs, which accept would edit
 * too. `search` returns the paragraphs containing a text with their labels
 * (Docxodus autoNumberPrefix); `shown` returns a paragraph's text with its
 * label. */
export function findPastListLabel<T extends { id: string; autoNumberPrefix?: string }>(
  find: string, replace: string, search: (text: string) => T[], shown: (id: string) => string | undefined,
  { recorded = false } = {},
): { find: string; replace: string; targets: T[]; error?: string } {
  const targets = search(find)
  if (targets.length) return { find, replace, targets }
  // A label is one to three words ("2.1", "(a)", "Article 1").
  const words = /^\s*(\S+)(?:[ \t\u00a0]+(\S+))?(?:[ \t\u00a0]+(\S+))?/.exec(find)?.slice(1).filter(Boolean) ?? []
  for (let count = 1; count <= words.length; count++) {
    const label = words.slice(0, count).join(" ")
    const rest = withoutListLabel(find, label)
    if (rest === find) continue
    const found = search(rest)
    const labelled = found.filter(target => target.autoNumberPrefix?.trim() === label && withoutListLabel(shown(target.id) ?? "", label).startsWith(rest))
    if (!labelled.length) continue
    const error = recorded && found.length > labelled.length
      ? `${JSON.stringify(find)} begins with ${label}, the number Word prints before that paragraph, which is not text. Without it, ${JSON.stringify(rest)} is also in ${found.length - labelled.length} other paragraph(s), where accepting the change could apply it. Quote more of the text of ${label}, enough to be found only there, and leave the number out.`
      : undefined
    return { find: rest, replace: withoutListLabel(replace, label), targets: labelled, ...(error ? { error } : {}) }
  }
  return { find, replace, targets }
}
