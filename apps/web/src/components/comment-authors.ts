import { STRIP_RE } from "../../../../dochaus/lib/untrusted-text"

// Who wrote a tracked change or comment, as Word marks them: the author's
// initials and a colour of their own. Shared by the Tracked changes panel, the
// tracked changes and comments in the page and a clause's change history, so
// an author has one colour everywhere.

// The number of author colours, --author-1 to --author-6 in design-system.css.
export const TONES = 6

export type Tone = number | "self"

// An author's initials, as Word shows them on a comment: the file's own when
// it gives short ones, else the first letters of the first and last names.
export function initialsOf(name: string, given?: string | null) {
  const own = given?.trim()
  if (own && Array.from(own).length <= 3 && /^\p{L}+$/u.test(own)) return own.toUpperCase()
  const words = name.trim().split(/\s+/).filter((word) => /\p{L}/u.test(word))
  if (!words.length) return "?"
  const first = (word: string) => Array.from(word.replace(/^[^\p{L}]+/u, ""))[0] ?? ""
  return (first(words[0]!) + (words.length > 1 ? first(words.at(-1)!) : "")).toUpperCase()
}

export const nameKey = (name: string) => name.trim().toLowerCase()

// Each author's colour, as Word colours changes by author: the lawyer's own
// (`self`) in the brand colour, and every other author the next of the TONES
// colours in the order they first appear in `names` (every author in the
// file, in document order), starting again after the last.
export function authorTones(names: string[], self?: (author: string) => boolean): (author: string) => Tone {
  const tones = new Map<string, number>()
  for (const name of names) if (!self?.(name) && !tones.has(nameKey(name))) tones.set(nameKey(name), tones.size % TONES)
  return (author) => (self?.(author) ? "self" : tones.get(nameKey(author)) ?? 0)
}

// The class that paints an author's colour: tone-1 to tone-6, or tone-self.
export const toneClass = (tone: Tone) => `tone-${tone === "self" ? "self" : tone + 1}`
// The colour itself, for a style property.
export const toneColour = (tone: Tone) => (tone === "self" ? "var(--accent)" : `var(--author-${tone + 1})`)

const authorOf = (node: HTMLElement) => (node.dataset.author ?? "").replace(STRIP_RE, "")

// Every author a rendered page names, in the order their changes and comments
// first appear in its text: a comment where its text (or marker) is, not
// where its card sits, so the order holds however the cards are laid out.
function pageAuthors(root: HTMLElement) {
  const notes = new Map([...root.querySelectorAll<HTMLElement>(".comment-margin-note")].map((note) => [note.dataset.commentId ?? "", authorOf(note)]))
  const inText = [...root.querySelectorAll<HTMLElement>("[data-author]:not(.comment-margin-note), .comment-highlight[data-comment-id], a.comment-marker")].map((node) =>
    node.dataset.author !== undefined ? authorOf(node) : notes.get(node.dataset.commentId ?? node.id.replace(/^comment-ref-/, "")))
  return [...inText, ...notes.values()].filter((name) => name !== undefined)
}

// Each author's colour on a rendered page, one list for its tracked changes
// and its comment cards: the authors in `names` (the file's own, in document
// order, as the panel colours them) first, then any other the page names.
export const pageTones = (root: HTMLElement, names: string[] = [], self?: (author: string) => boolean) =>
  authorTones([...names, ...pageAuthors(root)], self)

// Each tracked change in a rendered page in its author's colour: an insertion
// or deletion, a whole paragraph, row or cell, a formatting change
// (styles.css reads --author-tone).
export function colourChanges(root: HTMLElement, names: string[], self?: (author: string) => boolean) {
  const tone = pageTones(root, names, self)
  for (const mark of root.querySelectorAll<HTMLElement>("[data-author]:not(.comment-margin-note)")) mark.style.setProperty("--author-tone", toneColour(tone(authorOf(mark))))
}
