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

// Each tracked insertion and deletion in a rendered page in its author's
// colour (styles.css reads --author-tone). The page's authors follow those in
// `names`, so the file's authors keep the panel's colours.
export function colourChanges(root: HTMLElement, names: string[], self?: (author: string) => boolean) {
  const marks = [...root.querySelectorAll<HTMLElement>(":is(ins, del)[data-author]")]
  const authorOf = (mark: HTMLElement) => mark.dataset.author!.replace(STRIP_RE, "")
  const tone = authorTones([...names, ...marks.map(authorOf)], self)
  for (const mark of marks) mark.style.setProperty("--author-tone", toneColour(tone(authorOf(mark))))
}
