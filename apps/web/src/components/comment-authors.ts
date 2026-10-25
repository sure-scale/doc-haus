// Who wrote a tracked change or comment, as Word marks them: the author's
// initials and a colour of their own. Shared by the Tracked changes panel and
// the comments beside the document, so an author has one colour everywhere.

// The number of author colours, --author-1 to --author-6 in design-system.css.
export const TONES = 6

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

function hashName(name: string) {
  let hash = 0
  for (const char of name) hash = (hash * 31 + char.codePointAt(0)!) >>> 0
  return hash
}

// Each author's colour: picked by the name, so it stays put between versions,
// and moved on to the next free one when another author in the file has it.
export function authorTones(names: string[]): Map<string, number> {
  const tones = new Map<string, number>()
  const taken = new Set<number>()
  for (const name of [...new Set(names.map(nameKey))].sort()) {
    let tone = hashName(name) % TONES
    for (let step = 0; step < TONES && taken.has(tone); step++) tone = (tone + 1) % TONES
    taken.add(tone)
    if (taken.size === TONES) taken.clear()
    tones.set(name, tone)
  }
  return tones
}

// One author's colour among `names` (every author in the file): "self" for the
// lawyer's own name, shown in the brand colour, else 0 to TONES - 1.
export function toneOf(tones: Map<string, number>, author: string, self?: (author: string) => boolean): number | "self" {
  if (self?.(author)) return "self"
  return tones.get(nameKey(author)) ?? hashName(nameKey(author)) % TONES
}

// The class that paints an author's colour: tone-1 to tone-6, or tone-self.
export const toneClass = (tone: number | "self") => `tone-${tone === "self" ? "self" : tone + 1}`
