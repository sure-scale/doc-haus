// Shared matcher for "does this text use the defined term?". A use is the
// recorded term, case-sensitive (defined terms are capitalized; case folding
// would count ordinary prose such as "third party" as a use of "Party"), with
// no letter, combining mark, digit or underscore immediately before or after it
// ("Term" is not used in "Terminate"). The neighbours are read as whole code
// points, so a letter outside the Basic Multilingual Plane also blocks a match.
// Every occurrence is checked, so a longer word earlier in the text
// ("Agreements") cannot hide a later real use ("Agreement").
// Plural and other inflected forms ("Terms") are not uses; an apostrophe is not
// a word character, so a possessive ("Party's") still counts.

const WORD_CHARACTER = /[\p{L}\p{M}\p{N}_]/u

function isWordCharacter(character: string | undefined) {
  return character !== undefined && WORD_CHARACTER.test(character)
}

// The whole code point ending just before `index`; a lone or leading surrogate
// half stays as it is.
function characterBefore(text: string, index: number) {
  return Array.from(text.slice(Math.max(0, index - 2), index)).at(-1)
}

function characterAt(text: string, index: number) {
  const codePoint = text.codePointAt(index)
  return codePoint === undefined ? undefined : String.fromCodePoint(codePoint)
}

/** Index of the first whole-word, case-sensitive occurrence of `term` in `text`, or -1. */
export function findTermUse(text: string, term: string) {
  if (!term) return -1
  for (let at = text.indexOf(term); at !== -1; at = text.indexOf(term, at + 1)) {
    if (!isWordCharacter(characterBefore(text, at)) && !isWordCharacter(characterAt(text, at + term.length))) return at
  }
  return -1
}
