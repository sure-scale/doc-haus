// Shared matcher for "does this text use the defined term?". A use is the
// recorded term, case-sensitive (defined terms are capitalized; case folding
// would count ordinary prose such as "third party" as a use of "Party"), with
// no letter, digit or underscore immediately before or after it ("Term" is not
// used in "Terminate"). Every occurrence is checked, so a longer word earlier in
// the text ("Agreements") cannot hide a later real use ("Agreement").
// Plural and other inflected forms ("Terms") are not uses; an apostrophe is not
// a word character, so a possessive ("Party's") still counts.

const WORD_CHARACTER = /[\p{L}\p{N}_]/u

function isWordCharacter(character: string | undefined) {
  return character !== undefined && WORD_CHARACTER.test(character)
}

/** Index of the first whole-word, case-sensitive occurrence of `term` in `text`, or -1. */
export function findTermUse(text: string, term: string) {
  if (!term) return -1
  for (let at = text.indexOf(term); at !== -1; at = text.indexOf(term, at + 1)) {
    if (!isWordCharacter(text[at - 1]) && !isWordCharacter(text[at + term.length])) return at
  }
  return -1
}
