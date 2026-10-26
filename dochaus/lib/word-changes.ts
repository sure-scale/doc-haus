// The tracked change a proposal makes, as Word records one. Pure text, shared
// by the docx writer (docx-amendments.ts) and the web proposal card, so the
// card marks exactly the words the file changes. No imports on purpose: the
// browser bundle pulls this file directly.

/** From `offset` (UTF-16 units) of the old text, `replaced` becomes `changed`. */
export type WordChange = { offset: number; replaced: string; changed: string }

// Words, runs of white space and single other characters, as changes compare them.
// A clause number (2.1, 4.1.2, 2.1A) is one word, so a reference added beside
// it is never aligned into it; an amount's decimals after a thousands comma
// (1,000.50) stay words of their own.
function wordTokens(value: string) {
  return value.match(/(?<!\p{N},)\p{N}+(?:\.\p{N}+)+[\p{L}\p{N}_]*|[\p{L}\p{N}_]+|\s+|[^\p{L}\p{N}_\s]/gu) ?? []
}

// Tokens before[b0, b1) of the old text become after[a0, a1) of the new.
type Run = { b0: number; b1: number; a0: number; a1: number }

/** One change of whole words, from the first word that differs to the last. */
export function changedWords(text: string, replacement: string): WordChange {
  const before = wordTokens(text), after = wordTokens(replacement)
  return change(before, after, placed(before, after, [differing(before, after)])[0]!)
}

/** The changes Word records for `text` becoming `replacement`: whole words,
 * only those that differ, each run of them its own change, left to right. Words
 * kept between two changes are never deleted and typed again. Where all the
 * words between the first and last that differ stand, together and once,
 * inside the new words (two insertions Prepare joined into one proposal), the
 * changes are the insertion on each side of them; otherwise the words between
 * are aligned word by word, and changes with only white space between them
 * are one. */
export function wordChanges(text: string, replacement: string): WordChange[] {
  const before = wordTokens(text), after = wordTokens(replacement)
  const run = differing(before, after)
  const kept = before.slice(run.b0, run.b1), added = after.slice(run.a0, run.a1)
  const found = kept.length ? added.flatMap((_, index) => kept.every((token, k) => added[index + k] === token) ? [index] : []) : []
  const runs = found.length === 1
    ? [{ b0: run.b0, b1: run.b0, a0: run.a0, a1: run.a0 + found[0]! }, { b0: run.b1, b1: run.b1, a0: run.a0 + found[0]! + kept.length, a1: run.a1 }]
    : kept.length && added.length ? aligned(before, after, run) : [run]
  return placed(before, after, runs).map(row => change(before, after, row))
}

function change(before: string[], after: string[], run: Run): WordChange {
  return { offset: before.slice(0, run.b0).join("").length, replaced: before.slice(run.b0, run.b1).join(""),
    changed: after.slice(run.a0, run.a1).join("") }
}

// The tokens between the first and last that differ.
function differing(before: string[], after: string[]): Run {
  let first = 0
  let lastBefore = before.length
  let lastAfter = after.length
  while (first < lastBefore && first < lastAfter && before[first] === after[first]) first++
  while (lastBefore > first && lastAfter > first && before[lastBefore - 1] === after[lastAfter - 1]) {
    lastBefore--
    lastAfter--
  }
  return { b0: first, b1: lastBefore, a0: first, a1: lastAfter }
}

// The runs of a replacement whose words do not align: its old and new words
// compared as their longest common sequence, and runs with only white space
// between them joined. Too long to compare, it is one run.
function aligned(before: string[], after: string[], run: Run): Run[] {
  const x = before.slice(run.b0, run.b1), y = after.slice(run.a0, run.a1)
  if (x.length * y.length > 250_000) return [run]
  // Longest common run of tokens from x[i] and y[j] onwards.
  const width = y.length + 1
  const common = new Uint32Array((x.length + 1) * width)
  for (let i = x.length - 1; i >= 0; i--)
    for (let j = y.length - 1; j >= 0; j--)
      common[i * width + j] = x[i] === y[j] ? common[(i + 1) * width + j + 1]! + 1 : Math.max(common[(i + 1) * width + j]!, common[i * width + j + 1]!)
  const runs: Run[] = []
  let i = 0, j = 0
  while (i < x.length || j < y.length) {
    if (i < x.length && j < y.length && x[i] === y[j]) {
      i++
      j++
      continue
    }
    const b = run.b0 + i, a = run.a0 + j
    let open = runs.at(-1)
    if (!open || open.b1 !== b || open.a1 !== a) runs.push(open = { b0: b, b1: b, a0: a, a1: a })
    if (i < x.length && (j === y.length || common[(i + 1) * width + j]! >= common[i * width + j + 1]!)) {
      i++
      open.b1++
    } else {
      j++
      open.a1++
    }
  }
  const joined: Run[] = []
  for (const next of runs) {
    const last = joined.at(-1)
    if (last && before.slice(last.b1, next.b0).every(token => /^\s/.test(token))) Object.assign(last, { b1: next.b1, a1: next.a1 })
    else joined.push(next)
  }
  return joined
}

// An insertion or deletion that could stand at several places, because the
// words around it repeat, stands at the last whose both ends are sentence
// boundaries (the quote's start or end, or either side of the space after a
// full stop): a whole sentence goes between sentences, never inside the one
// whose words it repeats. Otherwise it stands at the first place that follows
// a space or the quote's start, so an inserted word goes before the word it
// qualifies; with no such place, at the last. Each keeps at least one word
// between it and the changes beside it.
function placed(before: string[], after: string[], runs: Run[]) {
  const space = (token?: string) => token !== undefined && /^\s/.test(token)
  const stop = (token?: string) => token !== undefined && /^[.!?]$/.test(token)
  return runs.reduce<Run[]>((out, run, index) => {
    const inserted = run.b0 === run.b1
    if (inserted === (run.a0 === run.a1)) return [...out, run]
    // The text holding the inserted or deleted words, and how far they may move.
    const holding = inserted ? after : before
    const [from, to] = inserted ? [run.a0, run.a1] : [run.b0, run.b1]
    const previous = out.at(-1), next = runs[index + 1]
    const low = previous ? (inserted ? previous.a1 : previous.b1) + 1 : 0
    const high = next ? (inserted ? next.a0 : next.b0) - 1 : holding.length
    // The last place first: moved one token later, the words still read the
    // same where the token after them matches their own first.
    let slide = 0
    while (to + slide < high && holding[from + slide] === holding[to + slide]) slide++
    const start = from + slide, end = to + slide
    const boundary = (at: number) => at === 0 || at === holding.length
      || (space(holding[at - 1]) && stop(holding[at - 2])) || (stop(holding[at - 1]) && space(holding[at]))
    let sentence: number | undefined
    let spaced = 0
    for (let k = 0; k <= start - low && (!k || holding[end - k] === holding[start - k]); k++) {
      if (sentence === undefined && boundary(start - k) && boundary(end - k)) sentence = k
      if (k && (k === start || space(holding[start - k - 1]))) spaced = k
    }
    const move = slide - (sentence ?? spaced)
    return [...out, { b0: run.b0 + move, b1: run.b1 + move, a0: run.a0 + move, a1: run.a1 + move }]
  }, [])
}
