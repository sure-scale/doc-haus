// A document's text rebuilt from its stored chunks, the text their char_start
// and char_end offsets index. ingest rebuilds the same text by the same rule
// (services/ingest/src/structure.ts reconstructText); every reader here uses
// this one rebuild so all of them agree with it.
//
// Rows come in (char_start, rowid) order. Each chunk overwrites the range it
// covers and never truncates text past its end: a shorter chunk with the same
// start cannot cut the text short, and the next chunk's text replaces the extra
// trailing newline that chunks ingested before the chunker's fix carried one
// character past a split section. A gap no chunk covers held only whitespace
// at ingest; it is padded with spaces, or with `pad(length, index)` before row
// `index`, so the offsets after it hold. No later chunk starts before the
// current one, so text before its start is final and only the tail from there
// on is rewritten, which keeps the rebuild linear.
//
// covered: the merged [start, end) ranges stored chunk text covers.

export type ChunkTextRow = { text: string; char_start: number }

export function chunkText(rows: readonly ChunkTextRow[], pad?: (length: number, index: number) => string) {
  const done: string[] = []
  let doneLength = 0
  let tail = ""
  const covered: [number, number][] = []
  rows.forEach((row, index) => {
    const at = row.char_start - doneLength
    if (at > tail.length) tail += pad ? pad(at - tail.length, index) : " ".repeat(at - tail.length)
    if (at > 0) {
      done.push(tail.slice(0, at))
      tail = tail.slice(at)
      doneLength = row.char_start
    }
    tail = row.text.length >= tail.length ? row.text : row.text + tail.slice(row.text.length)
    const end = row.char_start + row.text.length
    const last = covered[covered.length - 1]
    if (last && row.char_start <= last[1]) last[1] = Math.max(last[1], end)
    else covered.push([row.char_start, end])
  })
  return { text: done.join("") + tail, covered }
}
