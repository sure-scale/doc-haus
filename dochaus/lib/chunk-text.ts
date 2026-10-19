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

export type SectionChunkRow = ChunkTextRow & { doc_name: string; section: string }
export type SectionPassage = { doc_name: string; section: string; char_start: number; text: string }

// Section text for get-section and current-section. Rows come in (doc_name,
// char_start, rowid) order. A passage is one section's run of chunks that touch
// or overlap, rebuilt by chunkText from the run's first char_start: ingest
// repeats the last ~12% of a long section's chunk at the start of the next, so
// joining chunk text would repeat it. A label that recurs further on in a
// document (an exhibit's own numbering) starts a new passage with its own
// offsets. Passages end on a chunk boundary once `cap` characters are used,
// always keeping the first chunk; `rest` is where the text not returned starts.
export function sectionPassages(rows: readonly SectionChunkRow[], cap: number) {
  const passages: SectionPassage[] = []
  let rest: { doc_name: string; section: string; char_start: number } | undefined
  let run: SectionChunkRow[] = []
  let end = 0
  let used = 0
  const close = () => {
    const first = run[0]
    if (!first) return
    const rebased = run.map((row) => ({ text: row.text, char_start: row.char_start - first.char_start }))
    passages.push({ doc_name: first.doc_name, section: first.section, char_start: first.char_start, text: chunkText(rebased).text })
    run = []
  }
  for (const row of rows) {
    const first = run[0]
    if (first && (row.doc_name !== first.doc_name || row.section !== first.section || row.char_start > end)) close()
    const rowEnd = row.char_start + row.text.length
    const added = run.length ? Math.max(0, rowEnd - end) : row.text.length
    if (used + added > cap && passages.length + run.length > 0) {
      rest = { doc_name: row.doc_name, section: row.section, char_start: run.length ? end : row.char_start }
      break
    }
    end = run.length ? Math.max(end, rowEnd) : rowEnd
    run.push(row)
    used += added
  }
  close()
  return { passages, rest }
}
