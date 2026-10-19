// get-section returns a section's text verbatim from the chunk index.
// Run from dochaus/: bun test test/get-section.test.ts
import { Database } from "bun:sqlite"
import { afterEach, expect, test } from "bun:test"
import { mkdirSync, mkdtempSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"

const getSection = (await import("../tool/get-section")).default
const verifyQuoteTool = (await import("../tool/verify-quote")).default

const directories: string[] = []
afterEach(() => directories.splice(0).forEach((directory) => rmSync(directory, { recursive: true, force: true })))

function ctxFor(directory: string) {
  return {
    sessionID: "test",
    messageID: "test",
    agent: "test",
    directory,
    worktree: directory,
    abort: new AbortController().signal,
    metadata() {},
    async ask() {},
  }
}

type Row = { section: string; text: string; char_start: number }

function matter(rows: Row[]) {
  const directory = mkdtempSync(path.join(tmpdir(), "get-section-"))
  directories.push(directory)
  mkdirSync(path.join(directory, ".dochaus"))
  const db = new Database(path.join(directory, ".dochaus", "legal.db"))
  db.run("CREATE TABLE documents (id INTEGER PRIMARY KEY, doc_path TEXT UNIQUE, name TEXT)")
  db.run(
    "CREATE TABLE chunks (id INTEGER PRIMARY KEY, doc_path TEXT, doc_name TEXT, section TEXT, text TEXT, char_start INTEGER, char_end INTEGER)",
  )
  db.run("INSERT INTO documents (doc_path, name) VALUES ('/m/agreement.docx', 'Agreement')")
  for (const row of rows) {
    db.query("INSERT INTO chunks (doc_path, doc_name, section, text, char_start, char_end) VALUES (?, ?, ?, ?, ?, ?)").run(
      "/m/agreement.docx",
      "Agreement",
      row.section,
      row.text,
      row.char_start,
      row.char_start + row.text.length,
    )
  }
  db.close()
  return directory
}

// Lines of an invented clause, each distinct so a repeated span is detectable.
function clause(label: string, lines: number) {
  const body = Array.from(
    { length: lines },
    (_, i) => `Paragraph ${i + 1} of invented clause ${label}, written only to exceed one chunk and describing no actual obligation.\n`,
  )
  return `${label}. Invented heading\n${body.join("")}`
}

// Split a section the way ingest's chunker does: whole lines packed to ~2048
// characters, the trailing lines (~12% of the budget) repeated at the start of
// the next chunk.
function overlappingChunks(section: string, text: string, start: number): Row[] {
  const lines = text.split(/(?<=\n)/)
  const rows: Row[] = []
  let first = 0
  let offset = 0
  while (first < lines.length) {
    let last = first
    let length = 0
    while (last < lines.length && (last === first || length + lines[last]!.length <= 2048)) length += lines[last++]!.length
    rows.push({ section, text: lines.slice(first, last).join(""), char_start: start + offset })
    if (last === lines.length) break
    let back = last
    let overlap = 0
    while (back > first + 1 && overlap < 246) overlap += lines[--back]!.length
    offset += lines.slice(first, back).join("").length
    first = back
  }
  return rows
}

function sectionBody(result: string, header: string) {
  const at = result.indexOf(header)
  expect(at).toBeGreaterThanOrEqual(0)
  const rest = result.slice(at + header.length + 1)
  const end = rest.indexOf("\n\n=== ")
  return end === -1 ? rest : rest.slice(0, end)
}

test("an overlapping multi-chunk section is returned once, verbatim, at its source offsets", async () => {
  const preamble = "This invented agreement is a fixture.\n"
  const long = clause("1", 55)
  const next = "2. Next clause\nShort invented text.\n"
  const source = preamble + long + next
  const start = preamble.length
  const chunks = overlappingChunks("1", long, start)
  expect(chunks.length).toBeGreaterThan(2)
  const directory = matter([
    { section: "", text: preamble, char_start: 0 },
    ...chunks,
    { section: "2", text: next, char_start: start + long.length },
  ])

  const result = await getSection.execute({ section: "1", document: "Agreement" }, ctxFor(directory))
  const header = `=== Agreement — § 1 (chars ${start}–${start + long.length}) ===`
  expect(sectionBody(result, header)).toBe(source.slice(start, start + long.length))

  // A quote copied from the output across the first chunk seam is source text.
  const seam = chunks[0]!.text.length
  const quote = sectionBody(result, header).slice(seam - 120, seam + 120)
  const verified = await verifyQuoteTool.execute({ quote, document: "Agreement" }, ctxFor(directory))
  expect(verified).toStartWith("VERIFIED")
})

test("a label that recurs later in a document is a separate passage with its own offsets", async () => {
  const main = clause("1", 30)
  const between = clause("2", 3)
  const exhibit = "1. Exhibit heading\nAn invented exhibit clause.\n"
  const source = main + between + exhibit
  const directory = matter([
    ...overlappingChunks("1", main, 0),
    { section: "2", text: between, char_start: main.length },
    { section: "1", text: exhibit, char_start: main.length + between.length },
  ])

  const result = await getSection.execute({ section: "1" }, ctxFor(directory))
  const exhibitStart = main.length + between.length
  expect(sectionBody(result, `=== Agreement — § 1 (chars 0–${main.length}) ===`)).toBe(main)
  expect(sectionBody(result, `=== Agreement — § 1 (chars ${exhibitStart}–${source.length}) ===`)).toBe(exhibit)
  expect(result).not.toContain(between)
})

test("the text cap cuts an overlapping section at a chunk boundary and names where the rest starts", async () => {
  const long = clause("1", 260)
  const chunks = overlappingChunks("1", long, 0)
  const directory = matter(chunks)

  const result = await getSection.execute({ section: "1", document: "Agreement" }, ctxFor(directory))
  const match = result.match(/^=== Agreement — § 1 \(chars 0–(\d+)\) ===\n/)
  expect(match).not.toBeNull()
  const end = Number(match![1])
  expect(end).toBeLessThanOrEqual(20000)
  expect(chunks.some((chunk) => chunk.char_start + chunk.text.length === end)).toBe(true)
  const body = result.slice(match![0].length, result.indexOf("\n\n[Truncated"))
  expect(body).toBe(long.slice(0, end))
  expect(result).toContain(`on a chunk boundary before Agreement § 1 (char ${end})`)
})
