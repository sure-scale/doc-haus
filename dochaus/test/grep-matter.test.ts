// grep-matter lists every exact-text match once, at its offset in the
// document's text. Run from dochaus/: bun test test/grep-matter.test.ts
import { Database } from "bun:sqlite"
import { afterEach, expect, test } from "bun:test"
import { mkdirSync, mkdtempSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"

// Every fixture below is invented test content: no actual person,
// organization, transaction or obligation.

const grepMatter = (await import("../tool/grep-matter")).default

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

type Row = { doc_name: string; section: string; text: string; char_start: number }

function matter(rows: Row[]) {
  const directory = mkdtempSync(path.join(tmpdir(), "grep-matter-"))
  directories.push(directory)
  mkdirSync(path.join(directory, ".dochaus"))
  const db = new Database(path.join(directory, ".dochaus", "legal.db"))
  db.run(
    "CREATE TABLE chunks (id INTEGER PRIMARY KEY, doc_path TEXT, doc_name TEXT, section TEXT, text TEXT, char_start INTEGER, char_end INTEGER)",
  )
  for (const row of rows)
    db.query("INSERT INTO chunks (doc_path, doc_name, section, text, char_start, char_end) VALUES (?, ?, ?, ?, ?, ?)").run(
      `/m/${row.doc_name}.docx`,
      row.doc_name,
      row.section,
      row.text,
      row.char_start,
      row.char_start + row.text.length,
    )
  db.close()
  return directory
}

const grep = (directory: string, args: { pattern: string; regex?: boolean; ignoreCase?: boolean; document?: string }) =>
  grepMatter.execute(args as never, ctxFor(directory) as never) as Promise<string>

// One section stored as two chunks that share `overlap` characters, as ingest
// repeats the end of a long section's chunk at the start of the next.
function overlapping(docName: string, section: string, text: string, split: number, overlap: number): Row[] {
  return [
    { doc_name: docName, section, text: text.slice(0, split), char_start: 0 },
    { doc_name: docName, section, text: text.slice(split - overlap), char_start: split - overlap },
  ]
}

const filler = (label: string, count: number) =>
  Array.from({ length: count }, (_, index) => `Line ${index + 1} of invented clause ${label} describes no actual obligation.\n`).join("")

test("lists a match inside two chunks' overlap once, at its offset in the document", async () => {
  const text = `1. Scope\n${filler("one", 6)}The Invented Widget Fee is payable monthly.\n${filler("two", 6)}`
  const at = text.indexOf("Invented Widget Fee")
  // The match lies wholly inside the 200 characters both chunks hold.
  const split = at + 120
  const rows = overlapping("Agreement", "1. Scope", text, split, 200)
  expect(rows[0]!.text).toContain("Invented Widget Fee")
  expect(rows[1]!.text).toContain("Invented Widget Fee")
  const directory = matter(rows)

  const result = await grep(directory, { pattern: "invented widget fee" })
  expect(result).toBe(`1 match(es) for "invented widget fee":\nAgreement [1. Scope] @ ${at}: The Invented Widget Fee is payable monthly.`)

  const regex = await grep(directory, { pattern: "Widget\\s+Fee", regex: true, ignoreCase: false })
  expect(regex).toBe(`1 match(es) for regex /Widget\\s+Fee/:\nAgreement [1. Scope] @ ${at + "Invented ".length}: The Invented Widget Fee is payable monthly.`)
})

test("lists matches in document order, then by offset, each at its exact offset", async () => {
  const agreement = `1. Scope\nThe Example Supplier delivers.\n${filler("one", 6)}The Example Supplier invoices.\n${filler("two", 6)}`
  const schedule = "Schedule 1\nThe Example Supplier attends.\n"
  const split = agreement.indexOf("The Example Supplier invoices.") - 40
  const directory = matter([
    { doc_name: "Schedule", section: "Schedule 1", text: schedule, char_start: 0 },
    ...overlapping("Agreement", "1. Scope", agreement, split, 150),
  ])
  const first = agreement.indexOf("Example Supplier")
  const second = agreement.indexOf("Example Supplier", first + 1)
  expect(await grep(directory, { pattern: "Example Supplier" })).toBe(
    [
      `3 match(es) for "Example Supplier":`,
      `Agreement [1. Scope] @ ${first}: The Example Supplier delivers.`,
      `Agreement [1. Scope] @ ${second}: The Example Supplier invoices.`,
      `Schedule [Schedule 1] @ ${schedule.indexOf("Example Supplier")}: The Example Supplier attends.`,
    ].join("\n"),
  )
})

test("finds a match that straddles two chunks, attributed to the chunk it starts in", async () => {
  const text = "1. Payment\nThe invented fee is due on the first day.\n2. Term\nThe invented term is one year.\n"
  const boundary = text.indexOf("2. Term")
  const directory = matter([
    { doc_name: "Agreement", section: "1. Payment", text: text.slice(0, boundary), char_start: 0 },
    { doc_name: "Agreement", section: "2. Term", text: text.slice(boundary), char_start: boundary },
  ])
  const at = text.indexOf("first day.\n2. Term")
  expect(await grep(directory, { pattern: "first day\\.\\s+2\\. Term", regex: true })).toBe(
    `1 match(es) for regex /first day\\.\\s+2\\. Term/i:\nAgreement [1. Payment] @ ${at}: The invented fee is due on the first day.`,
  )
})

test("keeps case-insensitive offsets exact when lower-casing would change the text's length", async () => {
  // "İ" lower-cases to two UTF-16 code units, which would shift every later offset.
  const text = "1. Parties\nİnvented Holdings and the Sample Lessee agree.\n"
  const directory = matter([{ doc_name: "Agreement", section: "1. Parties", text, char_start: 0 }])
  expect(await grep(directory, { pattern: "sample lessee" })).toBe(
    `1 match(es) for "sample lessee":\nAgreement [1. Parties] @ ${text.indexOf("Sample Lessee")}: İnvented Holdings and the Sample Lessee agree.`,
  )
})
