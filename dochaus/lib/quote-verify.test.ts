import { afterEach, expect, test } from "bun:test"
import { Database } from "bun:sqlite"
import { mkdirSync, mkdtempSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { elidedParts, partsInOrder, verifyQuote, type QuoteChunk } from "./quote-verify"

const directories: string[] = []
afterEach(() => directories.splice(0).forEach(directory => rmSync(directory, { recursive: true, force: true })))
function fixture(chunks: QuoteChunk[]) {
  const directory = mkdtempSync(path.join(tmpdir(), "quote-chunks-")); directories.push(directory)
  mkdirSync(path.join(directory, ".dochaus"))
  const db = new Database(path.join(directory, ".dochaus/legal.db"))
  db.run("CREATE TABLE chunks (doc_name TEXT, section TEXT, text TEXT, char_start INTEGER, char_end INTEGER)")
  for (const chunk of chunks) db.query("INSERT INTO chunks VALUES (?, ?, ?, ?, ?)").run("original.docx", chunk.section, chunk.text, chunk.char_start, chunk.char_end)
  db.close()
  return (quote: string, document = "original.docx") => verifyQuote({ directory, quote, document })
}
function chunk(text: string, start: number, section: string): QuoteChunk { return { text, char_start: start, char_end: start + text.length, section } }

test("overlapping chunks verify the original cross-boundary text without duplicating overlap", () => {
  const check = fixture([chunk("alpha beta ", 0, "first"), chunk("beta gamma delta", 6, "second")])
  expect(check("alpha beta gamma delta")).toMatchObject({ found: true, hits: [{ charStart: 0, charEnd: 22, verbatim: "alpha beta gamma delta" }] })
  expect(check("alpha beta beta gamma delta").found).toBe(false)
  expect(check("alpha beta changed delta").found).toBe(false)
  expect(check("alpha beta gamma delta", "another.docx")).toMatchObject({ found: false, reason: "no-docs" })
})

test("contained chunks neither duplicate the source nor truncate evidence when the last chunk ends earlier", () => {
  const text = "alpha beta gamma delta epsilon"
  const check = fixture([chunk(text.slice(0, 11), 0, "first"), chunk(text.slice(6), 6, "long"), chunk(text.slice(17, 22), 17, "contained")])
  expect(check(text)).toMatchObject({ found: true, hits: [{ charStart: 0, charEnd: text.length, verbatim: text }] })
  expect(check("alpha beta gamma delta delta epsilon").found).toBe(false)
})

test("whitespace gaps retain source offsets and existing normalization including Markdown source escapes", () => {
  const text = '“Alpha”\\.\n\nBeta gamma delta'
  const split = text.indexOf("Beta")
  const check = fixture([chunk(text.slice(0, split - 2), 0, "first"), chunk(text.slice(split, split + 10), split, "second"), chunk(text.slice(split + 5), split + 5, "overlap")])
  expect(check('"Alpha". Beta gamma delta')).toMatchObject({ found: true, hits: [{ charStart: 0, charEnd: text.length }] })
  expect(check('"Alpha". Beta changed delta').found).toBe(false)
})

test("a quotation matched without regard to case still reads a currency code as written", () => {
  const request = (amount: string) => `Please cap the liability at ${amount} the charges paid in the year before the claim.`
  const parts = elidedParts("please cap the liability at … the charges paid in the year before the claim")!
  expect(partsInOrder(request("the greater of GBP 500000 and"), parts, { ignoreCase: true })).toBe(false)
  expect(partsInOrder(request("the amount, for 30 days, of"), parts, { ignoreCase: true })).toBe(true)
})
