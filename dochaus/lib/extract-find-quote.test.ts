import { afterEach, expect, test } from "bun:test"
import { Database } from "bun:sqlite"
import { mkdirSync, mkdtempSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { findQuote } from "./extract"
import { verifyQuote } from "./quote-verify"

// findQuote matches a quotation through the readings verify-quote uses for the
// same document (sourceReadings: curly quotes, whitespace, escapes, and only the
// markup the document's format has) but returns the raw span, so a stored
// excerpt is still a slice of the live text. Synthetic text only.

const directories: string[] = []
afterEach(() => directories.splice(0).forEach(directory => rmSync(directory, { recursive: true, force: true })))

function anchored(text: string, quote: string, document?: string, near?: number) {
  const match = findQuote(text, quote, document, near)
  if (match) expect(text.slice(match.start, match.end)).toBe(match.excerpt)
  return match
}

/** verifyQuote over a one-chunk index of `text` named `name`. */
function verified(name: string, text: string, quote: string) {
  const directory = mkdtempSync(path.join(tmpdir(), "find-quote-")); directories.push(directory)
  mkdirSync(path.join(directory, ".dochaus"))
  const db = new Database(path.join(directory, ".dochaus/legal.db"))
  db.run("CREATE TABLE chunks (doc_name TEXT, section TEXT, text TEXT, char_start INTEGER, char_end INTEGER)")
  db.query("INSERT INTO chunks VALUES (?, ?, ?, ?, ?)").run(name, "s", text, 0, text.length)
  db.close()
  return verifyQuote({ directory, quote, document: name }).found
}

test("a straight-quoted quotation anchors to the document's curly quotes, as written", () => {
  const text = "1.1 “Synthetic Price” means the sum stated in Schedule 2.\n"
  expect(anchored(text, '"Synthetic Price" means the sum')?.excerpt).toBe("“Synthetic Price” means the sum")
  expect(findQuote(text, '"Synthetic Cost" means the sum')).toBeUndefined()
})

test("a quotation anchors past the source's Markdown escapes and, in Markdown, its bold marks", () => {
  const text = "The Buyer signs here: \\_\\_\\_\\_ on the date below.\n**Deposit:** payable within ten days.\n"
  expect(anchored(text, "signs here: ____ on the date")?.excerpt).toBe("signs here: \\_\\_\\_\\_ on the date")
  expect(anchored(text, "Deposit: payable within ten days", "terms.md")?.excerpt).toBe("Deposit:** payable within ten days")
  expect(findQuote(text, "Deposit: payable within ten days", "terms.docx")).toBeUndefined()
  expect(findQuote(text, "signs here: ____ on the day")).toBeUndefined()
  expect(findQuote(text, "Deposit: payable within five days", "terms.md")).toBeUndefined()
})

test("findQuote accepts exactly what verifyQuote accepts in the same document", () => {
  const csv = '## Terms\nItem,Term\nPrice,"The ""Buyer"" pays within 30 days, net"\n'
  const link = "The Buyer shall [not](#opt) assign this Agreement.\n"
  const reply = "Agreed.\n\n&gt; The Seller will deliver\n&gt; on Friday.\n"
  const cases: [name: string, text: string, quote: string, found: boolean][] = [
    // A spreadsheet's CSV quoting is read past.
    ["Terms.xlsx", csv, 'The "Buyer" pays within 30 days, net', true],
    // A Word document's "[not](#opt)" is what the page shows, not a link.
    ["Agreement.docx", link, "The Buyer shall not assign", false],
    ["Agreement.md", link, "The Buyer shall not assign", true],
    // "&gt; " is a reply mark only in an email.
    ["Note.docx", reply, "The Seller will deliver on Friday", false],
    ["Mail.eml", reply, "The Seller will deliver on Friday", true],
  ]
  for (const [name, text, quote, found] of cases) {
    expect([name, verified(name, text, quote)]).toEqual([name, found])
    expect([name, anchored(text, quote, name) !== undefined]).toEqual([name, found])
  }
})

test("a match ending on an entity keeps the whole entity in its excerpt", () => {
  const text = "Hello,\nthe &quot;Price&quot; is fixed.\n"
  expect(anchored(text, 'the "Price"', "Mail.eml")?.excerpt).toBe("the &quot;Price&quot;")
})

test("a literal occurrence wins; otherwise the tolerant one nearest `near`", () => {
  const mixed = "“Synthetic Term” first. Then \"Synthetic Term\" again.\n"
  expect(anchored(mixed, '"Synthetic Term"')?.start).toBe(mixed.indexOf('"Synthetic Term"'))
  const curly = "“Synthetic Term” first. Later “Synthetic Term” again.\n"
  const last = curly.lastIndexOf("“Synthetic Term”")
  expect(anchored(curly, '"Synthetic Term"')?.start).toBe(0)
  expect(anchored(curly, '"Synthetic Term"', undefined, last - 2)?.start).toBe(last)
})
