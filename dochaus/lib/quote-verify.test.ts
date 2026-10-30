import { afterEach, expect, test } from "bun:test"
import { Database } from "bun:sqlite"
import { mkdirSync, mkdtempSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { elidedParts, elisionGap, locateElidedQuote, normalizeWithMap, sourceReadings, verifyQuote, type QuoteChunk } from "./quote-verify"

const directories: string[] = []
afterEach(() => directories.splice(0).forEach(directory => rmSync(directory, { recursive: true, force: true })))
function fixture(chunks: QuoteChunk[], name = "original.docx") {
  const directory = mkdtempSync(path.join(tmpdir(), "quote-chunks-")); directories.push(directory)
  mkdirSync(path.join(directory, ".dochaus"))
  const db = new Database(path.join(directory, ".dochaus/legal.db"))
  db.run("CREATE TABLE chunks (doc_name TEXT, section TEXT, text TEXT, char_start INTEGER, char_end INTEGER)")
  for (const chunk of chunks) db.query("INSERT INTO chunks VALUES (?, ?, ?, ?, ?)").run(name, chunk.section, chunk.text, chunk.char_start, chunk.char_end)
  db.close()
  return (quote: string, document = name) => verifyQuote({ directory, quote, document })
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

test("a Word document's escapes are read past alone, its other marks as the page shows them", () => {
  const check = fixture([chunk("- **Fees:** 2\\.5% per annum. See [Schedule 2](#s2), clause 2\\.1. Smith &amp; Jones LLP\\. agree.", 0, "first")])
  expect({
    bold: check("**Fees:** 2.5% per annum").found,
    link: check("See [Schedule 2](#s2), clause 2.1").found,
    entity: check("Smith &amp; Jones LLP. agree").found,
    boldLeftOut: check("Fees: 2.5% per annum").found,
    linkLeftOut: check("See Schedule 2, clause 2.1").found,
    entityDecoded: check("Smith & Jones LLP. agree").found,
  }).toEqual({ bold: true, link: true, entity: true, boldLeftOut: false, linkLeftOut: false, entityDecoded: false })
})

test("Markdown bold marks around words are not part of the quoted words, but unpaired or joined stars are", () => {
  const text = "- **Notice Period:** thirty days after delivery of the notice"
  const check = fixture([chunk(text, 0, "first")], "terms.md")
  expect(check("Notice Period: thirty days after delivery")).toMatchObject({ found: true, hits: [{ charStart: 0, charEnd: text.length }] })
  expect(check("Notice Period: sixty days after delivery").found).toBe(false)
  expect(fixture([chunk("a literal \\*\\*Notice", 0, "first")], "terms.md")("a literal **Notice").found).toBe(true)
  const stars = fixture([chunk("The cap is £10,000,000** in total. Use 10**6 or 12**34 units. Cap**: £1m; Basket**: £100k. The Buyer shall pay *** within ten days.", 0, "first")], "terms.md")
  expect({
    footnote: stars("The cap is £10,000,000 in total").found,
    figures: stars("Use 106 or 1234 units").found,
    labels: stars("Cap: £1m; Basket: £100k").found,
    blank: stars("The Buyer shall pay within ten days").found,
    word: fixture([chunk(text, 0, "first")])("Notice Period: thirty days after delivery").found,
  }).toEqual({ footnote: false, figures: false, labels: false, blank: false, word: false })
})

test("single emphasis marks and link targets are not part of the quoted words, but placeholder brackets are", () => {
  const text = "The Lender shall *not* assign it. See [Annex 2](#annex-2) and [insert date], [●]. Under clause [4](a) only. A \\*literal\\* mark. " +
    "Year 1: 5%*; Year 2: 7%*. Cap (per claim)*: £1m; Basket (aggregate)*: £100k. Price [£5m](TBC:board). Fee £5[1](#fn1),000."
  const check = fixture([chunk(text, 0, "first")], "terms.md")
  expect({
    emphasis: check("The Lender shall not assign it").found,
    emphasisDropped: check("The Lender shall assign it").found,
    link: check("See Annex 2 and [insert date], [●]").found,
    linkChanged: check("See Annex 3 and [insert date]").found,
    placeholder: check("and insert date").found,
    reference: check("Under clause 4 only").found,
    escaped: check("A *literal* mark").found,
    footnotes: check("Year 1: 5%; Year 2: 7%").found,
    bracketFootnotes: check("Cap (per claim): £1m; Basket (aggregate): £100k").found,
    pending: check("Price £5m.").found,
    inFigure: check("Fee £5,000.").found,
    word: fixture([chunk(text, 0, "first")])("The Lender shall not assign it").found,
  }).toEqual({
    emphasis: true, emphasisDropped: false, link: true, linkChanged: false, placeholder: false, reference: false, escaped: true,
    footnotes: false, bracketFootnotes: false, pending: false, inFigure: false, word: false,
  })
})

test("a rule line between PDF pages reads as a space, in whole and elided quotations alike, but not in a Word document or a form", () => {
  const text = "The Tenant shall pay within fourteen days of the date of\n\n-----\n\nthe invoice, without set-off. Signed for the Tenant:\n---\nDate of signature."
  const check = fixture([chunk(text, 0, "first")], "statement.pdf")
  const locate = (quote: string) => locateElidedQuote({ directory: directories.at(-1)!, parts: elidedParts(quote)!, document: "statement.pdf" })
  expect({
    whole: check("within fourteen days of the date of the invoice").found,
    changed: check("within fourteen days of the date of an invoice").found,
    passage: locate("… within fourteen days of the date of the invoice"),
    passageChanged: locate("… within twenty days of the date of the invoice"),
    joined: locate("within fourteen days of the date of … the invoice, without set-off"),
    formLine: check("Signed for the Tenant: Date of signature").found,
    word: fixture([chunk(text, 0, "first")])("within fourteen days of the date of the invoice").found,
  }).toEqual({
    whole: true, changed: false, passage: { kind: "near", omitted: "" }, passageChanged: null, joined: { kind: "near", omitted: "" }, formLine: false, word: false,
  })
})

test("reading long runs of blank lines and spaces takes linear time", () => {
  const started = performance.now()
  for (const reading of [...sourceReadings("statement.pdf"), ...sourceReadings("thread.eml")]) normalizeWithMap(`${"\n".repeat(200_000)}${" ".repeat(200_000)}x`, reading)
  expect(performance.now() - started).toBeLessThan(1000)
})

test("hyphens and dashes a page shows as a hyphen match a hyphen on either side", () => {
  const check = fixture([chunk("The Licensor grants a non\u2011exclusive licence for 2024\u20132026 only.", 0, "first")])
  expect({
    hyphen: check("grants a non-exclusive licence for 2024-2026").found,
    dash: check("grants a non\u2010exclusive licence for 2024\u22122026").found,
    changed: check("grants an exclusive licence for 2024-2026").found,
    year: check("grants a non-exclusive licence for 2024-2027").found,
    emDash: check("grants a non\u2014exclusive licence").found,
  }).toEqual({ hyphen: true, dash: true, changed: false, year: false, emDash: false })
})

test("reply marks opening an email's lines and its HTML entities are not part of the quoted words", () => {
  const text = [
    "On Monday the Buyer wrote:", "&gt; The price is reduced by the amount of", "&gt; any Leakage before Completion.",
    "> Agreed by the Seller", "> in writing.", "See the Data Room &lt;https://example.com/dr&gt; for R&amp;D costs. A fee > 500 applies.",
    "The Buyer&#8217;s&nbsp;consent &ndash; given on &#x2019;signing&rsquo; &mdash; is final.",
  ].join("\n")
  const check = fixture([chunk(text, 0, "first")], "thread.eml")
  expect({
    entityReply: check("reduced by the amount of any Leakage before Completion").found,
    entityReplyChanged: check("reduced by the amount of Leakage before Completion").found,
    reply: check("Agreed by the Seller in writing").found,
    replyChanged: check("Agreed by the Buyer in writing").found,
    decoded: check("See the Data Room <https://example.com/dr> for R&D costs").found,
    asWritten: check("See the Data Room &lt;https://example.com/dr&gt; for R&amp;D costs").found,
    linkLeftOut: check("See the Data Room for R&D costs").found,
    inLine: check("A fee 500 applies").found,
    entities: check("The Buyer\u2019s consent \u2013 given on 'signing' \u2014 is final").found,
    word: fixture([chunk(text, 0, "first")])("reduced by the amount of any Leakage before Completion").found,
  }).toEqual({
    entityReply: true, entityReplyChanged: false, reply: true, replyChanged: false, decoded: true, asWritten: true, linkLeftOut: false, inLine: false,
    entities: true, word: false,
  })
})

test("a quotation never runs across a change of reply depth, nor drops a Word document's reply marks", () => {
  const thread = fixture([chunk("Agreed, subject to the following:\n&gt; The Seller will give no tax indemnity.", 0, "first")], "thread.eml")
  const word = fixture([chunk("The ratchet is:\n&gt; £5,000,000: 100%", 0, "first")])
  expect({
    across: thread("Agreed, subject to the following: The Seller will give no tax indemnity").found,
    reply: thread("Agreed, subject to the following:").found,
    quoted: thread("The Seller will give no tax indemnity.").found,
    wordAsWritten: word("The ratchet is: &gt; £5,000,000: 100%").found,
    wordLeftOut: word("The ratchet is: £5,000,000: 100%").found,
  }).toEqual({ across: false, reply: true, quoted: true, wordAsWritten: true, wordLeftOut: false })
})

test("a match ending in an entity a chunk boundary splits keeps the whole entity as its evidence", () => {
  const text = "See the Data Room &lt;dr&gt; for the costs."
  const split = text.indexOf("&gt;") + 2
  const check = fixture([chunk(text.slice(0, split), 0, "first"), chunk(text.slice(split), split, "second")], "thread.eml")
  expect(check("See the Data Room <dr>")).toMatchObject({ found: true, hits: [{ charStart: 0, charEnd: text.length, verbatim: text }] })
})

test("a spreadsheet's cells are quoted as the sheet shows them, without the CSV quoting, but not changed or run together", () => {
  const text = [
    "## Escrow", "Item,Amount,Status", 'Escrow,"1,000,000",held', 'Fees,"2,000",paid', 'Note,"The Buyer called it ""final"" today",open',
    'Split,"5,000",250', 'Terms,"Pay on signing', 'and on completion",due',
  ].join("\n")
  const second = text.indexOf("Fees"), third = text.indexOf("and on completion")
  const docx = fixture([chunk(text, 0, "first")])
  const check = fixture([chunk(text.slice(0, third), 0, "Escrow"), chunk(text.slice(second), second, "Escrow")], "escrow.xlsx")
  const locate = (quote: string) => locateElidedQuote({ directory: directories.at(-1)!, parts: elidedParts(quote)!, document: "escrow.xlsx" })
  expect({
    asWritten: check('Escrow,"1,000,000",held').found,
    row: check("Escrow,1,000,000,held").found,
    rows: check("Escrow,1,000,000,held Fees,2,000,paid").found,
    amountChanged: check("Escrow,1,000,001,held").found,
    cellChanged: check("Escrow,1,000,000,released").found,
    otherRow: check("Escrow,2,000,held").found,
    doubled: check('The Buyer called it "final" today').found,
    doubledDropped: check("The Buyer called it final today").found,
    runTogether: check("Split,5,000,250").found,
    lineBreak: check("Terms,Pay on signing and on completion,due").found,
    passage: locate("Escrow,1,000,000,held Fees …"),
    passageChanged: locate("Escrow,1,000,001,held Fees …"),
    notSpreadsheet: docx("Escrow,1,000,000,held").found,
  }).toEqual({
    asWritten: true, row: true, rows: true, amountChanged: false, cellChanged: false, otherRow: false, doubled: true, doubledDropped: false,
    runTogether: false, lineBreak: true, passage: { kind: "near", omitted: "" }, passageChanged: null, notSpreadsheet: false,
  })
  // A match one chunk holds is that chunk's evidence, even where chunks overlap.
  expect(check("Fees,2,000,paid")).toMatchObject({ found: true, hits: [{ charStart: 0, charEnd: third }] })
  expect(check("and on completion,due")).toMatchObject({ found: true, hits: [{ charStart: second, charEnd: text.length }] })
})

test("chunks stored with an extra newline past a split section keep the next heading whole", () => {
  // Chunks ingested before the chunker stopped counting one newline past a
  // split section: the section's last chunk ends with "\n" where the next
  // heading's first character is, and the next chunk's text replaces it.
  const text = "1. First\nAn invented clause, describing no actual obligation.\n2. Next clause\nMore invented text.\n"
  const second = text.indexOf("describing")
  const next = text.indexOf("2. Next")
  const check = fixture([chunk(text.slice(0, second), 0, "1"), chunk(`${text.slice(second, next)}\n`, second, "1"), chunk(text.slice(next), next, "2")])
  const directory = directories.at(-1)!
  expect({
    verified: check("describing no actual obligation. 2. Next clause"),
    located: locateElidedQuote({ directory, parts: elidedParts("An invented clause, describing … obligation. 2. Next clause")!, document: "original.docx" }),
  }).toMatchObject({
    verified: { found: true, hits: [{ charStart: second, charEnd: text.length, verbatim: text.slice(second) }] },
    located: { kind: "near", omitted: "no actual" },
  })
})

test("an elided quotation located without regard to case or Markdown escapes shows the words left out as written", () => {
  const request = "Please cap the liability at the greater of GBP 500,000 and the charges paid in the year before the claim."
  const parts = elidedParts("please cap the liability at … the charges paid in the year before the claim")!
  expect(elisionGap(request, parts)).toBeNull()
  expect(elisionGap(request, parts, { ignoreCase: true })).toEqual({ kind: "near", omitted: "the greater of GBP 500,000 and" })
  expect(elisionGap("The cap in clause 4\\.2 applies \\(save for fraud\\) to all goods and services.", elidedParts("The cap in clause 4.2 applies … to all goods and services")!, { escapes: true }))
    .toEqual({ kind: "near", omitted: "(save for fraud)" })
})
