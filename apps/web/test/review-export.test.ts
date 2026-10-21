import { expect, test } from "bun:test"
import { CFB, read, utils } from "xlsx"
import { hashQuestion, isStale, type Grid, type GridRecheck } from "../src/api/ingest/grid"
import { describeRecheck } from "../src/api/review-recheck"
import { reviewCsv, reviewDocx, reviewExport, reviewXlsx, type ReviewExport } from "../src/components/review-export"
import { formatDate } from "../src/lib/format"

const at = Date.UTC(2026, 9, 11, 12)
const recheck: GridRecheck = { outcome: "not_found", passages: 12, read: 4, candidates: 0, reread: 0, terms: ["break clause", "break"], channels: ["lexical", "phrase"] }
const citation = { documentName: "Unit 4 lease.docx", docPath: "/matter/Unit 4 lease.docx", section: "9.2", excerpt: "This lease is governed by the law of England and Wales.", charStart: 10, charEnd: 64 }
const q1 = "What is the governing law?"
const q2 = "Is there a tenant break clause?"
// The reviewed cell and the supplier letter's q2 answer were computed for an
// earlier wording of their question; only the unlocked one is stale.
const grid: Grid = {
  columns: [{ id: "q1", question: q1 }, { id: "q2", question: q2 }],
  cells: {
    "Unit 4 lease.docx::q1": { answer: "England and Wales", citation, status: "reviewed", questionHash: "earlier", comments: [{ id: "c1", text: "Confirmed against the signed copy", at }] },
    "Unit 4 lease.docx::q2": { answer: "Not addressed", status: "filled", questionHash: hashQuestion(q2), recheck },
    "Supplier letter.docx::q2": { answer: "Not addressed", status: "filled", questionHash: "earlier" },
  },
}
const model = reviewExport("Aldgate Mills", grid, ["Unit 4 lease.docx", "Supplier letter.docx"])
const comment = `Comment (${formatDate(at)}): Confirmed against the signed copy`
const stale = "Stale: question changed since this answer"

test("an answer is stale when its question changed after it was computed, unless it is reviewed", () => {
  expect(isStale(grid.cells["Unit 4 lease.docx::q2"], q2)).toBe(false)
  expect(isStale(grid.cells["Supplier letter.docx::q2"], q2)).toBe(true)
  expect(isStale(grid.cells["Unit 4 lease.docx::q1"], q1)).toBe(false)
  expect(isStale(undefined, q1)).toBe(false)
})

test("one model carries questions, documents, answers, sources, reviewed and stale state, re-check notes and comments", () => {
  expect(model).toEqual({
    title: "Aldgate Mills",
    questions: ["What is the governing law?", "Is there a tenant break clause?"],
    rows: [
      {
        document: "Unit 4 lease.docx",
        cells: [
          { answer: "England and Wales", source: "Unit 4 lease.docx § 9.2", reviewed: true, stale: false, comments: [comment] },
          { answer: "Not addressed", recheck: describeRecheck(recheck), reviewed: false, stale: false, comments: [] },
        ],
      },
      {
        document: "Supplier letter.docx",
        cells: [
          { answer: "", reviewed: false, stale: false, comments: [] },
          { answer: "Not addressed", reviewed: false, stale: true, comments: [] },
        ],
      },
    ],
  })
  expect(model.rows[0].cells[1].recheck).toStartWith("Not found after re-check.")
})

// --- CSV ---------------------------------------------------------------------

const parseCsv = (csv: string) => utils.sheet_to_json<string[]>(read(csv, { type: "string", raw: true }).Sheets.Sheet1, { header: 1, raw: true, defval: "" })

test("CSV is UTF-8 with a BOM and CRLF records, one row per document and question", () => {
  const csv = reviewCsv(model)
  expect(csv.startsWith("﻿")).toBe(true)
  const records = csv.slice(1).split("\r\n")
  expect(records[0]).toBe("Document,Question,Answer,Source,Reviewed,Stale,Re-check,Comments")
  expect(records[1]).toBe(`Unit 4 lease.docx,What is the governing law?,England and Wales,Unit 4 lease.docx § 9.2,Yes,No,,${comment}`)
  expect(records.at(-1)).toBe("")
  expect(parseCsv(csv.slice(1))).toEqual([
    ["Document", "Question", "Answer", "Source", "Reviewed", "Stale", "Re-check", "Comments"],
    ["Unit 4 lease.docx", q1, "England and Wales", "Unit 4 lease.docx § 9.2", "Yes", "No", "", comment],
    ["Unit 4 lease.docx", q2, "Not addressed", "", "No", "No", describeRecheck(recheck), ""],
    ["Supplier letter.docx", q1, "", "", "No", "No", "", ""],
    ["Supplier letter.docx", q2, "Not addressed", "", "No", "Yes", "", ""],
  ])
})

const one = (document: string, question: string, answer: string, comments: string[] = []): ReviewExport => ({
  title: "T", questions: [question], rows: [{ document, cells: [{ answer, reviewed: false, stale: false, comments }] }],
})

test("CSV quotes fields holding commas, quotes or line breaks (RFC 4180)", () => {
  const csv = reviewCsv(one('Smith, Jones "draft".docx', "Who signs?", "Line one\nLine two", ["Comment (1 Jan 2026): a", "Comment (2 Jan 2026): b"]))
  expect(csv.slice(1).split("\r\n")[1]).toBe('"Smith, Jones ""draft"".docx",Who signs?,"Line one\nLine two",,No,No,,"Comment (1 Jan 2026): a\nComment (2 Jan 2026): b"')
})

test("CSV neutralises cells a spreadsheet would read as a formula", () => {
  for (const [value, expected] of [
    ['=HYPERLINK("http://x.test","open")', `"'=HYPERLINK(""http://x.test"",""open"")"`],
    ["+44 20 7946 0000", "'+44 20 7946 0000"],
    ["-5% of the price", "'-5% of the price"],
    ["@SUM(A1:A2)", "'@SUM(A1:A2)"],
    ["\tindented", "'\tindented"],
    ["\rreturn", `"'\rreturn"`],
  ]) {
    expect(reviewCsv(one(value, value, value)).slice(1).split("\r\n")[1]).toBe(`${expected},${expected},${expected},,No,No,,`)
  }
  expect(reviewCsv(one("A=B.docx", "Net 30?", "1 + 1")).slice(1).split("\r\n")[1]).toBe("A=B.docx,Net 30?,1 + 1,,No,No,,")
})

// --- Word --------------------------------------------------------------------

function unzip(bytes: ArrayBuffer) {
  const zip = CFB.read(new Uint8Array(bytes), { type: "array" })
  return (name: string) => {
    const index = zip.FullPaths.findIndex((path: string) => path.endsWith(`/${name}`))
    return index < 0 ? undefined : new TextDecoder().decode(zip.FileIndex[index].content)
  }
}
const decode = (text: string) => text.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&")
const paragraphs = (xml: string) => [...xml.matchAll(/<w:p[ >][\s\S]*?<\/w:p>/g)].map((p) => decode([...p[0].matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map((t) => t[1]).join("")))
const rows = (xml: string) => [...xml.matchAll(/<w:tr[ >][\s\S]*?<\/w:tr>/g)].map((tr) => [...tr[0].matchAll(/<w:tc>[\s\S]*?<\/w:tc>/g)].map((tc) => paragraphs(tc[0])))

test("Word export is one table with the grid's content, header row repeated on each page", async () => {
  const xml = unzip(await reviewDocx(model))("word/document.xml")!
  expect(paragraphs(xml)[0]).toBe("Tabular review — Aldgate Mills")
  expect(xml.match(/<w:tbl>/g)).toHaveLength(1)
  expect(xml).toContain("<w:tblHeader/>")
  expect(rows(xml)).toEqual([
    [["Document"], ["What is the governing law?"], ["Is there a tenant break clause?"]],
    [["Unit 4 lease.docx"], ["England and Wales", "Reviewed", "Source: Unit 4 lease.docx § 9.2", comment], ["Not addressed", describeRecheck(recheck)]],
    [["Supplier letter.docx"], [""], ["Not addressed", stale]],
  ])
  expect(xml).not.toContain('w:orient="landscape"')
})

test("Word export turns landscape for three or more questions and keeps answer line breaks", async () => {
  const wide: ReviewExport = {
    title: "Wide",
    questions: ["Q1", "Q2", "Q3"],
    rows: [{ document: "Lease.docx", cells: [{ answer: "First line\nSecond line", reviewed: false, stale: false, comments: [] }, { answer: "a & <b>", reviewed: false, stale: false, comments: [] }, { answer: "", reviewed: false, stale: false, comments: [] }] }],
  }
  const xml = unzip(await reviewDocx(wide))("word/document.xml")!
  expect(xml).toMatch(/<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"\/>/)
  expect(rows(xml)[1]).toEqual([["Lease.docx"], ["First line", "Second line"], ["a & <b>"], [""]])
})

test("Word export drops characters XML cannot carry, so a stray form feed cannot corrupt the file", async () => {
  const dirty: ReviewExport = {
    title: "Mills\u0007",
    questions: ["Term?\u000B"],
    rows: [{ document: "Lease\u0000.docx", cells: [{ answer: "Five\fyears\u0001 from\ttoday\uFFFE", reviewed: false, stale: false, comments: ["Comment (1 Jan 2026): see\u001F p. 4"] }] }],
  }
  const xml = unzip(await reviewDocx(dirty))("word/document.xml")!
  expect(xml).not.toMatch(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/)
  expect(paragraphs(xml)[0]).toBe("Tabular review — Mills")
  expect(rows(xml)).toEqual([[["Document"], ["Term?"]], [["Lease.docx"], ["Fiveyears from\ttoday", "Comment (1 Jan 2026): see p. 4"]]])
})

// --- Excel -------------------------------------------------------------------

test("Excel export keeps the grid layout: answers in cells, reviewed or stale state, source, re-check and comments as a hidden note", () => {
  const bytes = reviewXlsx(model)
  const sheet = read(bytes, { type: "array" }).Sheets["Tabular review"]
  expect(utils.sheet_to_json<string[]>(sheet, { header: 1, defval: "" })).toEqual([
    ["Document", "What is the governing law?", "Is there a tenant break clause?"],
    ["Unit 4 lease.docx", "England and Wales", "Not addressed"],
    ["Supplier letter.docx", "", "Not addressed"],
  ])
  expect(sheet.B2.c.map((note: { a: string; t: string }) => [note.a, note.t])).toEqual([["doc.haus", `Reviewed\nSource: Unit 4 lease.docx § 9.2\n${comment}`]])
  expect(sheet.C2.c.map((note: { t: string }) => note.t)).toEqual([describeRecheck(recheck)])
  expect(sheet.C3.c.map((note: { t: string }) => note.t)).toEqual([stale])
  expect(sheet.B3?.c).toBeUndefined()
  expect(unzip(bytes)("xl/drawings/vmlDrawing1.vml")).toContain("visibility:hidden")
})
