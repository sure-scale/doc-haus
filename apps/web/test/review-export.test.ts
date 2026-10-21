import { expect, test } from "bun:test"
import { CFB, read, utils } from "xlsx"
import type { Grid, GridRecheck } from "../src/api/ingest/grid"
import { describeRecheck } from "../src/api/review-recheck"
import { reviewCsv, reviewDocx, reviewExport, reviewXlsx, type ReviewExport } from "../src/components/review-export"
import { formatDate } from "../src/lib/format"

const at = Date.UTC(2026, 9, 11, 12)
const recheck: GridRecheck = { outcome: "not_found", passages: 12, read: 4, candidates: 0, reread: 0, terms: ["break clause", "break"], channels: ["lexical", "phrase"] }
const citation = { documentName: "Unit 4 lease.docx", docPath: "/matter/Unit 4 lease.docx", section: "9.2", excerpt: "This lease is governed by the law of England and Wales.", charStart: 10, charEnd: 64 }
const grid: Grid = {
  columns: [{ id: "q1", question: "What is the governing law?" }, { id: "q2", question: "Is there a tenant break clause?" }],
  cells: {
    "Unit 4 lease.docx::q1": { answer: "England and Wales", citation, status: "reviewed", questionHash: "h1", comments: [{ id: "c1", text: "Confirmed against the signed copy", at }] },
    "Unit 4 lease.docx::q2": { answer: "Not addressed", status: "filled", questionHash: "h2", recheck },
    "Supplier letter.docx::q2": { answer: "Not addressed", status: "filled", questionHash: "h2" },
  },
}
const model = reviewExport("Aldgate Mills", grid, ["Unit 4 lease.docx", "Supplier letter.docx"])
const comment = `Comment (${formatDate(at)}): Confirmed against the signed copy`

test("one model carries questions, documents, answers, sources, reviewed state, re-check notes and comments", () => {
  expect(model).toEqual({
    title: "Aldgate Mills",
    questions: ["What is the governing law?", "Is there a tenant break clause?"],
    rows: [
      {
        document: "Unit 4 lease.docx",
        cells: [
          { answer: "England and Wales", source: "Unit 4 lease.docx § 9.2", reviewed: true, comments: [comment] },
          { answer: "Not addressed", recheck: describeRecheck(recheck), reviewed: false, comments: [] },
        ],
      },
      {
        document: "Supplier letter.docx",
        cells: [
          { answer: "", reviewed: false, comments: [] },
          { answer: "Not addressed", reviewed: false, comments: [] },
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
  expect(records[0]).toBe("Document,Question,Answer,Source,Reviewed,Re-check,Comments")
  expect(records[1]).toBe(`Unit 4 lease.docx,What is the governing law?,England and Wales,Unit 4 lease.docx § 9.2,Yes,,${comment}`)
  expect(records.at(-1)).toBe("")
  expect(parseCsv(csv.slice(1))).toEqual([
    ["Document", "Question", "Answer", "Source", "Reviewed", "Re-check", "Comments"],
    ["Unit 4 lease.docx", "What is the governing law?", "England and Wales", "Unit 4 lease.docx § 9.2", "Yes", "", comment],
    ["Unit 4 lease.docx", "Is there a tenant break clause?", "Not addressed", "", "No", describeRecheck(recheck), ""],
    ["Supplier letter.docx", "What is the governing law?", "", "", "No", "", ""],
    ["Supplier letter.docx", "Is there a tenant break clause?", "Not addressed", "", "No", "", ""],
  ])
})

const one = (document: string, question: string, answer: string, comments: string[] = []): ReviewExport => ({
  title: "T", questions: [question], rows: [{ document, cells: [{ answer, reviewed: false, comments }] }],
})

test("CSV quotes fields holding commas, quotes or line breaks (RFC 4180)", () => {
  const csv = reviewCsv(one('Smith, Jones "draft".docx', "Who signs?", "Line one\nLine two", ["Comment (1 Jan 2026): a", "Comment (2 Jan 2026): b"]))
  expect(csv.slice(1).split("\r\n")[1]).toBe('"Smith, Jones ""draft"".docx",Who signs?,"Line one\nLine two",,No,,"Comment (1 Jan 2026): a\nComment (2 Jan 2026): b"')
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
    expect(reviewCsv(one(value, value, value)).slice(1).split("\r\n")[1]).toBe(`${expected},${expected},${expected},,No,,`)
  }
  expect(reviewCsv(one("A=B.docx", "Net 30?", "1 + 1")).slice(1).split("\r\n")[1]).toBe("A=B.docx,Net 30?,1 + 1,,No,,")
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
    [["Supplier letter.docx"], [""], ["Not addressed"]],
  ])
  expect(xml).not.toContain('w:orient="landscape"')
})

test("Word export turns landscape for three or more questions and keeps answer line breaks", async () => {
  const wide: ReviewExport = {
    title: "Wide",
    questions: ["Q1", "Q2", "Q3"],
    rows: [{ document: "Lease.docx", cells: [{ answer: "First line\nSecond line", reviewed: false, comments: [] }, { answer: "a & <b>", reviewed: false, comments: [] }, { answer: "", reviewed: false, comments: [] }] }],
  }
  const xml = unzip(await reviewDocx(wide))("word/document.xml")!
  expect(xml).toMatch(/<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"\/>/)
  expect(rows(xml)[1]).toEqual([["Lease.docx"], ["First line", "Second line"], ["a & <b>"], [""]])
})

// --- Excel (unchanged output) ------------------------------------------------

test("Excel export keeps its layout: answers in cells, source, re-check and comments as a hidden note", () => {
  const bytes = reviewXlsx(model)
  const sheet = read(bytes, { type: "array" }).Sheets["Tabular review"]
  expect(utils.sheet_to_json<string[]>(sheet, { header: 1, defval: "" })).toEqual([
    ["Document", "What is the governing law?", "Is there a tenant break clause?"],
    ["Unit 4 lease.docx", "England and Wales", "Not addressed"],
    ["Supplier letter.docx", "", "Not addressed"],
  ])
  expect(sheet.B2.c.map((note: { a: string; t: string }) => [note.a, note.t])).toEqual([["doc.haus", `Source: Unit 4 lease.docx § 9.2\n${comment}`]])
  expect(sheet.C2.c.map((note: { t: string }) => note.t)).toEqual([describeRecheck(recheck)])
  expect(sheet.C3.c).toBeUndefined()
  expect(sheet.B3?.c).toBeUndefined()
  expect(unzip(bytes)("xl/drawings/vmlDrawing1.vml")).toContain("visibility:hidden")
})
