import { utils, write } from "xlsx"
import { cellKey, type Grid, type GridCellData } from "../api/ingest/grid"
import { describeRecheck } from "../api/review-recheck"
import { formatDate } from "../lib/format"

// The review grid as exported: one model of what the grid shows, read by the
// Excel, CSV and Word serializers, so the three files carry the same content.
// Rows follow the matter's documents and cells follow the question columns.

export type ExportCell = {
  answer: string
  /** "<document> § <section>" of the verified citation. */
  source?: string
  /** The #248 re-check record as the lawyer reads it. */
  recheck?: string
  reviewed: boolean
  comments: string[]
}
export type ReviewExport = { title: string; questions: string[]; rows: Array<{ document: string; cells: ExportCell[] }> }

export function reviewExport(title: string, grid: Grid, documents: string[]): ReviewExport {
  return {
    title,
    questions: grid.columns.map((column) => column.question),
    rows: documents.map((document) => ({ document, cells: grid.columns.map((column) => exportCell(grid.cells[cellKey(document, column.id)])) })),
  }
}

function exportCell(cell: GridCellData | undefined): ExportCell {
  return {
    answer: cell?.answer ?? "",
    source: cell?.citation && `${cell.citation.documentName} § ${cell.citation.section}`,
    recheck: cell?.recheck && describeRecheck(cell.recheck),
    reviewed: cell?.status === "reviewed",
    comments: (cell?.comments ?? []).map((comment) => `Comment (${formatDate(comment.at)}): ${comment.text}`),
  }
}

/** The lines that travel with an answer: source, re-check, then comments. */
function notes(cell: ExportCell) {
  return [cell.source && `Source: ${cell.source}`, cell.recheck, ...cell.comments].filter((line): line is string => Boolean(line))
}

/** Excel: the grid's shape, with each answer's notes as a native hidden note on its cell. */
export function reviewXlsx(model: ReviewExport): ArrayBuffer {
  const sheet = utils.aoa_to_sheet([["Document", ...model.questions], ...model.rows.map((row) => [row.document, ...row.cells.map((cell) => cell.answer)])])
  model.rows.forEach((row, r) =>
    row.cells.forEach((cell, c) => {
      const note = notes(cell).join("\n")
      if (note) sheet[utils.encode_cell({ r: r + 1, c: c + 1 })].c = Object.assign([{ a: "doc.haus", t: note }], { hidden: true })
    }),
  )
  const book = utils.book_new()
  utils.book_append_sheet(book, sheet, "Tabular review")
  return write(book, { type: "array", bookType: "xlsx" })
}

/** CSV: one record per document and question, so every note keeps its own column.
 * RFC 4180 with CRLF records; the BOM makes Excel read UTF-8 (every source has a §). */
export function reviewCsv(model: ReviewExport) {
  const records = [
    ["Document", "Question", "Answer", "Source", "Reviewed", "Re-check", "Comments"],
    ...model.rows.flatMap((row) =>
      row.cells.map((cell, c) => [row.document, model.questions[c], cell.answer, cell.source ?? "", cell.reviewed ? "Yes" : "No", cell.recheck ?? "", cell.comments.join("\n")]),
    ),
  ]
  return `﻿${records.map((record) => record.map(csvField).join(",")).join("\r\n")}\r\n`
}

// Answers and names come from documents, so a value a spreadsheet would run as a
// formula is kept as text with a leading apostrophe (OWASP CSV injection).
function csvField(value: string) {
  const text = /^[=+\-@\t\r]/.test(value) ? `'${value}` : value
  return /[",\r\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text
}

const PAGE = { width: 11906, height: 16838, margin: 1440 } // A4 in twips, 1" margins

/** Word: a heading and one table in the grid's shape; landscape from three questions. */
export async function reviewDocx(model: ReviewExport): Promise<ArrayBuffer> {
  const { Document, HeadingLevel, Packer, PageOrientation, Paragraph, ShadingType, Table, TableCell, TableLayoutType, TableRow, TextRun, WidthType } = await import("docx")
  const landscape = model.questions.length >= 3
  const width = (landscape ? PAGE.height : PAGE.width) - 2 * PAGE.margin
  const first = Math.round(width * 0.2)
  const rest = model.questions.length ? Math.floor((width - first) / model.questions.length) : 0
  const columnWidths = [first, ...model.questions.map(() => rest)]
  const lines = (text: string, run: { bold?: boolean; size?: number; color?: string } = {}) =>
    text.split(/\r?\n|\r/).map((line) => new Paragraph({ children: [new TextRun({ text: line, ...run })] }))
  const cell = (c: number, children: InstanceType<typeof Paragraph>[], header = false) =>
    new TableCell({
      children,
      width: { size: columnWidths[c], type: WidthType.DXA },
      margins: { top: 60, bottom: 60, left: 100, right: 100 },
      shading: header ? { type: ShadingType.CLEAR, color: "auto", fill: "EDEDED" } : undefined,
    })
  const note = { size: 16, color: "595959" }
  const table = new Table({
    width: { size: width, type: WidthType.DXA },
    columnWidths,
    layout: TableLayoutType.FIXED,
    rows: [
      new TableRow({ tableHeader: true, children: ["Document", ...model.questions].map((text, c) => cell(c, lines(text, { bold: true }), true)) }),
      ...model.rows.map((row) =>
        new TableRow({
          children: [
            cell(0, lines(row.document)),
            ...row.cells.map((value, c) =>
              cell(c + 1, [...lines(value.answer), ...(value.reviewed ? lines("Reviewed", note) : []), ...notes(value).flatMap((line) => lines(line, note))]),
            ),
          ],
        }),
      ),
    ],
  })
  const document = new Document({
    styles: { default: { document: { run: { font: "Calibri", size: 20 } } } },
    sections: [
      {
        properties: { page: { size: { width: PAGE.width, height: PAGE.height, orientation: landscape ? PageOrientation.LANDSCAPE : PageOrientation.PORTRAIT }, margin: { top: PAGE.margin, right: PAGE.margin, bottom: PAGE.margin, left: PAGE.margin } } },
        children: [new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(`Tabular review — ${model.title}`)] }), table],
      },
    ],
  })
  return Packer.toArrayBuffer(document)
}
