import type { WordChatCellRun, WordChatTableCapture } from "../../../dochaus/lib/word-chat-amendment-contract"

type Collection<T> = { items: T[]; load(properties: string): void }
type Paragraph = { uniqueLocalId: string; text: string; tableNestingLevel: number }
type Font = { bold: boolean; italic: boolean; underline: string; load(properties: string): void }
type Range = { font: Font }
type Cell = { body: { paragraphs: Collection<Paragraph>; insertText(text: string, location: "End"): Range } }
type Table = {
  rowCount: number
  isUniform: boolean
  values: string[][]
  load(properties: string): void
  getCell(row: number, column: number): Cell
  addRows(location: "End", count: number, values: string[][]): unknown
}
export type NativeRowContext = {
  document: {
    changeTrackingMode: string
    trackFormatting: boolean
    load(properties: string): void
    body: { tables: Collection<Table> }
  }
  sync(): Promise<void>
}
export type PreparedNativeRow = { table: Table; rowCount: number; width: number; tracking: string; trackFormatting: boolean }

/** Read and bind all cells, independently of the cursor. No writes are queued. */
export async function prepareNativeRow(context: NativeRowContext, binding: WordChatTableCapture): Promise<PreparedNativeRow> {
  const fail = (): never => { throw new Error("The reviewed Word table has changed or cannot be identified.") }
  if (!binding || !Number.isSafeInteger(binding.tableIndex) || binding.tableIndex < 0 || !Array.isArray(binding.rows) || !binding.rows.length || binding.rows.length > 1000 || !Array.isArray(binding.rows[0]) || !binding.rows[0].length || binding.rows[0].length > 32) return fail()
  const width = binding.rows[0].length
  if (binding.rows.length * width > 20_000 || binding.rows.some(row => !Array.isArray(row) || row.length !== width || row.some(cell => !cell || typeof cell.id !== "string" || !cell.id || typeof cell.text !== "string"))) return fail()
  const ids = binding.rows.flat().map(cell => cell.id)
  if (new Set(ids).size !== ids.length) return fail()
  const tables = context.document.body.tables
  tables.load("items")
  context.document.load("changeTrackingMode,trackFormatting")
  await context.sync()
  for (const table of tables.items) table.load("rowCount,isUniform,values")
  await context.sync()
  const table = tables.items[binding.tableIndex]
  if (!table || !table.isUniform || table.rowCount !== binding.rows.length || !Array.isArray(table.values) || table.values.length !== table.rowCount || table.values.some(row => !Array.isArray(row) || row.length !== width)) return fail()
  const cells = binding.rows.map((row, ri) => row.map((_, ci) => {
    const paragraphs = table.getCell(ri, ci).body.paragraphs
    paragraphs.load("items/text,items/uniqueLocalId,items/tableNestingLevel")
    return paragraphs
  }))
  await context.sync()
  if (cells.some((row, ri) => row.some((paragraphs, ci) => {
    const p = paragraphs.items[0], expected = binding.rows[ri]![ci]!
    return paragraphs.items.length !== 1 || !p || p.uniqueLocalId !== expected.id || p.text !== expected.text || p.tableNestingLevel !== 1
  }))) return fail()
  if (typeof context.document.changeTrackingMode !== "string" || typeof context.document.trackFormatting !== "boolean") return fail()
  return { table, rowCount: table.rowCount, width, tracking: context.document.changeTrackingMode, trackFormatting: context.document.trackFormatting }
}

/** Queue the native insertion without syncing. The caller must fence first and
 * perform exactly one final sync; any lost sync leaves the result uncertain.
 * A cell with reviewed runs is inserted run by run, and every run sets all
 * three flags because text appended at End inherits its neighbour's font. */
export function queueNativeRow(context: NativeRowContext, prepared: PreparedNativeRow, cells: string[], runs?: WordChatCellRun[][]): { ranges: Range[] } {
  if (!Array.isArray(cells) || cells.length !== prepared.width || cells.some(cell => typeof cell !== "string")) throw new Error("The reviewed row does not match the table width.")
  if (runs !== undefined && (!Array.isArray(runs) || runs.length !== cells.length || runs.some((cell, column) => !Array.isArray(cell) ||
    cell.some(run => !run || typeof run.text !== "string" || !run.text || [run.bold, run.italic, run.underline].some(flag => typeof flag !== "boolean")) ||
    cell.map(run => run.text).join("") !== cells[column]))) throw new Error("The reviewed row formatting does not match its cells.")
  context.document.changeTrackingMode = "TrackAll"
  prepared.table.addRows("End", 1, [cells.map(() => "")])
  const ranges = cells.flatMap((text, column) => {
    const body = prepared.table.getCell(prepared.rowCount, column).body
    const planned = runs?.[column]?.length ? runs[column]! : [{ text, bold: false, italic: false, underline: false }]
    return planned.map(run => {
      const range = body.insertText(run.text, "End")
      range.font.bold = run.bold
      range.font.italic = run.italic
      range.font.underline = run.underline ? "Single" : "None"
      return range
    })
  })
  context.document.changeTrackingMode = prepared.tracking
  context.document.trackFormatting = prepared.trackFormatting
  return { ranges }
}
