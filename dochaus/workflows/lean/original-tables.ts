import { readFileSync, realpathSync } from "node:fs"
import path from "node:path"
import { carryOriginalTables, matchTables, pipeTable, plain, readDocxTables, type DocxTable, type TableCarry } from "../../lib/docx-tables"
import { withParagraphPointers } from "../../lib/drafting-amendment"
import { visibleInlineText } from "../../lib/markdown-visible-text"
import { appendLeanEvent } from "./events"
import { plannedDocument, type DocID } from "./plan"
import { excerpt, matterSourceIDs, openItem } from "./quotes"
import type { OpenItem } from "./state"
import { templateSource } from "./template-contract"

// The tables of the original a revised document is drafted from. The
// canonical source text the drafter reads gives each table cell as a paragraph
// of its own and an empty cell as blank lines, so the draft step shows the
// drafter each table in its place as a Markdown table (originalTablesBlock),
// asks once for a table where a part keeps the passage before one but no
// table after it (partTableIssues), and, when the draft is accepted, carries
// the original's own table over each drafted counterpart (carryRevisionTables,
// lib/docx-tables.ts). The apply step records an original table the final
// body no longer has, or the rows it has lost, as an omission (tableItems).

export type OriginalTables = { name: string; tables: DocxTable[] }

/** The original of a revised document: the indexed matter .docx its plan
 * entry is drafted from (templateSource), read from inside the matter. */
export function revisionOriginal(directory: string, document: DocID) {
  const entry = plannedDocument(directory, document)
  const source = entry?.mode === "revise" ? templateSource(entry) : undefined
  if (!source || path.basename(source) !== source || !/\.docx$/i.test(source) || !matterSourceIDs(directory).includes(source)) return
  try {
    const root = realpathSync(directory)
    const file = realpathSync(path.join(root, source))
    const relative = path.relative(root, file)
    if (relative.startsWith("..") || path.isAbsolute(relative)) return
    return { name: source, bytes: new Uint8Array(readFileSync(file)) }
  } catch {
    return
  }
}

/** The original's tables, when the document is a revision of a .docx that has any. */
export function originalTables(directory: string, document: DocID): OriginalTables | undefined {
  const original = revisionOriginal(directory, document)
  if (!original) return
  try {
    const tables = readDocxTables(original.bytes)
    return tables.length ? { name: original.name, tables } : undefined
  } catch {
    return
  }
}

/** The draft prompt's view of the original's tables. */
export function originalTablesBlock(original: OriginalTables | undefined) {
  if (!original) return ""
  return `<original_tables>\n${[
    `The original "${original.name}" has ${original.tables.length === 1 ? "a table" : `${original.tables.length} tables`}. Its text as read-document and search-document give it shows each table cell as a paragraph of its own and an empty cell as blank lines. ` +
    "Here each table is in its place, after the passage it follows:",
    ...original.tables.map((table, index) => `${tableName(table, index)}:\n${pipeTable(table.rows)}`),
    "Write each table the revision keeps as a Markdown table directly after the passage it follows, with every row and column as shown: leave an empty cell empty, as it is left for completion, " +
    "and change a cell only where the instructions change it. A kept table is given the original's own formatting, and a changed cell is changed alone. " +
    "Leave a table out only where the instructions remove it, together with the passage that introduces it.",
  ].join("\n\n")}\n</original_tables>`
}

/** Retry feedback for a part that keeps the passage an original table follows
 * with no table after it before the next heading or the next such passage. */
export function partTableIssues(part: string, tables: DocxTable[] | undefined) {
  if (!tables?.length) return []
  const lines = part.split("\n")
  const passages = tables.map(table => passage(table.follows))
  const kept = new Set<number>()
  lines.forEach((line, at) => {
    const read = passage(line)
    const index = read ? passages.indexOf(read) : -1
    if (index < 0) return
    const next = lines.findIndex((other, after) => after > at && (/^\s*#{1,6}\s/.test(other) || Boolean(passage(other)) && passages.includes(passage(other))))
    const region = lines.slice(at + 1, next < 0 ? undefined : next)
    if (!region.some(row => /^\s*\|.*\|\s*$/.test(row))) kept.add(index)
  })
  return [...kept].map(index => `This part keeps the passage "${excerpt(tables[index]!.follows)}", which table ${index + 1} of the original follows, but has no table after it. ` +
    "Unless the instructions remove or replace that table, write it directly after the passage as a Markdown table, with every row and column and empty cells left empty, changing only the cells the instructions change:\n" +
    `${pipeTable(tables[index]!.rows)}\nIf the instructions do remove or replace it, reply with the part as it was.`)
}

/** The rendered draft with the original's tables carried over their drafted
 * counterparts, and its paragraph pointers completed; what became of each
 * table is recorded in the run's event log. */
export function carryRevisionTables(directory: string, document: DocID, rendered: Uint8Array) {
  const original = revisionOriginal(directory, document)
  if (!original) return rendered
  try {
    const { bytes, carries } = carryOriginalTables(rendered, original.bytes)
    if (carries.length) appendLeanEvent(directory, { stage: "draft", document, event: "tables", step: "tables",
      outcome: (["carried", "edited", "drafted", "missing"] as const).map(kind => `${kind} ${carries.filter(row => row.outcome === kind).length}`).join(", "),
      issues: carries.map(carryText) })
    return bytes === rendered ? rendered : withParagraphPointers(bytes)
  } catch (error) {
    appendLeanEvent(directory, { stage: "draft", document, event: "tables", step: "tables", outcome: "failed", reason: error instanceof Error ? error.message : String(error) })
    return rendered
  }
}

/** An omission for each original table the final body lacks, and for each
 * one that has lost rows, by the text of their first cells. */
export function tableItems(directory: string, document: DocID, final: Uint8Array): OpenItem[] {
  const original = originalTables(directory, document)
  if (!original) return []
  const tables = readDocxTables(final)
  const pairs = matchTables(original.tables, tables)
  return original.tables.flatMap((table, index) => {
    const where = `${tableName(table, index)} of the original "${original.name}"`
    const ref = { source_id: original.name, table: index + 1 }
    const at = pairs[index]
    if (at === undefined)
      return [openItem("apply", "omission", `${where} (rows ${labels(table.rows).map(label => `"${label}"`).join(", ") || "without text"}) is not in the final text; check whether the revision should keep it.`, ref)]
    const present = new Set(tables[at]!.rows.map(row => plain(row[0] ?? "")))
    const lost = labels(table.rows).filter(label => !present.has(plain(label)))
    return lost.length
      ? [openItem("apply", "omission", `${where} has lost ${lost.length === 1 ? "its row" : "its rows"} ${lost.map(label => `"${label}"`).join(", ")} in the final text; check whether the revision should keep ${lost.length === 1 ? "it" : "them"}.`, ref)]
      : []
  })
}

function carryText(carry: TableCarry) {
  const name = `Table ${carry.table}${carry.follows ? ` (after "${excerpt(carry.follows)}")` : ""}`
  if (carry.outcome === "carried") return `${name}: kept as in the original.`
  if (carry.outcome === "missing") return `${name}: not in the draft.`
  if (carry.outcome === "drafted") return `${name}: kept as drafted, because ${carry.reason}.`
  return `${name}: kept from the original with ${[
    carry.cells.length ? `${carry.cells.join("; ")} changed` : "",
    carry.added ? `${carry.added} ${carry.added === 1 ? "row" : "rows"} added` : "",
    carry.removed.length ? `${carry.removed.map(label => `"${label}"`).join(", ")} removed` : "",
  ].filter(Boolean).join(", ")}.`
}

function tableName(table: DocxTable, index: number) {
  return `Table ${index + 1}${table.follows ? `, after "${excerpt(table.follows)}"` : ""}`
}

/** The text of each row's first cell, where it has any. */
function labels(rows: string[][]) {
  return rows.map(row => (row[0] ?? "").split("\n")[0]!.trim()).filter(Boolean)
}

/** A line's words for finding a table's passage: Markdown heading marks and
 * emphasis aside, in any case, closing punctuation aside. */
function passage(line: string) {
  return plain(visibleInlineText(line.trim().replace(/^#{1,6}\s+/, "")).replace(/[*_`]/g, "")).replace(/[\s.:;,]+$/, "")
}
