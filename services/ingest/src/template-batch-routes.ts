import { Hono, type Context } from "hono"
import { randomUUID } from "node:crypto"
import { existsSync, mkdirSync, readdirSync } from "node:fs"
import path from "node:path"
import { read, utils } from "xlsx"
import { readDurableJson, writeDurableJson } from "../../../dochaus/lib/durable-json"
import {
  autoMatch,
  batchDocumentName,
  batchFolderName,
  identifyingKey,
  rowFill,
  templateBatchFields,
} from "../../../dochaus/lib/template-batch"
import { fillTemplateDocx } from "../../../dochaus/lib/template-fill"
import { attachmentDocumentDisposition } from "./content-disposition"
import { writeDeliverable } from "./deliverable"
import { oneLine, sha256Hex } from "./document-versions"
import { sha256File } from "../../../dochaus/lib/lineage"
import { docxodus } from "./docxodus"
import { getMatter } from "./matter"
import { templatePath, TEMPLATES_DIR } from "./template"

// Many documents from one library template and a spreadsheet (one row, one
// document). The spreadsheet is read here, the documents come from the shared
// deterministic template fill (dochaus/lib/template-batch.ts), and each one is
// written as a generated deliverable. No model is called and no value is
// guessed: what a row does not supply stays a highlighted placeholder.

const MAX_ROWS = 1000
const MAX_COLUMNS = 200
const MAX_ENTERED = 2000
const DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

type Mapping = Record<string, string | null>
// `entered`: values typed in the preview for this row, by field key.
type Row = { row: number; values: string[]; entered?: Record<string, string> }
type BatchDocument = { name: string; row: number; missing: string[] }
// What a document was written with: its bytes' sha256 and the placeholder
// text of each missing label (missingKeys[i] for missing[i]).
type RecordedDocument = BatchDocument & { sha256: string; missingKeys: string[] }
type BatchRecord = {
  id: string
  folder: string
  createdAt: number
  template: { name: string; sha256: string }
  source: string
  mapping: Mapping
  documents: RecordedDocument[]
  skipped: number[]
  failed: { row: number; error: string }[]
}

class BatchRequestError extends Error {
  constructor(message: string, readonly status: 400 | 404 = 400) { super(message) }
}

const mappingsFile = () => path.join(TEMPLATES_DIR, "batch-mappings.json")
const recordsDir = (matterDir: string) => path.join(matterDir, ".dochaus", "template-batches")
const stem = (name: string) => name.replace(/\.docx$/i, "")

export function templateBatchRoutes() {
  const app = new Hono()
  app.onError((error, c) => {
    if (error instanceof BatchRequestError) return c.json({ error: error.message }, error.status)
    throw error
  })

  app.post("/templates/batch/read", async (c) => {
    const body = await c.req.parseBody()
    const file = body["file"]
    if (!(file instanceof File)) throw new BatchRequestError("Choose a spreadsheet (.csv, .xlsx or .xls).")
    const template = await loadTemplate(body["name"])
    const sheet = readSpreadsheet(file.name, Buffer.from(await file.arrayBuffer()))
    const remembered = readMappings()[template.name] ?? {}
    const mapping = autoMatch(template.fields, sheet.columns, remembered)
    return c.json({
      template: { name: template.name, sha256: template.sha256 },
      fields: template.fields,
      unnamedBlanks: template.unnamedBlanks,
      columns: sheet.columns,
      rows: sheet.rows,
      dropped: sheet.dropped,
      mapping,
      remembered: Object.entries(remembered).some(([key, column]) => mapping[key] === column),
      identifyingKey: identifyingKey(template.fields, mapping),
      source: sourceName(file.name),
    })
  })

  app.post("/templates/batch/preview", async (c) => {
    const body = await c.req.json().catch(() => null) as Record<string, unknown> | null
    const template = await loadTemplate(body?.name)
    const columns = strings(body?.columns, "columns")
    const values = strings(body?.values, "values")
    const fill = rowFill(template.fields, columns, values, mappingOf(body?.mapping), enteredOf(body?.entered))
    const { bytes } = await fillTemplateDocx(template.bytes, fill, { markFilled: true })
    return new Response(Buffer.from(bytes), {
      headers: { "Content-Type": DOCX, "X-DocHaus-Batch-Missing": encodeURIComponent(JSON.stringify(fill.missing)) },
    })
  })

  app.get("/templates/batch/sample", async (c) => {
    const template = await loadTemplate(c.req.query("name"))
    const labels = [...template.fields.filter((f) => f.kind === "fill"), ...template.fields.filter((f) => f.kind === "optional")]
      .map((field) => field.label)
    // The byte order mark makes Excel read the file as UTF-8.
    return new Response(`\uFEFF${labels.map(csvCell).join(",")}\r\n`, {
      headers: {
        "Content-Type": "text/csv; charset=utf-8",
        "Content-Disposition": attachmentDocumentDisposition(`${stem(template.name)} – spreadsheet.csv`),
      },
    })
  })

  app.post("/matters/:id/template-batches", async (c) => {
    const matter = getMatter(c.req.param("id"))
    const body = await c.req.json().catch(() => null) as Record<string, unknown> | null
    const template = await loadTemplate(body?.name)
    if (body?.sha256 !== template.sha256)
      return c.json({ error: `${template.name} changed since the spreadsheet was read. Read the spreadsheet again.`, code: "template_changed" }, 409)
    const columns = strings(body?.columns, "columns")
    const mapping = mappingOf(body?.mapping)
    const rows = rowsOf(body?.rows)
    const skipped = Array.isArray(body?.skipped) ? body.skipped.filter((n): n is number => Number.isSafeInteger(n)) : []
    const source = sourceName(typeof body?.source === "string" ? body.source : "")
    const skip = new Set(skipped)
    const idKey = identifyingKey(template.fields, mapping)
    const width = Math.max(3, String(Math.max(0, ...rows.map((row) => row.row))).length)
    const actor = requestActor(c)
    const documents: RecordedDocument[] = []
    const failed: { row: number; error: string }[] = []
    const taken = new Set<string>()
    // One at a time: each document is its own registered write.
    for (const { row, values, entered } of rows) {
      if (skip.has(row)) continue
      try {
        const fill = rowFill(template.fields, columns, values, mapping, entered)
        const { bytes } = await fillTemplateDocx(template.bytes, fill)
        const name = freeName(matter.dir, batchDocumentName(template.name, idKey ? fill.fills[idKey] : undefined, row, width), taken)
        taken.add(name)
        const written = await writeDeliverable(matter.dir, name, Buffer.from(bytes), {
          precondition: { kind: "create" },
          version: { origin: "generated", actor, label: `From ${stem(template.name)} · spreadsheet row ${row} of ${source}` },
        })
        documents.push({ name, row, missing: fill.missing, missingKeys: fill.missingKeys, sha256: sha256File(written.docPath)?.sha256 ?? "" })
      } catch (error) {
        failed.push({ row, error: error instanceof Error ? error.message : String(error) })
      }
    }
    const record: BatchRecord = {
      id: randomUUID(),
      folder: batchFolderName(template.name, new Date()),
      createdAt: Date.now(),
      template: { name: template.name, sha256: template.sha256 },
      source,
      mapping,
      documents,
      skipped: [...skip].filter((row) => rows.some((r) => r.row === row)),
      failed,
    }
    mkdirSync(recordsDir(matter.dir), { recursive: true })
    writeDurableJson(path.join(recordsDir(matter.dir), `${record.id}.json`), record)
    const used = Object.fromEntries(Object.entries(mapping).filter((entry): entry is [string, string] => typeof entry[1] === "string"))
    mkdirSync(TEMPLATES_DIR, { recursive: true })
    writeDurableJson(mappingsFile(), { ...readMappings(), [template.name]: used })
    return c.json({ id: record.id, folder: record.folder, documents: documents.map(({ name, row, missing }) => ({ name, row, missing })), skipped: record.skipped, failed })
  })

  return app
}

/** The matter's documents, each one a batch created carrying where it came
 * from and what it still needs. A document unchanged since the batch wrote it
 * (same sha256) reports the recorded missing labels; a changed one is read
 * again and keeps only the labels whose placeholder text is still in it (an
 * optional component's choice leaves no marker, so it drops once changed).
 * A document whose file is gone carries nothing. */
export async function withTemplateBatches<T extends { name: string; role?: unknown; doc_path?: unknown }>(matterDir: string, documents: T[]) {
  const dir = recordsDir(matterDir)
  if (!existsSync(dir)) return documents
  const byName = new Map<string, { record: BatchRecord; document: RecordedDocument }>()
  const records = readdirSync(dir)
    .filter((file) => file.endsWith(".json"))
    .flatMap((file) => {
      const record = readDurableJson(path.join(dir, file)) as BatchRecord | undefined
      return record && Array.isArray(record.documents) ? [record] : []
    })
    .sort((a, b) => a.createdAt - b.createdAt)
  for (const record of records) for (const document of record.documents) byName.set(document.name, { record, document })
  return Promise.all(documents.map(async (document) => {
    const found = document.role === "deliverable" ? byName.get(document.name) : undefined
    if (!found) return document
    const file = typeof document.doc_path === "string" ? document.doc_path : path.join(matterDir, document.name)
    const current = sha256File(file)
    if (!current) return document
    const { record, document: written } = found
    const missing = current.sha256 === written.sha256 || !written.missing.length ? written.missing : await stillMissing(file, written)
    return { ...document, batch: { id: record.id, folder: record.folder, row: written.row, source: record.source, missing } }
  }))
}

async function stillMissing(file: string, written: RecordedDocument) {
  const session = (await docxodus()).openDocxSession(await Bun.file(file).bytes(), {})
  try {
    const present = new Set(session.findPlaceholders().map((p) => p.match.text))
    return written.missing.filter((_, index) => present.has(written.missingKeys?.[index] ?? ""))
  } finally {
    session.close()
  }
}

async function loadTemplate(value: unknown) {
  const requested = typeof value === "string" ? value : ""
  const file = templatePath(requested)
  if (!requested || !file.toLowerCase().endsWith(".docx") || !existsSync(file)) throw new BatchRequestError("That template is not in the library.", 404)
  const bytes = await Bun.file(file).bytes()
  const session = (await docxodus()).openDocxSession(bytes, {})
  try {
    const placeholders = session.findPlaceholders().map((p) => ({ text: p.match.text, kind: p.kind, hint: p.hint }))
    return { name: path.basename(file), bytes, sha256: sha256Hex(bytes), ...templateBatchFields(placeholders) }
  } finally {
    session.close()
  }
}

// The first sheet as text, each cell as the spreadsheet displays it. The first
// non-blank row is the header; the row numbers are the spreadsheet's own.
function readSpreadsheet(fileName: string, buffer: Buffer) {
  const ext = path.extname(fileName).toLowerCase()
  if (![".csv", ".xlsx", ".xls"].includes(ext)) throw new BatchRequestError("Choose a spreadsheet (.csv, .xlsx or .xls).")
  const workbook = (() => {
    try {
      if (ext !== ".csv") return read(buffer, { type: "buffer" })
      // A CSV saved by Excel is often not UTF-8; it reads as Windows-1252.
      const utf8 = new TextDecoder().decode(buffer)
      const text = utf8.includes("\uFFFD") ? new TextDecoder("windows-1252").decode(buffer) : utf8
      return read(text.replace(/^\uFEFF/, ""), { type: "string", raw: true })
    } catch {
      throw new BatchRequestError(`${path.basename(fileName)} could not be read as a spreadsheet.`)
    }
  })()
  const sheet = workbook.SheetNames[0] ? workbook.Sheets[workbook.SheetNames[0]] : undefined
  if (!sheet?.["!ref"]) throw new BatchRequestError("The spreadsheet is empty.")
  const first = utils.decode_range(sheet["!ref"]).s.r
  const grid = utils.sheet_to_json<unknown[]>(sheet, { header: 1, raw: false, defval: "", blankrows: true })
    .map((cells) => cells.map((cell) => String(cell ?? "").trim()))
  const blank = (cells: string[]) => cells.every((cell) => !cell)
  const headerAt = grid.findIndex((cells) => !blank(cells))
  if (headerAt < 0) throw new BatchRequestError("The spreadsheet is empty.")
  const lastAt = grid.findLastIndex((cells) => !blank(cells))
  const body = grid.slice(headerAt, lastAt + 1)
  const width = Math.max(0, ...body.map((cells) => cells.findLastIndex(Boolean) + 1))
  if (width > MAX_COLUMNS) throw new BatchRequestError(`The spreadsheet has ${width} columns; use at most ${MAX_COLUMNS}.`)
  const header = Array.from({ length: width }, (_, index) => grid[headerAt]![index] ?? "")
  const columns: string[] = []
  for (const [index, cell] of header.entries()) {
    const name = cell || `Column ${utils.encode_col(index)}`
    let unique = name
    for (let n = 2; columns.includes(unique); n++) unique = `${name} (${n})`
    columns.push(unique)
  }
  const dropped = { blank: 0, header: 0 }
  const rows: Row[] = []
  for (let at = headerAt + 1; at <= lastAt; at++) {
    const values = Array.from({ length: width }, (_, index) => grid[at]![index] ?? "")
    if (blank(values)) dropped.blank++
    else if (values.every((value, index) => value.toLowerCase() === header[index]!.toLowerCase())) dropped.header++
    else rows.push({ row: first + at + 1, values })
  }
  if (!rows.length) throw new BatchRequestError("The spreadsheet has no rows below its header row.")
  if (rows.length > MAX_ROWS) throw new BatchRequestError(`The spreadsheet has ${rows.length} rows; use at most ${MAX_ROWS} at a time.`)
  return { columns, rows, dropped }
}

function readMappings() {
  const value = readDurableJson(mappingsFile())
  return value && typeof value === "object" && !Array.isArray(value) ? value as Record<string, Record<string, string>> : {}
}

// A name not yet in the matter folder nor in this batch: "<name> (2).docx" and on.
function freeName(matterDir: string, name: string, taken: Set<string>) {
  const base = name.replace(/\.docx$/i, "")
  let candidate = name
  for (let n = 2; taken.has(candidate) || existsSync(path.join(matterDir, candidate)); n++) candidate = `${base} (${n}).docx`
  return candidate
}

function sourceName(name: string) {
  return oneLine(path.basename(name)).slice(0, 200) || "spreadsheet"
}

function requestActor(c: Context) {
  return oneLine(c.req.header("x-dochaus-actor")) || undefined
}

function strings(value: unknown, field: string) {
  if (!Array.isArray(value) || !value.every((item) => typeof item === "string")) throw new BatchRequestError(`"${field}" must be a list of text values.`)
  return value as string[]
}

function mappingOf(value: unknown): Mapping {
  if (!value || typeof value !== "object" || Array.isArray(value)) throw new BatchRequestError(`"mapping" must map each field to a column or null.`)
  return Object.fromEntries(Object.entries(value).map(([key, column]) => [key, typeof column === "string" ? column : null]))
}

function rowsOf(value: unknown): Row[] {
  if (!Array.isArray(value) || value.length > MAX_ROWS) throw new BatchRequestError(`"rows" must list at most ${MAX_ROWS} rows.`)
  return value.map((row) => {
    const { row: number, values, entered } = (row ?? {}) as { row?: unknown; values?: unknown; entered?: unknown }
    if (!Number.isSafeInteger(number) || (number as number) < 1) throw new BatchRequestError("Each row needs its spreadsheet row number.")
    return { row: number as number, values: strings(values, "values"), entered: enteredOf(entered) }
  })
}

// Values typed in the preview, field key -> text. rowFill applies its own
// checks; only the shape and size are checked here.
function enteredOf(value: unknown): Record<string, string> {
  if (value === undefined || value === null) return {}
  if (typeof value !== "object" || Array.isArray(value) || !Object.values(value).every((text) => typeof text === "string" && text.length <= MAX_ENTERED))
    throw new BatchRequestError(`"entered" must map fields to text of at most ${MAX_ENTERED} characters.`)
  return value as Record<string, string>
}

function csvCell(value: string) {
  return /[",\r\n]|^\s|\s$/.test(value) ? `"${value.replaceAll('"', '""')}"` : value
}
