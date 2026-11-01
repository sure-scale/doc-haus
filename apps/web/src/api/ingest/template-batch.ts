import { INGEST_URL } from "../../config"
import { ingestError } from "./errors"
import { rowFill, UNSAFE, type BatchField } from "../../../../../dochaus/lib/template-batch"

// Batch documents from one library template and a spreadsheet: one document
// per row, filled deterministically by ingest (no model calls). The types
// mirror the ingest routes in services/ingest/src/template-batch-routes.ts and
// the shared rules in dochaus/lib/template-batch.ts.

// key is the exact placeholder text ("[Client Name]"); label is what the
// lawyer reads ("Client Name", "Include: Non-solicit").
export type { BatchField }
// Field key → spreadsheet column (null: left for the lawyer to fill).
export type BatchMapping = Record<string, string | null>
// row is the spreadsheet's own row number: header = 1, first data row = 2.
export type BatchRow = { row: number; values: string[] }
// Values the lawyer typed in the preview for one row, field key → text.
export type BatchEntered = Record<string, string>

export type BatchSheet = {
  template: { name: string; sha256: string }
  fields: BatchField[]
  unnamedBlanks: number
  columns: string[]
  rows: BatchRow[]
  dropped: { blank: number; header: number }
  mapping: BatchMapping
  remembered: boolean
  identifyingKey: string | null
  source: string
}

export type BatchDocument = { name: string; row: number; missing: string[] }
export type BatchResult = { id: string; folder: string; documents: BatchDocument[]; skipped: number[]; failed: { row: number; error: string }[] }

// The template file changed between reading the spreadsheet and creating.
export class TemplateChangedError extends Error {
  constructor() {
    super("This template changed after you chose the spreadsheet. Close this and start again to use the new version.")
    this.name = "TemplateChangedError"
  }
}

export async function readTemplateBatch(name: string, file: File): Promise<BatchSheet> {
  const form = new FormData()
  form.append("name", name)
  form.append("file", file)
  const res = await fetch(`${INGEST_URL}/templates/batch/read`, { method: "POST", body: form })
  if (!res.ok) throw new Error(await ingestError(res, `Could not read ${file.name}`))
  return res.json()
}

// One row filled into the template, for the preview. `missing` lists the field
// labels the row leaves for the lawyer to fill.
export async function previewTemplateBatch(input: { name: string; columns: string[]; values: string[]; mapping: BatchMapping; entered?: BatchEntered }): Promise<{ bytes: Uint8Array; missing: string[] }> {
  const res = await fetch(`${INGEST_URL}/templates/batch/preview`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(input),
  })
  if (!res.ok) throw new Error(await ingestError(res, "Could not preview this row"))
  const header = res.headers.get("x-dochaus-batch-missing")
  let missing: string[] = []
  try {
    const parsed: unknown = header ? JSON.parse(decodeURIComponent(header)) : []
    if (Array.isArray(parsed)) missing = parsed.filter((item): item is string => typeof item === "string")
  } catch { /* a malformed header leaves the list empty; the preview still shows its highlights */ }
  return { bytes: new Uint8Array(await res.arrayBuffer()), missing }
}

// A CSV whose header row names the template's fields, downloaded by the browser.
export function templateBatchSampleUrl(name: string) {
  return `${INGEST_URL}/templates/batch/sample?name=${encodeURIComponent(name)}`
}

export async function createTemplateBatch(matterId: string, input: {
  name: string
  sha256: string
  source: string
  columns: string[]
  rows: (BatchRow & { entered?: BatchEntered })[]
  skipped: number[]
  mapping: BatchMapping
}): Promise<BatchResult> {
  const res = await fetch(`${INGEST_URL}/matters/${encodeURIComponent(matterId)}/template-batches`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(input),
  })
  if (res.status === 409) {
    const body = await res.clone().json().catch(() => undefined) as { code?: unknown } | undefined
    if (body?.code === "template_changed") throw new TemplateChangedError()
  }
  if (!res.ok) throw new Error(await ingestError(res, "Could not create the documents"))
  return res.json()
}

// One row as the created document will have it, by the server's own rules
// (dochaus/lib/template-batch.ts), so the preview counts every row exactly as
// Create will. `missing` lists the labels still left for the lawyer once their
// typed values are in; `gaps` the fields the spreadsheet itself leaves, the
// ones the preview offers to type.
export function rowStatus(fields: BatchField[], columns: string[], values: string[], mapping: BatchMapping, entered: BatchEntered = {}) {
  const gaps = new Set(rowFill(fields, columns, values, mapping).missingKeys)
  return {
    missing: rowFill(fields, columns, values, mapping, entered).missing,
    gaps: fields.filter((field) => gaps.has(field.key)),
  }
}

// Why a typed value cannot go into the document, or nothing when it can.
export function enteredProblem(value: string) {
  // Line breaks and tabs read as spaces, as in a cell.
  return UNSAFE.test(value.replace(/[\r\n\t]+/g, " ")) ? "Square brackets can't be used here." : undefined
}
