import { safeFileName } from "./safe-file-name"

// Many documents from one library template and a spreadsheet: one row, one
// document. Column-to-field matching, per-row values and file naming are pure
// here (and shared with the web preview, so keep this file free of Node and
// docx imports); the document itself comes from the same deterministic fill
// draft-document uses (fillTemplateDocx in template-fill.ts). No model is involved and no value is
// ever guessed: anything the row does not supply stays a visible placeholder.

export type BatchField = { key: string; label: string; kind: "fill" | "optional"; hint?: string }

export const OPTIONAL_FIELD = /^\[\s*optional:\s*([^\]]*?)\s*\]$/i

/** The fields a spreadsheet can fill, one per distinct placeholder text. A
 * placeholder with no name ("[___]", "[●]") is not offered; it is counted and
 * stays highlighted in every document. */
export function templateBatchFields(placeholders: { text: string; kind: string; hint?: string }[]) {
  const fields: BatchField[] = []
  const seen = new Set<string>()
  let unnamedBlanks = 0
  for (const placeholder of placeholders) {
    const key = placeholder.text
    if (seen.has(key)) continue
    seen.add(key)
    const optional = OPTIONAL_FIELD.exec(key)
    if (optional) {
      if (optional[1]) fields.push({ key, label: `Include: ${optional[1]}`, kind: "optional" })
      continue
    }
    const label = placeholder.kind === "blank_fill" ? "" : fieldLabel(key)
    if (!/[\p{L}\p{N}]/u.test(label)) {
      unnamedBlanks++
      continue
    }
    fields.push({ key, label, kind: "fill", ...(placeholder.hint ? { hint: placeholder.hint } : {}) })
  }
  return { fields, unnamedBlanks }
}

// "[Client Name]" -> "Client Name"; "[insert notice date]" -> "Notice date".
function fieldLabel(key: string) {
  const inner = key.replace(/^\[\s*|\s*\]$/g, "").replace(/\s+/g, " ")
  const instruction = /^insert\s+(.+)$/i.exec(inner)
  if (!instruction) return inner
  return instruction[1]!.charAt(0).toUpperCase() + instruction[1]!.slice(1)
}

/** A header or field name reduced for matching: lower case, no brackets, no
 * leading "insert " or "include:", punctuation as spaces. */
export function normaliseHeader(text: string) {
  return text
    .toLowerCase()
    .replace(/[[\]]/g, " ")
    .trim()
    .replace(/^insert\s+/, "")
    .replace(/^include\b\s*:?/, "")
    .replace(/[\p{P}\p{S}_-]+/gu, " ")
    .replace(/\s+/g, " ")
    .trim()
}

/** Field key -> column name, or null where nothing matches. A remembered
 * column still present wins, then a column equal to the label, then one equal
 * once both are normalised. Each column fills at most one field. */
export function autoMatch(fields: BatchField[], columns: string[], remembered: Record<string, string> = {}) {
  const mapping: Record<string, string | null> = Object.fromEntries(fields.map((field) => [field.key, null]))
  const used = new Set<string>()
  const assign = (pick: (field: BatchField) => string | undefined) => {
    for (const field of fields) {
      if (mapping[field.key]) continue
      const column = pick(field)
      if (column === undefined || used.has(column)) continue
      mapping[field.key] = column
      used.add(column)
    }
  }
  assign((field) => (columns.includes(remembered[field.key] ?? "") ? remembered[field.key] : undefined))
  assign((field) => columns.find((column) => !used.has(column) && column === field.label))
  assign((field) => {
    const names = new Set([normaliseHeader(field.label), normaliseHeader(field.key)].filter(Boolean))
    return columns.find((column) => !used.has(column) && names.has(normaliseHeader(column)))
  })
  return mapping
}

/** `missingKeys[i]` is the placeholder text of `missing[i]`. */
export type RowFill = { fills: Record<string, string>; include: Record<string, boolean>; missing: string[]; missingKeys: string[] }

const YES = new Set(["yes", "y", "true", "1", "include", "included"])
const NO = new Set(["no", "n", "false", "0", "omit", "exclude"])

/** One row's values for the template. A fill with no usable value is missing
 * and keeps its placeholder; a value holding a bracket or a control character
 * is never written into the document (line breaks and tabs read as spaces).
 * An optional component with no column or a blank cell keeps the template's
 * default (kept); one whose cell is not a clear yes or no is missing, and is
 * kept with its marker left in for the lawyer to resolve. */
export function rowFill(fields: BatchField[], columns: string[], row: string[], mapping: Record<string, string | null>): RowFill {
  const fills: Record<string, string> = {}
  const include: Record<string, boolean> = {}
  const missing: string[] = []
  const missingKeys: string[] = []
  const lack = (field: BatchField) => {
    missing.push(field.label)
    missingKeys.push(field.key)
  }
  for (const field of fields) {
    const column = mapping[field.key]
    const index = column ? columns.indexOf(column) : -1
    const value = index >= 0 ? (row[index] ?? "").replace(/[\r\n\t]+/g, " ").replace(/ {2,}/g, " ").trim() : undefined
    if (field.kind === "optional") {
      if (!value) {
        include[field.key] = true
        continue
      }
      const answer = value.toLowerCase()
      if (YES.has(answer)) include[field.key] = true
      else if (NO.has(answer)) include[field.key] = false
      else {
        include[field.key] = true
        lack(field)
      }
      continue
    }
    if (!value || /[[\]\p{Cc}]/u.test(value)) lack(field)
    else fills[field.key] = value
  }
  return { fills, include, missing, missingKeys }
}

const IDENTIFYING = /name|party|client|company|counterparty|recipient|addressee|tenant|employee/i

/** The field that tells one row's document from another (a party's name),
 * else the first matched fill. */
export function identifyingKey(fields: BatchField[], mapping: Record<string, string | null>) {
  const mapped = fields.filter((field) => field.kind === "fill" && mapping[field.key])
  return (mapped.find((field) => IDENTIFYING.test(field.label)) ?? mapped[0])?.key ?? null
}

const stem = (template: string) => (template.split(/[\\/]/).pop() ?? template).replace(/\.docx$/i, "")

/** "<Template> – <value> – row 008.docx", or without the value when blank. */
export function batchDocumentName(template: string, identifying: string | undefined, sheetRow: number, width: number) {
  const value = (identifying ?? "").replace(/\s+/g, " ").trim().slice(0, 80).trim()
  const row = `row ${String(sheetRow).padStart(width, "0")}`
  return `${safeFileName([stem(template), value, row].filter(Boolean).join(" – ")) || row}.docx`
}

/** "<Template> – YYYY-MM-DD" in local time. */
export function batchFolderName(template: string, date: Date) {
  const day = [date.getFullYear(), date.getMonth() + 1, date.getDate()].map((part) => String(part).padStart(2, "0")).join("-")
  return `${stem(template)} – ${day}`
}
