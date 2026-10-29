import type { DraftingPreferences } from "../../api/ingest/preferences"

export type DateInputFormat = DraftingPreferences["dateFormat"]
export type DateInputPrecision = "day" | "month" | "year"

export function dateInputPlaceholder(format: DateInputFormat, precision: DateInputPrecision = "day") {
  if (precision === "year") return "yyyy"
  if (precision === "month") return format === "iso" ? "yyyy-mm" : "mm/yyyy"
  return format === "iso" ? "yyyy-mm-dd" : format === "day-month-year" ? "dd/mm/yyyy" : "mm/dd/yyyy"
}

/** Insert separators without supplying any missing date part. */
export function maskDateInput(text: string, format: DateInputFormat, precision: DateInputPrecision = "day") {
  const parts =
    precision === "year"
      ? [4]
      : precision === "month"
        ? format === "iso"
          ? [4, 2]
          : [2, 4]
        : format === "iso"
          ? [4, 2, 2]
          : [2, 2, 4]
  const digits = text.replace(/\D/g, "").slice(
    0,
    parts.reduce((sum, part) => sum + part, 0),
  )
  const result: string[] = []
  let offset = 0
  for (const length of parts) {
    if (digits.length <= offset) break
    result.push(digits.slice(offset, offset + length))
    offset += length
  }
  return result.join(format === "iso" ? "-" : "/")
}

/** Empty is an explicit clear; undefined is incomplete or invalid. No Date parsing/timezone coercion. */
export function parseDateInput(
  text: string,
  format: DateInputFormat,
  precision: DateInputPrecision = "day",
): string | undefined {
  if (!text) return ""
  const expected = dateInputPlaceholder(format, precision)
  if (text.length !== expected.length) return undefined
  let year: string, month: string | undefined, day: string | undefined
  if (precision === "year") year = text
  else {
    const parts = text.split(format === "iso" ? "-" : "/")
    if (precision === "month") [year, month] = format === "iso" ? parts : [parts[1]!, parts[0]!]
    else if (format === "iso") [year, month, day] = parts
    else if (format === "day-month-year") [day, month, year] = parts
    else [month, day, year] = parts
  }
  if (!/^\d{4}$/.test(year!) || Number(year) < 1) return undefined
  if (precision === "year") return year!
  if (!month || !/^\d{2}$/.test(month) || Number(month) < 1 || Number(month) > 12) return undefined
  if (precision === "month") return `${year!}-${month}`
  const leap = Number(year) % 4 === 0 && (Number(year) % 100 !== 0 || Number(year) % 400 === 0)
  const days = [31, leap ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
  if (!day || !/^\d{2}$/.test(day) || Number(day) < 1 || Number(day) > days[Number(month) - 1]!) return undefined
  return `${year!}-${month}-${day}`
}

export function formatDateInput(
  value: string | undefined,
  format: DateInputFormat,
  precision: DateInputPrecision = "day",
) {
  if (!value) return ""
  const [year, month, day] = value.split("-")
  if (precision === "year") return year ?? ""
  if (precision === "month")
    return format === "iso" ? [year, month].filter(Boolean).join("-") : month ? `${month}/${year}` : (year ?? "")
  if (format === "iso") return value
  return month && day ? (format === "day-month-year" ? `${day}/${month}/${year}` : `${month}/${day}/${year}`) : value
}
