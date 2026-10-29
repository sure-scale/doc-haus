import { describe, expect, test } from "bun:test"
import {
  dateInputPlaceholder,
  formatDateInput,
  maskDateInput,
  parseDateInput,
} from "../src/components/chronology/date-input"

describe("chronology numeric dates", () => {
  test("explicit preferences determine day order and digit separators", () => {
    expect(maskDateInput("03042026", "day-month-year")).toBe("03/04/2026")
    expect(parseDateInput("03/04/2026", "day-month-year")).toBe("2026-04-03")
    expect(parseDateInput("03/04/2026", "month-day-year")).toBe("2026-03-04")
    expect(maskDateInput("20260403", "iso")).toBe("2026-04-03")
    expect(parseDateInput("2026-04-03", "iso")).toBe("2026-04-03")
  })
  test("incomplete and invalid inputs cannot replace saved dates; empty is an explicit clear", () => {
    for (const text of ["0", "03/0", "03/04/202", "31/04/2026", "00/04/2026", "03/00/2026", "03/13/2026", "03/04/0000"])
      expect(parseDateInput(text, "day-month-year")).toBeUndefined()
    expect(parseDateInput("", "day-month-year")).toBe("")
    expect(parseDateInput("29/02/2024", "day-month-year")).toBe("2024-02-29")
    expect(parseDateInput("29/02/2026", "day-month-year")).toBeUndefined()
    expect(parseDateInput("1900-02-29", "iso")).toBeUndefined()
    expect(parseDateInput("2000-02-29", "iso")).toBe("2000-02-29")
    expect(parseDateInput("2026-11-31", "iso")).toBeUndefined()
  })
  test("month/year precision never invents a day", () => {
    expect(maskDateInput("052026", "day-month-year", "month")).toBe("05/2026")
    expect(parseDateInput("05/2026", "day-month-year", "month")).toBe("2026-05")
    expect(parseDateInput("2026-05", "iso", "month")).toBe("2026-05")
    expect(parseDateInput("2026", "month-day-year", "year")).toBe("2026")
    expect(parseDateInput("13/2026", "day-month-year", "month")).toBeUndefined()
    expect(parseDateInput("2026", "iso")).toBeUndefined()
    expect(formatDateInput("2026-05", "day-month-year", "month")).toBe("05/2026")
    expect(formatDateInput("2026", "iso", "year")).toBe("2026")
  })
  test("saved dates and placeholders follow the chosen numeric format", () => {
    expect(formatDateInput("2026-12-03", "day-month-year")).toBe("03/12/2026")
    expect(formatDateInput("2026-12-03", "month-day-year")).toBe("12/03/2026")
    expect(formatDateInput("2026-12-03", "iso")).toBe("2026-12-03")
    expect(dateInputPlaceholder("month-day-year")).toBe("mm/dd/yyyy")
    expect(dateInputPlaceholder("day-month-year")).toBe("dd/mm/yyyy")
    expect(dateInputPlaceholder("iso", "month")).toBe("yyyy-mm")
  })
})
