import { describe, expect, test } from "bun:test"
import type { ChronologyDate, ChronologyEvent } from "../lib/chronology"
import { chronologyExportDate, chronologyExportEventNotes } from "../lib/chronology-export-format"

const entry = (date: ChronologyDate): ChronologyEvent => ({
  id: "inspection",
  title: "Inspection reported",
  detail: "The inspector reported an inspection.",
  date,
  accounts: [],
  origin: "manual",
  reviewed: false,
  opened: false,
  hidden: false,
  key: false,
  disputed: false,
  note: "",
  lawyerEdited: false,
  createdAt: 1,
  updatedAt: 1,
  previousIds: [],
})

describe("chronology export date presentation", () => {
  test("compacts equivalent calendar dates while notes retain the original and normalized values", () => {
    const date: ChronologyDate = { original: "May 2026", start: "2026-05", precision: "month" }
    expect(chronologyExportDate(date)).toBe("May 2026 · Month only")
    expect(chronologyExportEventNotes(entry(date))[0].text).toBe("Date: 2026-05; May 2026; Date precision: month")
    expect(chronologyExportDate({ original: "2026", start: "2026", precision: "year" })).toBe("2026 · Year only")
    expect(chronologyExportDate({ original: "01 September 2026", start: "2026-09-01", precision: "exact" })).toBe(
      "1 Sep 2026",
    )
  })

  test("does not infer equivalence for source qualifications, ambiguous numeric dates or invalid dates", () => {
    expect(chronologyExportDate({ original: "May or June 2026", start: "2026-05", precision: "ambiguous" })).toBe(
      "May 2026 · May or June 2026 · Ambiguous date",
    )
    expect(chronologyExportDate({ original: "05/06/2026", start: "2026-06-05", precision: "ambiguous" })).toBe(
      "5 Jun 2026 · 05/06/2026 · Ambiguous date",
    )
    expect(chronologyExportDate({ original: "29 February 2026", start: "2026-02-29", precision: "exact" })).toBe(
      "2026-02-29 · 29 February 2026",
    )
  })

  test("keeps uncertainty distinct from a known duration and retains the qualifier", () => {
    const date: ChronologyDate = {
      original: "1 May 2026 to 3 May 2026",
      start: "2026-05-01",
      end: "2026-05-03",
      precision: "approximate",
      rangeKind: "uncertainty",
      qualifier: "No earlier than the first delivery",
    }
    expect(chronologyExportDate(date)).toBe(
      "1 May 2026 to 3 May 2026 · Approximate · Uncertain date range · No earlier than the first delivery",
    )
    expect(chronologyExportDate({ ...date, precision: "exact", rangeKind: "duration" })).toBe(
      "1 May 2026 to 3 May 2026 · Known period · No earlier than the first delivery",
    )
  })

  test("preserves times, absent timezones, relative expressions and stale anchor qualifications", () => {
    const date: ChronologyDate = {
      original: "Three days after receipt",
      start: "2026-05-04",
      precision: "relative",
      time: "09:30",
      timezone: "Europe/London",
      anchorFreshness: "stale",
      anchorIssue: "Receipt date changed",
      relative: {
        expression: "Three days after receipt",
        anchorDocument: "Receipt record.pdf",
        anchorDate: "2026-05-01",
        anchorVersion: "receipt-v1",
        calculation: "2026-05-01 + 3 calendar days",
        interpretation: "document-stated",
      },
    }
    expect(chronologyExportDate(date)).toBe(
      "4 May 2026 · Three days after receipt · Relative date · Time: 09:30 Europe/London · Date anchor stale: Receipt date changed",
    )
    const notes = chronologyExportEventNotes(entry(date)).map((note) => note.text)
    expect(notes).toContain("Anchor: Receipt record.pdf dated 2026-05-01 (version receipt-v1)")
    expect(notes).toContain("Calculation: 2026-05-01 + 3 calendar days")
    expect(notes).toContain("Interpretation: document-stated")
    expect(chronologyExportDate({ ...date, timezone: undefined })).toContain("Time: 09:30 (timezone not supplied)")
  })

  test("retains original expressions for corrected and lawyer-edited entries in evidence notes", () => {
    const date: ChronologyDate = {
      original: "21 May 2026",
      start: "2026-05-21",
      precision: "exact",
      correction: {
        actor: "lawyer",
        at: 1,
        original: { original: "May 2026", start: "2026-05", precision: "month" },
      },
    }
    const event = entry(date)
    event.original = {
      title: "Inspection planned",
      detail: "The original account gave only a year.",
      date: { original: "2026", start: "2026", precision: "year" },
    }
    const notes = chronologyExportEventNotes(event).map((note) => note.text)
    expect(notes).toContain("Date: 2026-05-21; 21 May 2026")
    expect(notes).toContain(
      "Date corrected by lawyer at 1970-01-01T00:00:00.001Z; original: 2026-05; May 2026; Date precision: month",
    )
    expect(notes).toContain("Original date: 2026; Date precision: year")
  })
})
