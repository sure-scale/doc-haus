import { describe, expect, test } from "bun:test"
import { formatAgo, formatCost, formatDate, formatDateTime, formatElapsed } from "../src/lib/format"
import { humanizeError, TIMED_OUT, UNREACHABLE } from "../src/api/ingest/errors"

describe("formatDate", () => {
  const ms = Date.UTC(2026, 9, 11, 12)

  test("renders day, short month and year", () => {
    expect(formatDate(ms)).toBe("11 Oct 2026")
  })

  test("accepts epoch seconds and ISO strings", () => {
    expect(formatDate(ms / 1000)).toBe("11 Oct 2026")
    expect(formatDate("2026-10-11T12:00:00Z")).toBe("11 Oct 2026")
  })

  test("is empty for missing or invalid input", () => {
    expect(formatDate(undefined)).toBe("")
    expect(formatDate(null)).toBe("")
    expect(formatDate("not a date")).toBe("")
  })

  test("date-time keeps the same date form", () => {
    expect(formatDateTime(ms)).toStartWith("11 Oct 2026, ")
  })
})

describe("formatElapsed", () => {
  test("seconds, minutes and hours", () => {
    expect(formatElapsed(0)).toBe("0s")
    expect(formatElapsed(45_900)).toBe("45s")
    expect(formatElapsed(252_000)).toBe("4m 12s")
    expect(formatElapsed(65_000)).toBe("1m 05s")
    expect(formatElapsed(3_900_000)).toBe("1h 05m")
  })

  test("is empty for negative or missing durations", () => {
    expect(formatElapsed(-1)).toBe("")
    expect(formatElapsed(undefined)).toBe("")
    expect(formatElapsed(Number.NaN)).toBe("")
  })
})

describe("formatAgo", () => {
  const now = Date.UTC(2026, 9, 11, 12)

  test("relative steps", () => {
    expect(formatAgo(now - 2_000, now)).toBe("just now")
    expect(formatAgo(now - 12_000, now)).toBe("12s ago")
    expect(formatAgo(now - 4 * 60_000, now)).toBe("4m ago")
    expect(formatAgo(now - 2 * 3_600_000, now)).toBe("2h ago")
    expect(formatAgo(now - 3 * 86_400_000, now)).toBe("8 Oct 2026")
  })

  test("accepts ISO timestamps", () => {
    expect(formatAgo(new Date(now - 30_000).toISOString(), now)).toBe("30s ago")
  })
})

describe("formatCost", () => {
  test("two decimals", () => {
    expect(formatCost(1.23456)).toBe("$1.23")
    expect(formatCost(0)).toBe("$0.00")
  })

  test("tiny costs do not round to zero", () => {
    expect(formatCost(0.0004)).toBe("< $0.01")
  })

  test("unpriced models never show a zero", () => {
    expect(formatCost(0, true)).toBe("Not priced")
    expect(formatCost(2.5, true)).toBe("≥ $2.50")
  })
})

describe("humanizeError", () => {
  test("network failures read as an unreachable service", () => {
    expect(humanizeError(new TypeError("Failed to fetch"))).toBe(UNREACHABLE)
    expect(humanizeError(new TypeError("Load failed"))).toBe(UNREACHABLE)
    expect(humanizeError(new TypeError("NetworkError when attempting to fetch resource."))).toBe(UNREACHABLE)
  })

  test("timeouts and aborts read as a slow service", () => {
    expect(humanizeError(new DOMException("signal timed out", "TimeoutError"))).toBe(TIMED_OUT)
    expect(humanizeError(new DOMException("aborted", "AbortError"))).toBe(TIMED_OUT)
  })

  test("server messages pass through", () => {
    expect(humanizeError(new Error("Only .docx and .pdf files are supported"))).toBe("Only .docx and .pdf files are supported")
    expect(humanizeError("plain text")).toBe("plain text")
  })

  test("unknown values fall back", () => {
    expect(humanizeError(undefined)).toBe("Something went wrong. Try again.")
    expect(humanizeError({}, "Could not save")).toBe("Could not save")
    expect(humanizeError(new Error("  "), "Could not save")).toBe("Could not save")
  })
})
