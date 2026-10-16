// Key-dates register guarantees, exercised through the real record-key-date and
// key-dates tools against a temp matter — no mocks beyond the tool ctx.
// Run from dochaus/: bun test test/key-dates.test.ts
//
// Lives in test/, not tool/: the engine's ToolRegistry imports every .ts under
// {tool,tools}/ as a custom tool, and importing a bun:test file outside the
// test runner throws ("Cannot use describe outside of the test runner").
import { describe, expect, test } from "bun:test"
import { mkdtempSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { listKeyDates, recordKeyDate } from "../lib/key-dates"

const recordTool = (await import("../tool/record-key-date")).default
const readTool = (await import("../tool/key-dates")).default

function ctxFor(directory: string) {
  return {
    sessionID: "test",
    messageID: "test",
    agent: "test",
    directory,
    worktree: directory,
    abort: new AbortController().signal,
    metadata() {},
    async ask() {},
  }
}

function matter() {
  return mkdtempSync(path.join(tmpdir(), "key-dates-matter-"))
}

describe("record-key-date tool", () => {
  test("writes a typed row into the matter register", async () => {
    const dir = matter()
    const result = await recordTool.execute(
      {
        document: "msa.docx",
        kind: "renewal",
        dateIso: "2027-01-01",
        dateFormula: undefined,
        noticeDays: 90,
        cancelByIso: "2026-10-03",
        party: "Customer",
        quote: "This Agreement renews automatically for successive one-year terms.",
        reason: "Auto-renewal anniversary",
        author: "A. Lawyer",
      },
      ctxFor(dir),
    )
    expect(result).toBeObject()
    if (typeof result === "string") throw new Error(result)
    expect(result.output).toContain("key date #1")
    expect(result.output).toContain("2027-01-01")

    const rows = listKeyDates(dir)
    expect(rows).toHaveLength(1)
    expect(rows[0].kind).toBe("renewal")
    expect(rows[0].date_iso).toBe("2027-01-01")
    expect(rows[0].date_formula).toBeNull()
    expect(rows[0].notice_days).toBe(90)
    expect(rows[0].cancel_by_iso).toBe("2026-10-03")
    expect(rows[0].party).toBe("Customer")
    expect(rows[0].author).toBe("A. Lawyer")
  })

  test("refuses a row with neither dateIso nor dateFormula", async () => {
    const dir = matter()
    const result = await recordTool.execute(
      {
        document: "msa.docx",
        kind: "expiry",
        dateIso: undefined,
        dateFormula: undefined,
        noticeDays: undefined,
        cancelByIso: undefined,
        party: undefined,
        quote: "The term ends as set forth above.",
        reason: "test",
        author: "test",
      },
      ctxFor(dir),
    )
    expect(result).toBeString()
    expect(result as string).toContain("dateIso")
    expect(listKeyDates(dir)).toEqual([])
  })

  test("accepts a formula-only row when no calendar date is resolvable", async () => {
    const dir = matter()
    const result = await recordTool.execute(
      {
        document: "msa.docx",
        kind: "termination_notice",
        dateIso: undefined,
        dateFormula: "90 days prior to the end of the then-current term",
        noticeDays: 90,
        cancelByIso: undefined,
        party: undefined,
        quote: "Either party may terminate on 90 days' notice prior to the end of the then-current term.",
        reason: "Notice window with no stated anchor date",
        author: "test",
      },
      ctxFor(dir),
    )
    if (typeof result === "string") throw new Error(result)
    const rows = listKeyDates(dir)
    expect(rows[0].date_iso).toBeNull()
    expect(rows[0].date_formula).toBe("90 days prior to the end of the then-current term")
  })
})

describe("key-dates tool", () => {
  test("an empty register returns an explicit all-clear", async () => {
    const result = await readTool.execute({ document: undefined }, ctxFor(matter()))
    if (typeof result === "string") throw new Error(result)
    expect(result.output).toBe("No key dates recorded in this matter.")
    expect(result.title).toBe("0 key date(s)")
  })

  test("lists soonest first with undated rows last, and filters by document", async () => {
    const dir = matter()
    const base = {
      noticeDays: undefined,
      cancelByIso: undefined,
      party: undefined,
      quote: "clause text fixing the date",
      reason: "test",
      author: "test",
    }
    recordKeyDate(dir, { ...base, docName: "msa.docx", kind: "expiry", dateIso: "2027-06-30" })
    recordKeyDate(dir, {
      ...base,
      docName: "sow.docx",
      kind: "deliverable",
      dateFormula: "30 days after acceptance",
    })
    recordKeyDate(dir, { ...base, docName: "sow.docx", kind: "payment", dateIso: "2026-09-15" })

    const all = await readTool.execute({ document: undefined }, ctxFor(dir))
    if (typeof all === "string") throw new Error(all)
    expect(all.metadata?.keyDates).toEqual([3, 1, 2])
    const lines = all.output.split("\n")
    expect(lines[0]).toContain("2026-09-15")
    expect(lines[1]).toContain("2027-06-30")
    expect(lines[2]).toContain('formula: "30 days after acceptance"')

    const filtered = await readTool.execute({ document: "sow.docx" }, ctxFor(dir))
    if (typeof filtered === "string") throw new Error(filtered)
    expect(filtered.metadata?.keyDates).toEqual([3, 2])
    expect(filtered.title).toBe("2 key date(s) in sow.docx")
    expect(filtered.output).not.toContain("msa.docx")
  })
})
