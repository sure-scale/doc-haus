import { afterEach, expect, test } from "bun:test"
import { mkdtempSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { recordKeyDate } from "../../../dochaus/lib/key-dates"
import { keyDateRoutes } from "./key-dates"

const dirs: string[] = []
afterEach(() => {
  for (const dir of dirs.splice(0)) rmSync(dir, { recursive: true, force: true })
})

function setup() {
  const directory = mkdtempSync(path.join(tmpdir(), "key-dates-route-"))
  dirs.push(directory)
  writeFileSync(path.join(directory, "matter.json"), "{}")
  return { app: keyDateRoutes((id) => (id === "synthetic" ? directory : path.join(directory, "missing"))), directory }
}

test("rejects an invalid matter id and a missing matter", async () => {
  const { app } = setup()
  expect((await app.request("/matters/UPPER/key-dates")).status).toBe(400)
  expect((await app.request("/matters/other/key-dates")).status).toBe(404)
})

test("an empty register is an explicit empty list, not an error", async () => {
  const { app } = setup()
  const response = await app.request("/matters/synthetic/key-dates")
  expect(response.status).toBe(200)
  expect(await response.json()).toEqual({ rows: [] })
})

test("returns stored rows soonest first with unresolved formulas kept distinct", async () => {
  const { app, directory } = setup()
  const base = {
    noticeDays: undefined,
    cancelByIso: undefined,
    party: undefined,
    quote: "clause text fixing the date",
    reason: "test",
    author: "Synthetic Counsel",
  }
  recordKeyDate(directory, {
    ...base,
    docName: "MSA.docx",
    kind: "renewal",
    dateIso: "2027-01-01",
    noticeDays: 90,
    cancelByIso: "2026-10-03",
    party: "Customer",
    quote: "This Agreement renews for successive one-year terms.",
    reason: "Auto-renewal anniversary",
  })
  recordKeyDate(directory, {
    ...base,
    docName: "SOW.docx",
    kind: "deliverable",
    dateFormula: "30 days after acceptance",
    reason: "Delivery window with no anchor date",
  })
  recordKeyDate(directory, { ...base, docName: "SOW.docx", kind: "payment", dateIso: "2026-09-15" })

  const response = await app.request("/matters/synthetic/key-dates")
  expect(response.status).toBe(200)
  const { rows } = (await response.json()) as { rows: Record<string, unknown>[] }
  expect(rows.map((row) => row.document)).toEqual(["SOW.docx", "MSA.docx", "SOW.docx"])
  expect(rows[0]).toMatchObject({ kind: "payment", date: "2026-09-15", formula: null })
  expect(rows[1]).toMatchObject({
    kind: "renewal",
    date: "2027-01-01",
    formula: null,
    noticeDays: 90,
    actBy: "2026-10-03",
    party: "Customer",
    quote: "This Agreement renews for successive one-year terms.",
    reason: "Auto-renewal anniversary",
    author: "Synthetic Counsel",
  })
  expect(rows[2]).toMatchObject({ kind: "deliverable", date: null, formula: "30 days after acceptance" })
})
