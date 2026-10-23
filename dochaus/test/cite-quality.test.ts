import { afterAll, expect, test } from "bun:test"
import { mkdtempSync, realpathSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { utils, write, type WorkSheet } from "xlsx"
import { ingestDocument } from "../../services/ingest/src/ingest"
import cite from "../tool/cite"

// cite reads the extraction quality ingest stored for the document (#239).
// Every workbook here is synthetic, built in the test.

const directories: string[] = []
afterAll(() => directories.forEach((directory) => rmSync(directory, { recursive: true, force: true })))

function workbook(sheet: WorkSheet) {
  const book = utils.book_new()
  utils.book_append_sheet(book, sheet, "Sheet1")
  return Buffer.from(write(book, { type: "buffer", bookType: "xlsx" }))
}

// Formulas saved without calculated values; `label` adds a text cell beside them.
function formulas(label?: string): WorkSheet {
  const sheet: WorkSheet = { "!ref": label ? "A1:B2" : "B1:B2", B1: { t: "n", f: "SUM(C1:C9)" }, B2: { t: "n", f: "B1*2" } }
  if (label) sheet.A1 = { t: "s", v: label }
  return sheet
}

const QUOTE = "Synthetic deposit is payable within thirty days"

async function citeIn(name: string, sheet: WorkSheet, quote = QUOTE) {
  const directory = realpathSync(mkdtempSync(path.join(tmpdir(), "dochaus-cite-quality-")))
  directories.push(directory)
  await ingestDocument(directory, name, workbook(sheet))
  const result = await cite.execute(
    { docPath: name, documentName: name, quote, reason: "Shows the deposit term.", confidence: 5 },
    { directory, sessionID: crypto.randomUUID() } as never,
  )
  return typeof result === "string" ? { output: result, citations: undefined } : { output: result.output, citations: result.metadata?.citations }
}

test("cite refuses a document with no readable text", async () => {
  const { output, citations } = await citeIn("Unreadable.xlsx", formulas(), "=SUM(C1:C9)")
  expect(output).toContain("has no readable text")
  expect(output).toContain("Do not use this quotation")
  expect(citations).toBeUndefined()
})

test("cite verifies a quote in a degraded document with the warning and without 'safe to quote'", async () => {
  const { output, citations } = await citeIn("Partly computed.xlsx", formulas(QUOTE))
  expect(output).toContain(`\n${QUOTE}\n`)
  expect(output).toContain("[⚠ DEGRADED EXTRACTION]")
  expect(output).toContain("may be incomplete or garbled")
  expect(output).not.toContain("safe to quote")
  expect(citations).toEqual([expect.objectContaining({ verified: true, degradedExtraction: ["2 of 3 filled cells are formulas saved without a calculated value"] })])
})

test("cite on a document whose extraction is ok is unchanged", async () => {
  const { output, citations } = await citeIn("Terms.xlsx", utils.aoa_to_sheet([[QUOTE], ["Synthetic notices are given in writing."]]))
  expect(output).toContain("safe to quote")
  expect(output).not.toContain("DEGRADED EXTRACTION")
  expect(citations).toEqual([expect.not.objectContaining({ degradedExtraction: expect.anything() })])
})
