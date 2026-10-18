import { afterEach, expect, test } from "bun:test"
import { writeFileSync, symlinkSync } from "node:fs"
import path from "node:path"
import { unzipSync, strFromU8 } from "fflate"
import { docxodus } from "../../lib/docxodus"
import { renderMarkdownToDocx } from "../../lib/render"
import { comparisonBaseline, comparisonDescription, comparisonBasis } from "./comparison"
import { writeLeanPlan } from "./plan"
import { writeLeanBytes } from "./state"
import { cleanupMatters, matter, SOURCE } from "./test-support"

afterEach(cleanupMatters)

const document = { title: "Supply note", purpose: "Record terms", parts: [{ heading: "Terms", scope: "Operative terms" }] }

test("template reuse does not select a source comparison baseline", () => {
  const directory = matter()
  writeLeanPlan(directory, [{ ...document, template: SOURCE }])
  writeLeanBytes(directory, "d1/body.docx", new Uint8Array([1, 2, 3]))
  writeFileSync(path.join(directory, SOURCE), new Uint8Array([4, 5, 6]))
  expect(comparisonBaseline(directory, "d1").bytes).toEqual(Buffer.from([1, 2, 3]))
  expect(comparisonDescription(directory, "d1")).toContain("our first draft")
  expect(comparisonDescription(directory, "d1")).not.toMatch(/\brun\b|generated/)
})

test("explicit comparison uses supplied original bytes and rejects an escaping source", () => {
  const directory = matter()
  writeLeanPlan(directory, [{ ...document, comparison: SOURCE }])
  writeFileSync(path.join(directory, SOURCE), new Uint8Array([4, 5, 6]))
  expect(comparisonBaseline(directory, "d1").bytes).toEqual(Buffer.from([4, 5, 6]))
  expect(comparisonDescription(directory, "d1")).toContain(SOURCE)
  const before = comparisonBasis(directory, "d1")
  writeFileSync(path.join(directory, SOURCE), new Uint8Array([4, 5, 7]))
  expect(comparisonBasis(directory, "d1")).not.toBe(before)
  const second = matter()
  writeFileSync(path.join(second, SOURCE), new Uint8Array([7]))
  writeLeanPlan(second, [{ ...document, comparison: SOURCE }])
  const escape = matter()
  writeLeanPlan(escape, [{ ...document, comparison: SOURCE }])
  symlinkSync(path.join(second, SOURCE), path.join(escape, SOURCE))
  expect(() => comparisonBaseline(escape, "d1")).toThrow("inside the matter")
})

test("Docxodus compares an unbookmarked supplied original against revised generated wording", async () => {
  const directory = matter()
  const original = (await renderMarkdownToDocx("# Supply note\n\nThe supplier delivers within twenty days.")).bytes
  const revised = (await renderMarkdownToDocx("# Supply note\n\nThe supplier delivers within thirty days.")).bytes
  writeLeanPlan(directory, [{ ...document, comparison: SOURCE }])
  writeFileSync(path.join(directory, SOURCE), original)
  const compared = await (await docxodus()).compareDocuments(comparisonBaseline(directory, "d1").bytes, revised, { authorName: "doc.haus" })
  const xml = strFromU8(unzipSync(compared)["word/document.xml"]!)
  expect(xml).toContain("<w:del")
  expect(xml).toContain("<w:ins")
  expect(xml).toContain("twenty")
  expect(xml).toContain("thirty")
}, 120_000)
