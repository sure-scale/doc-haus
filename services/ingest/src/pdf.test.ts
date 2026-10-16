import { test, expect } from "bun:test"
import { readFileSync } from "node:fs"
import path from "node:path"
import { extractText, getDocumentProxy } from "unpdf"
import { pdfToText, looksScanned, ocrPdfText } from "./pdf"
import { sectionize } from "./ingest"

// Authored source and provenance are committed beside these generated fixtures.
// The scan rasterizes only our text-layer fixture and has no embedded OCR text.
const fixture = (name: string) => readFileSync(path.join(import.meta.dir, "fixtures", `synthetic-${name}.pdf`))
const textPdf = fixture("text-layer")
const scannedPdf = fixture("scanned")
const continuationPdf = fixture("continuation")
const tablePdf = fixture("table")

test("pdfToText extracts structured headings and complete born-digital text", async () => {
  const text = await pdfToText(textPdf)
  expect(text).toContain("# SYNTHETIC WORKSHOP PLAN")
  expect(text).toContain("## 2.3 Schedule changes")
  expect(text).toContain("The organizer sends a written reminder before changing the workshop time.")
  expect(text).toContain("leave the demonstration desk clear")
  expect(text).not.toContain("(cid:")
}, 120_000)

test("pdfToText preserves a continued instruction across pages in its section", async () => {
  const text = await pdfToText(continuationPdf)
  const opening = "The carrier places the triangular counters into the green tray before"
  const continuation = "placing the circular counters into the amber tray."
  expect(text).toContain("# SYNTHETIC EQUIPMENT HANDOVER")
  expect(text.indexOf(opening)).toBeGreaterThanOrEqual(0)
  expect(text.indexOf(continuation)).toBeGreaterThan(text.indexOf(opening))
  const sections = sectionize(text)
  const transport = sections.find(section => section.label === "2")
  expect(transport).toBeDefined()
  expect(transport!.text).toContain(opening)
  expect(transport!.text).toContain(continuation)
  expect(transport!.text).not.toContain("The recipient compares both trays")
  expect(sections.find(section => section.label === "3")?.text).toContain("The recipient compares both trays")
}, 120_000)

test("pdfToText preserves table columns, row values and form labels", async () => {
  const text = await pdfToText(tablePdf)
  expect(text).toContain("## Appendix B")
  expect(text).toContain("| CHECK | EXPECTED RESULT |")
  expect(text).toContain("|---|---|")
  expect(text).toContain("| Tray colour | Amber |")
  expect(text).toContain("| Counter shape | Triangle |")
  expect(text).toContain("| Storage location | Upper shelf |")
  expect(text).toContain("| Checked by: [name] | Reviewed by: [name] |")
}, 120_000)

test("scan fixture is image-only and cannot pass extraction from a hidden text layer", async () => {
  const document = await getDocumentProxy(new Uint8Array(scannedPdf))
  expect(document.numPages).toBe(1)
  expect((await extractText(document, { mergePages: true })).text.trim()).toBe("")
})

test("pdfToText recovers the image-only workshop plan through OCR", async () => {
  const text = await pdfToText(scannedPdf)
  expect(text).toContain("SYNTHETIC WORKSHOP PLAN")
  expect(text).toContain("The organizer sends a written reminder before changing the workshop time.")
  expect(text).toContain("leave the demonstration desk clear")
}, 120_000)

test("ocrPdfText reads image-only headings and body text", async () => {
  const text = await ocrPdfText(scannedPdf)
  expect(text).not.toBeNull()
  expect(text).toContain("Closing checklist")
  expect(text).toContain("Count the reusable trays")
}, 120_000)

test("looksScanned flags empty and sparse text layers only", () => {
  expect(looksScanned("", 1)).toBe(true)
  expect(looksScanned("Reference mark 001", 3)).toBe(true)
  expect(looksScanned("x".repeat(5000), 3)).toBe(false)
})

test("sectionize splits markdown headings and retains their body text", () => {
  const sections = sectionize([
    "# SYNTHETIC WORKSHOP PLAN",
    "This invented plan describes a paper-folding workshop.",
    "## 2.3 Schedule changes",
    "The organizer sends a written reminder.",
  ].join("\n"))
  expect(sections.map(section => section.label)).toEqual(["SYNTHETIC WORKSHOP PLAN", "2.3"])
  expect(sections[1]!.text).toContain("sends a written reminder")
})

test("sectionize retains short form labels without creating spurious sections", () => {
  const sections = sectionize([
    "# SYNTHETIC WORKSHOP CHECKLIST",
    "# AND",
    "# BOX 42",
    "# Appendix B",
    "Workshop materials information",
  ].join("\n"))
  expect(sections.map(section => section.label)).toEqual(["SYNTHETIC WORKSHOP CHECKLIST", "Appendix B"])
  expect(sections[0]!.text).toContain("# AND")
  expect(sections[0]!.text).toContain("# BOX 42")
})
