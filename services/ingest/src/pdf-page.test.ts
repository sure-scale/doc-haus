import { afterEach, expect, test } from "bun:test"
import { mkdtempSync, readFileSync, rmSync, symlinkSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { readPdf } from "./pdf"
import { locatePdfPassage, MAX_EXCERPT_CHARS, pdfPassagePage, pdfPassageRoutes } from "./pdf-page"

// A five-page born-digital PDF of invented supply terms: one Helvetica line per
// text row and a running "Page N of 5" footer, which interrupts a sentence that
// runs from page 3 onto page 4 and separates page 4's last heading from its
// clause on page 5. Pages 2 and 5 share one sentence.
const PAGES = [
  ["SYNTHETIC SUPPLY TERMS", "1. Scope", "The supplier lends garden furniture samples for a showroom display.", "Each sample stays the property of the supplier."],
  ["2. Delivery", "Each delivery is checked against the packing list on arrival.", "Late deliveries are recorded in the showroom log."],
  ["3. Liability", "The supplier's total liabil-", "ity under these terms is capped at the fees paid.", "The buyer returns any rejected sample to the supplier"],
  ["within ten working days of the inspection.", "4. Payment", "Invoices are payable within thirty days of receipt.", "5. Notices"],
  ["Any notice under these terms is given in writing.", "6. Returns", "Each delivery is checked against the packing list on arrival.", "Returned samples are wrapped in the original covers."],
]

function syntheticPdf(pages: string[][]) {
  const objects = ["<< /Type /Catalog /Pages 2 0 R >>", `<< /Type /Pages /Kids [${pages.map((_, i) => `${4 + i * 2} 0 R`).join(" ")}] /Count ${pages.length} >>`, "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>"]
  pages.forEach((lines, i) => {
    const rows = [...lines.map((line, row) => [line, 740 - row * 22] as const), [`Page ${i + 1} of ${pages.length}`, 40] as const]
    const content = rows.map(([line, y]) => `BT /F1 12 Tf 56 ${y} Td (${line}) Tj ET`).join("\n")
    objects.push(`<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 3 0 R >> >> /Contents ${5 + i * 2} 0 R >>`, `<< /Length ${content.length} >>\nstream\n${content}\nendstream`)
  })
  let body = "%PDF-1.4\n"
  const offsets: number[] = []
  objects.forEach((object, i) => { offsets.push(body.length); body += `${i + 1} 0 obj\n${object}\nendobj\n` })
  const xref = body.length
  body += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n${offsets.map((o) => `${String(o).padStart(10, "0")} 00000 n \n`).join("")}trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`
  return new Uint8Array(Buffer.from(body, "latin1"))
}

const termsPdf = syntheticPdf(PAGES)
// pdf.js takes ownership of (detaches) the bytes it is given.
const terms = () => termsPdf.slice()
const fixture = (name: string) => new Uint8Array(readFileSync(path.join(import.meta.dir, "fixtures", `synthetic-${name}.pdf`)))
const folders: string[] = []
afterEach(() => { folders.splice(0).forEach((folder) => rmSync(folder, { recursive: true, force: true })) })

test("an excerpt on page 3 opens page 3, across a line-end hyphen", async () => {
  expect(await pdfPassagePage(terms(), "The supplier's total liability under these terms is capped at the fees paid.")).toEqual({ status: "matched", page: 3 })
})

test("an excerpt that runs from page 3 onto page 4 opens page 3, with or without the footer between", async () => {
  const sentence = "The buyer returns any rejected sample to the supplier within ten working days of the inspection."
  expect(await pdfPassagePage(terms(), sentence)).toEqual({ status: "matched", page: 3 })
  expect(await pdfPassagePage(terms(), sentence.replace("supplier ", "supplier\nPage 3 of 5\n"))).toEqual({ status: "matched", page: 3 })
})

// The indexed text drops the running footer and marks the page break.
test("a heading at the foot of page 4 with its clause on page 5 opens page 4, past the running footer", async () => {
  expect(await pdfPassagePage(terms(), "## 5. Notices\n\n-----\n\nAny notice under these terms is given in writing.")).toEqual({ status: "matched", page: 4 })
})

// A bundle with a draft and the executed copy: the draft has the clause's
// opening whole on page 3; the executed clause runs from page 9 onto page 10
// with a footer between, so only its opening is found, on both pages.
test("an excerpt found only by its opening, which starts on two pages, is ambiguous", () => {
  const pages = ["Cover", "Contents", "Schedule 2 (draft)\n7. Liability\nThe Supplier's total liability under this Agreement is capped at one million pounds in aggregate.\nPage 3 of 12", "a", "b", "c", "d", "e", "7. Liability\nThe Supplier's total liability under this\nPage 9 of 12", "Agreement is capped at five million pounds in aggregate.\nPage 10 of 12", "x", "y"]
  expect(locatePdfPassage(pages, "7. Liability\n\nThe Supplier's total liability under this\n\n-----\n\nAgreement is capped at five million pounds in aggregate.")).toEqual({ status: "ambiguous", pages: [3, 9] })
})

test("accented text in another Unicode form and Cyrillic text are found", () => {
  const french = "La responsabilité du fournisseur est limitée au montant payé."
  expect(locatePdfPassage(["Préambule", french.normalize("NFD")], french)).toEqual({ status: "matched", page: 2 })
  const russian = "Ответственность поставщика ограничена суммой уплаченных сборов."
  expect(locatePdfPassage(["Предмет договора", russian], russian)).toEqual({ status: "matched", page: 2 })
})

test("a file that is not a readable PDF is unavailable", async () => {
  expect(await pdfPassagePage(new Uint8Array(Buffer.from("This file is not a PDF.")), "The supplier's total liability")).toEqual({ status: "unavailable", reason: "unreadable" })
})

test("text that is not in the PDF is unavailable", async () => {
  expect(await pdfPassagePage(terms(), "The landlord may terminate the lease on one month's notice.")).toEqual({ status: "unavailable", reason: "not_found" })
})

test("text repeated on two pages is ambiguous unless the rest of the excerpt settles it", async () => {
  expect(await pdfPassagePage(terms(), "Each delivery is checked against the packing list on arrival.")).toEqual({ status: "ambiguous", pages: [2, 5] })
  expect(await pdfPassagePage(terms(), "Each delivery is checked against the packing list on arrival. Late deliveries are recorded")).toEqual({ status: "matched", page: 2 })
})

test("a scan with no text layer is unavailable, even for text OCR indexed", async () => {
  expect(await pdfPassagePage(fixture("scanned"), "The organizer sends a written reminder before changing the workshop time.")).toEqual({ status: "unavailable", reason: "no_text_layer" })
})

// The excerpt comes from the indexed projection (LiteParse Markdown), which drops
// the running footer and marks the page break; the PDF's text layer keeps it.
test("an indexed excerpt crossing a real page break opens the page where it starts", async () => {
  const bytes = fixture("continuation")
  const { text } = await readPdf(Buffer.from(bytes))
  const start = text.indexOf("The carrier places the triangular counters")
  expect(start).toBeGreaterThanOrEqual(0)
  expect(await pdfPassagePage(bytes, text.slice(start, start + 200))).toEqual({ status: "matched", page: 1 })
}, 120_000)

test("the route resolves only a registered PDF name in the matter and validates the request", async () => {
  const folder = mkdtempSync(path.join(tmpdir(), "pdf-page-"))
  folders.push(folder)
  writeFileSync(path.join(folder, "Supply terms.pdf"), terms())
  writeFileSync(path.join(folder, "Supply terms.docx"), "not a pdf")
  symlinkSync(path.join(folder, "Supply terms.pdf"), path.join(folder, "Alias.pdf"))
  const app = pdfPassageRoutes(() => folder)
  const post = (body: unknown) => app.request("/matters/example/documents/pdf-page", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) })
  const excerpt = "Invoices are payable within thirty days of receipt."
  const found = await post({ name: "Supply terms.pdf", excerpt })
  expect(found.status).toBe(200)
  expect(await found.json()).toEqual({ status: "matched", page: 4 })
  for (const name of ["Missing.pdf", "../Supply terms.pdf", `${folder}/Supply terms.pdf`, "Supply terms.docx", "Alias.pdf"])
    expect({ name, status: (await post({ name, excerpt })).status }).toEqual({ name, status: 404 })
  for (const body of [{ name: "Supply terms.pdf" }, { name: "Supply terms.pdf", excerpt: "  " }, { excerpt }, { name: "Supply terms.pdf", excerpt: "x".repeat(MAX_EXCERPT_CHARS + 1) }])
    expect((await post(body)).status).toBe(400)
  expect((await app.request("/matters/example/documents/pdf-page?name=Supply%20terms.pdf")).status).toBe(404)
})
