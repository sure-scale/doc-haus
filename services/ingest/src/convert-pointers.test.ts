import { expect, test } from "bun:test"
import { pdfToDocx } from "./convert"
import { readDocxCommentParagraphs } from "../../../dochaus/lib/docx-comments"
import { withParagraphPointers } from "../../../dochaus/lib/drafting-amendment"

test("a new DOCX converted from a PDF has body pointers while the supplied PDF remains unchanged", async () => {
  const wording = "This is a plain record of supplied information for conversion. Keep its words intact in the editable copy."
  const stream = `BT /F1 12 Tf 40 700 Td (${wording}) Tj ET`
  const objects = [
    "<< /Type /Catalog /Pages 2 0 R >>",
    "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
    "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>",
    `<< /Length ${stream.length} >>\nstream\n${stream}\nendstream`,
    "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
  ]
  let pdf = "%PDF-1.4\n"
  const offsets = objects.map((object, index) => {
    const offset = pdf.length
    pdf += `${index + 1} 0 obj\n${object}\nendobj\n`
    return offset
  })
  const xref = pdf.length
  pdf += `xref\n0 6\n0000000000 65535 f \n${offsets.map(offset => `${String(offset).padStart(10, "0")} 00000 n \n`).join("")}trailer\n<< /Size 6 /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`
  const input = Buffer.from(pdf)
  const before = Buffer.from(input)
  const output = await pdfToDocx(input)
  const paragraphs = readDocxCommentParagraphs(output)
  expect(paragraphs.map(row => row.paragraph_text).join(" ").replace(/\s+/g, " ")).toBe(wording)
  expect(paragraphs.every(row => row.pids.length === 1)).toBe(true)
  expect(new Set(paragraphs.flatMap(row => row.pids)).size).toBe(paragraphs.length)
  expect(withParagraphPointers(output)).toBe(output)
  expect(input).toEqual(before)
}, 120000)
