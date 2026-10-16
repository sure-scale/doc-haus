import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { Document, Packer, Paragraph } from "docx"
import { looksScanned, ocrPdfText } from "./pdf"
import { withParagraphPointers } from "../../../dochaus/lib/drafting-amendment"

// Best-effort PDF -> DOCX. There is no MIT pure-JS engine that preserves PDF
// layout, so the reliable floor is text-only: pull the text with unpdf and rebuild
// it as DOCX paragraphs. When LibreOffice is already installed on the host we use
// it instead for full-fidelity conversion (tables, columns, images) — we never
// bundle or download it, only call `soffice` if it happens to be on PATH.
export async function pdfToDocx(pdfBytes: Buffer): Promise<Buffer> {
  const viaSoffice = await sofficeConvert(pdfBytes)
  return Buffer.from(withParagraphPointers(viaSoffice ?? await textOnlyConvert(pdfBytes)))
}

// Lossless-ish path: shell out to a pre-installed LibreOffice. Returns null when
// soffice is absent or the conversion fails, so the caller falls back to text.
async function sofficeConvert(pdfBytes: Buffer): Promise<Buffer | null> {
  const soffice = Bun.which("soffice") ?? Bun.which("libreoffice")
  if (!soffice) return null
  const dir = mkdtempSync(path.join(tmpdir(), "dochaus-convert-"))
  try {
    const src = path.join(dir, "in.pdf")
    writeFileSync(src, pdfBytes)
    const proc = Bun.spawn([soffice, "--headless", "--convert-to", "docx", "--outdir", dir, src], {
      stdout: "ignore",
      stderr: "ignore",
    })
    if ((await proc.exited) !== 0) return null
    const output = path.join(dir, "in.docx")
    if (!existsSync(output)) return null
    return readFileSync(output)
  } finally {
    rmSync(dir, { recursive: true, force: true })
  }
}

// Text-only path: unpdf returns one string per page; each page becomes a run of
// paragraphs split on blank lines so clause breaks survive into the DOCX. Flat
// scans have no text layer to pull, so those go through OCR instead — otherwise
// a converted scanned filing would open as a blank document.
async function textOnlyConvert(pdfBytes: Buffer): Promise<Buffer> {
  const { extractText, getDocumentProxy } = await import("unpdf")
  const pdf = await getDocumentProxy(new Uint8Array(pdfBytes))
  const { text } = await extractText(pdf, { mergePages: false })
  const layerText = text.join("\n")
  const source = looksScanned(layerText, pdf.numPages) ? ((await ocrPdfText(pdfBytes)) ?? layerText) : layerText
  const paragraphs = source.split(/\n/).map((line) => new Paragraph({ text: line.trim() }))
  const doc = new Document({ sections: [{ children: paragraphs }] })
  return Buffer.from(await Packer.toBuffer(doc))
}
