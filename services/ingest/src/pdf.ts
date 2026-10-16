import { mkdtempSync, readdirSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"

// Below this many text-layer characters per page the PDF is treated as a flat
// scan (image-only pages with at most stray stamps or bates numbers) and routed
// through OCR.
const SCANNED_CHARS_PER_PAGE = 100

// PDF -> text for indexing. LiteParse is the primary born-digital extractor
// because it emits structured Markdown without Python startup overhead; markitdown
// remains the compatible fallback, and unpdf is the no-Python floor. Flat scans
// yield a near-empty text layer either way, so those fall through to OCR.
export async function pdfToText(pdfBytes: Buffer): Promise<string> {
  const { extractText, getDocumentProxy } = await import("unpdf")
  const pdf = await getDocumentProxy(new Uint8Array(pdfBytes))
  const text = (await liteparseText(pdfBytes)) ?? (await markitdownText(pdfBytes)) ?? (await extractText(pdf, { mergePages: true })).text
  if (!looksScanned(text, pdf.numPages)) return text
  return (await ocrPdfText(pdfBytes)) ?? text
}

export function looksScanned(text: string, pageCount: number) {
  return text.trim().length < SCANNED_CHARS_PER_PAGE * pageCount
}

// OCR for flat scans: rasterize with poppler's pdftoppm (300dpi grayscale), then
// run tesseract once over the full page list. Both tools must already be on
// PATH — we never bundle or download them. Null when either is absent.
export async function ocrPdfText(pdfBytes: Buffer): Promise<string | null> {
  const pdftoppm = Bun.which("pdftoppm")
  const tesseract = Bun.which("tesseract")
  if (!pdftoppm || !tesseract) return null
  const dir = mkdtempSync(path.join(tmpdir(), "dochaus-ocr-"))
  try {
    const src = path.join(dir, "in.pdf")
    writeFileSync(src, pdfBytes)
    const raster = Bun.spawn([pdftoppm, "-r", "300", "-gray", "-png", src, path.join(dir, "page")], {
      stdout: "ignore",
      stderr: "ignore",
    })
    if ((await raster.exited) !== 0) return null
    // pdftoppm zero-pads page numbers to a uniform width, so a plain sort is page order.
    const pages = readdirSync(dir)
      .filter((f) => f.endsWith(".png"))
      .sort()
    if (pages.length === 0) return null
    const list = path.join(dir, "pages.txt")
    writeFileSync(list, pages.map((f) => path.join(dir, f)).join("\n"))
    const ocr = Bun.spawn([tesseract, list, "stdout"], { stdout: "pipe", stderr: "ignore" })
    const text = await new Response(ocr.stdout).text()
    if ((await ocr.exited) !== 0) return null
    return text
  } finally {
    rmSync(dir, { recursive: true, force: true })
  }
}

async function liteparseText(pdfBytes: Buffer): Promise<string | null> {
  return import("@llamaindex/liteparse")
    .then(async ({ LiteParse }) => {
      const result = await new LiteParse({
        outputFormat: "markdown",
        imageMode: "off",
        extractLinks: false,
        ocrEnabled: false,
        quiet: true,
      }).parse(pdfBytes)
      return result.text
    })
    .catch(() => null)
}

async function markitdownText(pdfBytes: Buffer): Promise<string | null> {
  const cmd = markitdownCommand()
  if (!cmd) return null
  const dir = mkdtempSync(path.join(tmpdir(), "dochaus-markitdown-"))
  try {
    const src = path.join(dir, "in.pdf")
    writeFileSync(src, pdfBytes)
    const proc = Bun.spawn([...cmd, src], { stdout: "pipe", stderr: "ignore" })
    const text = await new Response(proc.stdout).text()
    if ((await proc.exited) !== 0) return null
    return text
  } finally {
    rmSync(dir, { recursive: true, force: true })
  }
}

// markitdown is Python, so it is never bundled. Use a PATH install when present,
// else run it through uvx, which caches the venv after the first call. Null when
// neither exists, in which case the caller falls back to unpdf.
function markitdownCommand() {
  const direct = Bun.which("markitdown")
  if (direct) return [direct]
  const uvx = Bun.which("uvx")
  if (uvx) return [uvx, "--quiet", "--from", "markitdown[pdf]", "markitdown"]
  return null
}
