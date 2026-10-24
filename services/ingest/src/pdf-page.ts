import { existsSync } from "node:fs"
import path from "node:path"
import { Hono } from "hono"
import { withArtifactMutationLocksAsync } from "../../../dochaus/capabilities/artifacts/mutation-lock"
import { looksScanned } from "./pdf"
import { readDocumentTarget, registeredDocumentName } from "./registered-document"

export type PdfPassagePage =
  | { status: "matched"; page: number }
  | { status: "ambiguous"; pages: number[] }
  | { status: "unavailable"; reason: "no_text_layer" | "not_found" | "unreadable" }

// A citation excerpt is chunk text; this is far above any chunk.
export const MAX_EXCERPT_CHARS = 20_000

// The opening of an excerpt (letters and digits) whose start decides the page
// when the whole excerpt is not found. A shorter excerpt must be found whole.
const MIN_LEADING_CHARS = 24

// The page of a matter PDF where a cited excerpt starts, so the web viewer can
// open the browser's PDF viewer at it (#page=N). Deterministic: the excerpt is
// matched against the PDF's own text layer, page by page; nothing is guessed.
// The document resolves exactly as /documents/content does: a registered
// basename only, never a path.
export function pdfPassageRoutes(resolveMatter: (id: string) => string) {
  const app = new Hono()
  app.post("/matters/:id/documents/pdf-page", async (c) => {
    const input = await c.req.json<{ name?: unknown; excerpt?: unknown }>().catch(() => null)
    if (!input || typeof input.name !== "string" || typeof input.excerpt !== "string" || !input.excerpt.trim())
      return c.json({ error: "Send the PDF's name and the cited excerpt." }, 400)
    if (input.excerpt.length > MAX_EXCERPT_CHARS)
      return c.json({ error: `The excerpt is longer than ${MAX_EXCERPT_CHARS} characters.` }, 400)
    const name = registeredDocumentName(input.name)
    if (!name || path.extname(name).toLowerCase() !== ".pdf") return c.notFound()
    const dir = resolveMatter(c.req.param("id"))
    const bytes = await withArtifactMutationLocksAsync(dir, [name], async () => {
      const target = readDocumentTarget(dir, name)
      return target && existsSync(target.docPath) ? Bun.file(target.docPath).bytes() : undefined
    })
    if (!bytes) return c.notFound()
    return c.json(await pdfPassagePage(bytes, input.excerpt))
  })
  return app
}

// A flat scan (the same test ingest uses to route a PDF through OCR) has no text
// layer the excerpt could be found in, and the browser could not find it either.
export async function pdfPassagePage(bytes: Uint8Array, excerpt: string): Promise<PdfPassagePage> {
  const { extractText, getDocumentProxy } = await import("unpdf")
  const pages = await getDocumentProxy(bytes)
    .then(async (pdf) => (await extractText(pdf, { mergePages: false })).text)
    .catch(() => undefined)
  if (!pages) return { status: "unavailable", reason: "unreadable" }
  if (looksScanned(pages.join("\n"), pages.length)) return { status: "unavailable", reason: "no_text_layer" }
  return locatePdfPassage(pages, excerpt)
}

// Compare letters and digits only, NFKC-folded and lower-cased. The indexed text
// the excerpt comes from (LiteParse Markdown, markitdown or unpdf) and the PDF's
// text layer differ in whitespace, line-end hyphenation, Markdown markup, quote
// marks and ligatures, not in the words.
function textKey(text: string) {
  return text.normalize("NFKC").toLowerCase().replace(/[^\p{L}\p{N}]+/gu, "")
}

// The page where the excerpt starts. Running headers and footers are dropped and
// the pages joined, so an excerpt that runs onto the next page still matches
// whole. When the whole excerpt is not in the text layer (the indexed projection
// differs further on), only its first MIN_LEADING_CHARS count: a longer opening
// can be shared by a draft and its executed copy. Starts on more than one page
// are ambiguous.
export function locatePdfPassage(pages: readonly string[], excerpt: string): PdfPassagePage {
  const keys = withoutRunningLines(pages).map(textKey)
  const starts: number[] = []
  let length = 0
  for (const key of keys) {
    starts.push(length)
    length += key.length
  }
  const text = keys.join("")
  const needle = textKey(excerpt)
  const lead = text.includes(needle) ? needle : needle.slice(0, MIN_LEADING_CHARS)
  if (!needle || !text.includes(lead)) return { status: "unavailable", reason: "not_found" }
  const matches = new Set<number>()
  let page = 0
  for (let at = text.indexOf(lead); at >= 0; at = text.indexOf(lead, at + 1)) {
    // The page holding this offset: the last one starting at or before it (an
    // empty page shares its start with the next page). Offsets only increase.
    while (page + 1 < starts.length && starts[page + 1]! <= at) page++
    matches.add(page + 1)
  }
  const ordered = [...matches].sort((a, b) => a - b)
  return ordered.length === 1 ? { status: "matched", page: ordered[0]! } : { status: "ambiguous", pages: ordered }
}

// The text layer keeps running headers, footers and page numbers that the
// indexed text drops. Treat a line among the first or last two of a page as one
// when, digits aside, it is at the top or foot of most pages, and of three or
// more (two pages that open with the same heading are not evidence).
function withoutRunningLines(pages: readonly string[]) {
  const lines = pages.map((page) => page.split("\n").filter((line) => line.trim()))
  const edges = lines.map((page) => new Set([...page.slice(0, 2), ...page.slice(-2)]))
  const signature = (line: string) => textKey(line).replace(/\p{N}+/gu, "")
  const counts = new Map<string, number>()
  for (const edge of edges)
    for (const sign of new Set([...edge].map(signature))) counts.set(sign, (counts.get(sign) ?? 0) + 1)
  const running = (line: string) => {
    const count = counts.get(signature(line)) ?? 0
    return count >= 3 && count > pages.length / 2
  }
  return lines.map((page, index) => page.filter((line) => !(edges[index]!.has(line) && running(line))).join("\n"))
}
