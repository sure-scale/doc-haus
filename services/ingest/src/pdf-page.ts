import { existsSync } from "node:fs"
import path from "node:path"
import { Hono } from "hono"
import { withArtifactMutationLocksAsync } from "../../../dochaus/capabilities/artifacts/mutation-lock"
import { deliverableDocumentName } from "./deliverable"
import { MatterDocumentIdentityError, resolveMatterDocumentTarget } from "./document-identity"
import { matterDocumentName } from "./ingest"
import { looksScanned } from "./pdf"

export type PdfPassagePage =
  | { status: "matched"; page: number }
  | { status: "ambiguous"; pages: number[] }
  | { status: "unavailable"; reason: "no_text_layer" | "not_found" | "unreadable" }

// A citation excerpt is chunk text; this is far above any chunk.
export const MAX_EXCERPT_CHARS = 20_000

// The shortest leading run of an excerpt (letters and digits) taken as evidence
// of where it starts when the whole excerpt is not found verbatim. A shorter
// excerpt must be found whole.
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
    const name = matterDocumentName(input.name) ?? deliverableDocumentName(input.name)
    if (!name || path.extname(name).toLowerCase() !== ".pdf") return c.notFound()
    const dir = resolveMatter(c.req.param("id"))
    const bytes = await withArtifactMutationLocksAsync(dir, [name], async () => {
      const target = documentTarget(dir, name)
      return target && existsSync(target.docPath) ? Bun.file(target.docPath).bytes() : undefined
    })
    if (!bytes) return c.notFound()
    return c.json(await pdfPassagePage(bytes, input.excerpt))
  })
  return app
}

function documentTarget(directory: string, name: string) {
  try {
    return resolveMatterDocumentTarget(directory, name)
  } catch (error) {
    if (error instanceof MatterDocumentIdentityError) return
    throw error
  }
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

// The page where the excerpt starts. The pages are joined, so an excerpt that
// runs onto the next page still matches whole. When the whole excerpt is not in
// the text layer (a running header or page number interrupts it, or the indexed
// projection differs further on), its longest leading part that is found counts,
// if it is at least MIN_LEADING_CHARS long. Starts on more than one page are
// ambiguous.
export function locatePdfPassage(pages: readonly string[], excerpt: string): PdfPassagePage {
  const keys = pages.map(textKey)
  const starts: number[] = []
  let length = 0
  for (const key of keys) {
    starts.push(length)
    length += key.length
  }
  const text = keys.join("")
  const needle = textKey(excerpt)
  let found = Math.min(needle.length, MIN_LEADING_CHARS)
  if (!needle || !text.includes(needle.slice(0, found))) return { status: "unavailable", reason: "not_found" }
  // Every shorter lead of a found lead is found too, so search for the longest.
  let high = needle.length
  while (found < high) {
    const middle = Math.ceil((found + high) / 2)
    if (text.includes(needle.slice(0, middle))) found = middle
    else high = middle - 1
  }
  const lead = needle.slice(0, found)
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
