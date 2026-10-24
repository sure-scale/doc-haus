import { existsSync, readFileSync } from "node:fs"
import { fileURLToPath } from "node:url"
import path from "node:path"
import { docxodus } from "./docxodus"
import { withParagraphPointers } from "./drafting-amendment"
import { polishMarkdownArtifactsInDocx } from "./docx-polish"
import { lintDraftMarkdown, splitBlocks } from "./markdown"
import { highlightOpenItemsInDocx } from "./open-items"
import { readFirmPreferences } from "./preferences"
import { normalizePreferences } from "./drafting-preferences"
import { prepareMarkdownRenderText } from "./markdown-render-text"
import { withGeneratedDocxAttribution } from "./docx-attribution"
import { applyPageFurniture, type PageFurniture } from "./page-furniture"
import { numberDocxClauses } from "./docx-numbering"

// Shared markdown -> .docx -> deliverable render core for doc.haus drafting.
// draft-document (from-scratch create) and the lean drafting stages compose a
// markdown body onto the styled blank seed (_base.docx) the same way, so the seed
// insert, the advisory consistency lint, the save, and the deliverable write live
// here once.
//
// Docxodus can edit a .docx but not create one from nothing, so every render
// starts from the seed bytes that carry the house style set (Title, Heading 1-3)
// — that is what makes markdown headings land as real Word styles. Starting from
// raw bytes instead would lose those styles.
//
// The finished bytes go to the ingest service's deliverable route, which writes
// and registers the .docx for the web viewer without treating generated output as
// source evidence. The render core never writes the file itself.

const baseSeed = path.join(path.dirname(fileURLToPath(import.meta.url)), "..", "templates", "_base.docx")
// Default to start.sh's DOC_HAUS_INGEST_PORT (4501) so the tool targets the
// ingest service start.sh launches. Override INGEST_URL to point at a different stack.
const ingestUrl = process.env.INGEST_URL ?? `http://127.0.0.1:${process.env.DOC_HAUS_INGEST_PORT ?? 4501}`

// Compose `content` (markdown) onto the styled seed and return the saved .docx
// bytes plus the advisory consistency lint. Caller owns the deliverable write via
// uploadDeliverable so each tool can name the file and report independently.
// A re-render of an earlier render's unchanged text keeps its paragraph ids.
// `furniture` (page-furniture.ts) adds the header and footer the document's
// form carries. Literal clause numbers become Word list numbering
// (docx-numbering.ts) unless `numbering` is false, for a caller that numbers
// after reshaping the headings itself.
export async function renderMarkdownToDocx(content: string, previous?: Uint8Array, furniture?: PageFurniture, options: { numbering?: false } = {}) {
  const text = prepareMarkdownRenderText(splitBlocks(content))
  const dx = await docxodus()
  const session = dx.openDocxSession(await Bun.file(baseSeed).bytes(), {})

  // The seed's single empty paragraph is the insertion anchor; the whole markdown
  // body goes in as one multi-block insert, then the seed is dropped.
  let saved: Uint8Array
  try {
    const seed = Object.keys(session.project().anchorIndex).find((id) => id.startsWith("p:"))!
    const inserted = session.insertParagraph(seed, "after", text.markdown)
    if (!inserted.success)
      throw new Error(`Render failed: ${inserted.error?.message ?? JSON.stringify(inserted.error)}`)
    session.deleteBlock(seed)
    saved = session.save()
  } finally {
    session.close()
  }
  const composed = text.restore(polishMarkdownArtifactsInDocx(saved))
  const accepted = dx.openDocxSession(composed, {})
  const projection = accepted.project()
  accepted.close()

  // Consistency lint over the composed result (defined-term hygiene,
  // cross-reference integrity, the firm's date/number style). Advisory — the
  // document still ships and the problems come back so they can be fixed. The
  // projection's anchor markers and synthetic "# Document" line are engine
  // bookkeeping, not document text, so they are stripped before linting.
  const lint = lintRenderedDraftMarkdown(
    projection
      .markdown.replace(/\{#[^}]+\} /g, "")
      .split("\n")
      .filter((line) => line !== "# Document")
      .join("\n"),
  )

  // Furniture goes in only now: the restore expects each placeholder once, in
  // document.xml, and a header date would trip the lint's date-style rule.
  // Highlighting and paragraph pointers read document.xml only, so header and
  // footer text takes neither.
  // Numbering follows the lint, which reads the literal numbers as Markdown.
  const numbered = options.numbering === false ? composed : numberDocxClauses(composed)
  const furnished = furniture ? applyPageFurniture(numbered, furniture) : numbered
  const bytes = withGeneratedDocxAttribution(withParagraphPointers(highlightOpenItemsInDocx(furnished), previous))
  return { bytes, lint }
}

export function lintRenderedDraftMarkdown(markdown: string) {
  return lintDraftMarkdown(markdown, normalizePreferences(readFirmPreferences()))
}

// Upload finished .docx bytes through ingest's non-source deliverable boundary.
// The file remains visible in the matter without entering retrieval or source maps.
export async function uploadDeliverable(
  directory: string,
  name: string,
  bytes: Uint8Array,
  precondition?:
    | { kind: "create" }
    | { kind: "replace"; prior_sha256: string },
) {
  const matter = JSON.parse(readFileSync(path.join(directory, "matter.json"), "utf8")) as { id: string }
  const form = new FormData()
  form.append("file", new File([new Uint8Array(bytes)], name))
  const headers = precondition ? new Headers() : undefined
  if (precondition?.kind === "create")
    headers!.set("x-dochaus-deliverable-precondition", "create")
  if (precondition?.kind === "replace") {
    headers!.set("x-dochaus-deliverable-precondition", "replace")
    headers!.set(
      "x-dochaus-deliverable-prior-sha256",
      precondition.prior_sha256,
    )
  }
  const res = await fetch(`${ingestUrl}/matters/${matter.id}/deliverables`, {
    method: "POST",
    body: form,
    headers,
  })
  if (!res.ok) return `deliverable upload failed (${res.status}): ${await res.text()}`
  if (!existsSync(path.join(directory, name))) return `deliverable upload succeeded but did not write ${name} into this matter`
  return undefined
}
