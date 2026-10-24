import { existsSync, readFileSync } from "node:fs"
import { fileURLToPath } from "node:url"
import path from "node:path"
import { strFromU8, unzipSync } from "fflate"
import { docxodus } from "./docxodus"
import { withParagraphPointers } from "./drafting-amendment"
import { polishMarkdownArtifactsInDocx } from "./docx-polish"
import { lintDraftMarkdown, splitBlocks } from "./markdown"
import { highlightOpenItemsInDocx } from "./open-items"
import { readFirmPreferences } from "./preferences"
import { normalizePreferences } from "./drafting-preferences"
import { prepareMarkdownRenderText } from "./markdown-render-text"
import { withGeneratedDocxAttribution } from "./docx-attribution"
import { applyPageFurniture, type DocumentForm, type PageFurniture } from "./page-furniture"
import { numberDocxClauses } from "./docx-numbering"
import { linkCrossReferences } from "./docx-crossrefs"
import { advisoryForm } from "../workflows/lean/consistency"
import { templateSeed, withTemplateStyleIds } from "./template-seed"

// Shared markdown -> .docx -> deliverable render core for doc.haus drafting.
// draft-document (from-scratch create) and the lean drafting stages compose a
// markdown body onto the styled blank seed (_base.docx) the same way, so the seed
// insert, the advisory consistency lint, the save, and the deliverable write live
// here once.
//
// Docxodus can edit a .docx but not create one from nothing, so every render
// starts from the seed bytes that carry the house style set (Title, Heading 1-3)
// — that is what makes markdown headings land as real Word styles. Starting from
// raw bytes instead would lose those styles. A document drafted from a template
// or precedent .docx, or for a firm that chose a library template as its style
// (Settings → Drafting), starts from that file's styles, numbering, fonts and
// page set-up instead (template-seed.ts).
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
// after reshaping the headings itself; that caller also points the headings at
// a template's own style ids (withTemplateStyleIds), which the reshaping does
// not read. `styleTemplate` is the .docx whose styles the document takes, or
// false for the house style; without one the firm's style template applies,
// and without that, or when the file is not a Word document, the house style;
// `templated` says whether a template's styles were used. Headings take the
// house spacing directly (`headingSpacing`) unless a template's own styles
// space them; a template whose own headings carry spacing directly (a document
// doc.haus rendered in the house style, revised from itself) keeps it.
// A numbered document's references to its own clauses become live
// cross-references (docx-crossrefs.ts) only when its `form` is known and is
// not advice or a letter, which cite the clauses of the documents they discuss;
// without a form they stay typed.
export async function renderMarkdownToDocx(content: string, previous?: Uint8Array, furniture?: PageFurniture, options: { numbering?: false; form?: DocumentForm; styleTemplate?: Uint8Array | false } = {}) {
  const text = prepareMarkdownRenderText(splitBlocks(content))
  const dx = await docxodus()
  const house = await Bun.file(baseSeed).bytes()
  const styleTemplate = options.styleTemplate === false ? undefined : options.styleTemplate ?? firmStyleTemplate()
  const templated = styleTemplate && templateSeed(styleTemplate, house)
  if (styleTemplate && !templated) warnOnce("a style template that is not a Word document with styles of its own; the house style applies")
  const headingSpacing = !templated || headingsSpaced(styleTemplate!)
  const session = dx.openDocxSession(templated || house, {})

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
  // The house headings carry their spacing directly; a template's keep its own.
  const composed = text.restore(polishMarkdownArtifactsInDocx(saved, { headingSpacing }))
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
  // Numbering follows the lint, which reads the literal numbers as Markdown,
  // and the furniture, whose header and footer take the template's style ids too.
  // Cross-references follow the numbering.
  const furnished = furniture ? applyPageFurniture(composed, furniture) : composed
  const numbered = options.numbering === false ? furnished : numberDocxClauses(templated ? withTemplateStyleIds(furnished) : furnished)
  const operative = options.form !== undefined && !advisoryForm(options.form)
  const linked = options.numbering === false || !operative ? numbered : await linkCrossReferences(numbered, { form: options.form })
  const bytes = withGeneratedDocxAttribution(withParagraphPointers(highlightOpenItemsInDocx(linked), previous))
  return { bytes, lint, templated: Boolean(templated), headingSpacing }
}

// Whether the template's own Heading 1-3 paragraphs carry spacing of their own
// rather than their style's, as the house headings do.
function headingsSpaced(template: Uint8Array) {
  try {
    const body = unzipSync(template, { filter: file => file.name === "word/document.xml" })["word/document.xml"]
    return Boolean(body) && [...strFromU8(body!).matchAll(/<w:pPr>((?:(?!<\/w:pPr>)[\s\S])*)<\/w:pPr>/g)]
      .some(([, properties]) => /<w:pStyle\b[^>]*w:val="Heading[123]"[^>]*\/>/.test(properties!) && /<w:spacing\b/.test(properties!.replace(/<w:rPr>[\s\S]*?<\/w:rPr>/g, "")))
  } catch {
    return false
  }
}

// The firm's style template: a file in the template library, read when a
// document renders. A template since removed from the library is no style.
function firmStyleTemplate() {
  const name = normalizePreferences(readFirmPreferences()).styleTemplate
  const root = process.env.WORKSPACE_ROOT
  if (!name || !root) return
  const file = path.join(root, ".templates", path.basename(name))
  if (existsSync(file)) return readFileSync(file)
  warnOnce(`the firm's style template "${path.basename(name)}", which is not in the template library; the house style applies`)
}

// A style template that cannot be used is reported once per process, not on every render.
const warned = new Set<string>()
function warnOnce(detail: string) {
  if (warned.has(detail)) return
  warned.add(detail)
  console.warn(`[render] Rendering with ${detail}.`)
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
  // Who made the write, kept on its version (#268): the agent and its session.
  provenance?: { actor: string; run: string },
) {
  const matter = JSON.parse(readFileSync(path.join(directory, "matter.json"), "utf8")) as { id: string }
  const form = new FormData()
  form.append("file", new File([new Uint8Array(bytes)], name))
  const headers = precondition || provenance ? new Headers() : undefined
  if (provenance) {
    headers!.set("x-dochaus-actor", provenance.actor)
    headers!.set("x-dochaus-run", provenance.run)
  }
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
