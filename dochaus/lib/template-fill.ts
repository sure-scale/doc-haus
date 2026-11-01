import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { docxodus } from "./docxodus"
import { withParagraphPointers } from "./drafting-amendment"
import { renumberHeadings, splitBlocks } from "./markdown"
import { lintRenderedDraftMarkdown } from "./render"
import { withGeneratedDocxAttribution } from "./docx-attribution"
import { numberDocxClauses } from "./docx-numbering"
import { linkCrossReferences } from "./docx-crossrefs"
import { withTemplateStyleIds } from "./template-seed"
import { shadeOpenItemsInDocx, shadeTextsInDocx } from "./open-items"
import { FILLED_VALUE_FILL } from "./marker-patterns"
import { OPTIONAL_FIELD, type RowFill } from "./template-batch"

// The deterministic template fill: the one path that turns library template
// bytes plus placeholder values and optional-component choices into a new
// draft. draft-document's template mode and the spreadsheet batch
// (template-batch.ts) both run it, so a batch document is exactly what the
// tool would draft from the same values. No model is involved.

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

export type TemplateFillOptions = {
  /** Exact placeholder text -> final value. */
  fills?: Record<string, string>
  /** Optional component names (the text after "[optional: ") to leave out. */
  omit?: string[]
  /** Optional component names kept with their marker still in the text, as an open item. */
  keepMarkers?: string[]
  replaces?: { clause: string; replacement: string }[]
  /** Markdown body inserted after the template's first paragraph, which is then dropped. */
  content?: string
}

/** The markdown body could not be inserted into the template. */
export class TemplateContentError extends Error {}

export async function fillTemplate(template: Uint8Array, options: TemplateFillOptions) {
  const dx = await docxodus()
  const session = dx.openDocxSession(template, {})

  if (options.content) {
    // The seed's single empty paragraph is the insertion anchor; the whole
    // markdown body goes in as one multi-block insert, then the seed is dropped.
    const seed = Object.keys(session.project().anchorIndex).find((id) => id.startsWith("p:"))!
    const inserted = session.insertParagraph(seed, "after", splitBlocks(options.content))
    if (!inserted.success) {
      session.close()
      throw new TemplateContentError(inserted.error?.message ?? JSON.stringify(inserted.error))
    }
    session.deleteBlock(seed)
  }

  // Omitted optional components go first: a component is its `[optional: name]`-marked
  // heading plus everything under it (deleteSection), or just the marked block
  // when the marker sits on a plain paragraph. The engine refuses to delete or
  // remove a bookmark a cross-reference still cites, so the component's own
  // cross-reference bookmarks are taken out of its XML first; the refresh below
  // marks each reference to them as an open item. A hyperlink's target stays,
  // and the refused omission is reported.
  const missingOmits: string[] = []
  const refusedOmits: string[] = []
  for (const clause of options.omit ?? []) {
    const anchor = session.findByText(`[optional: ${clause}]`, { ignoreWhitespace: true })
    if (!anchor) {
      missingOmits.push(clause)
      continue
    }
    const component = new Set(Object.keys(session.projectAnchor(anchor.id).anchorIndex))
    const linked = new Set(session.listHyperlinks().map((link) => link.target))
    for (const mark of session.listBookmarks()) {
      const targets = /^_Ref\d+$/.test(mark.name) && !linked.has(mark.name) ? session.findByBookmark(mark.name) : []
      if (!mark.range || !targets.length || !targets.every((target) => component.has(target.id))) continue
      for (const id of new Set([mark.range.startAnchorId, mark.range.endAnchorId])) {
        const xml = new DOMParser().parseFromString(session.raw.getXml(id), "application/xml")
        for (const node of Array.from(xml.getElementsByTagNameNS(W, "*")))
          if (["bookmarkStart", "bookmarkEnd"].includes(node.localName) && node.getAttributeNS(W, "id") === mark.bookmarkId) node.parentNode!.removeChild(node)
        session.raw.replaceXml(id, new XMLSerializer().serializeToString(xml))
      }
    }
    const removed = anchor.kind === "h" ? session.deleteSection(anchor.id) : session.deleteBlock(anchor.id)
    if (!removed.success) refusedOmits.push(`${clause} (${removed.error?.message ?? removed.error?.code})`)
  }

  const fills = options.fills ?? {}
  if (Object.keys(fills).length) session.fillPlaceholders((p) => fills[p.match.text] ?? null)

  // Component rewrites run after fills so the legacy `clause` anchor can include filled-in
  // text, and before the optional-marker strip so a rewrite may target a kept
  // optional component by its visible text. Same locate semantics as the redline
  // tool: match the block's flat text whitespace-tolerantly, then swap the
  // whole paragraph (insert the replacement after it, delete the original).
  const missedReplaces: string[] = []
  for (const r of options.replaces ?? []) {
    const anchor = session.findByText(r.clause, { ignoreWhitespace: true })
    if (!anchor) {
      missedReplaces.push(r.clause)
      continue
    }
    const inserted = session.insertParagraph(anchor.id, "after", splitBlocks(r.replacement))
    if (!inserted.success) {
      missedReplaces.push(r.clause)
      continue
    }
    session.deleteBlock(anchor.id)
  }

  // Kept optional components lose their marker so it never reaches the draft,
  // unless the caller keeps it as an open item. A marker that opens its
  // paragraph goes with the space after it; the strip below absorbs whitespace
  // only between words, so it would leave the text starting with a space.
  const keep = new Set(options.keepMarkers ?? [])
  const strip = (text: string) => text.startsWith("[optional:") && !keep.has(OPTIONAL_FIELD.exec(text)?.[1] ?? "")
  for (const { match } of session.findPlaceholders(dx.PlaceholderKinds.AlternativeClause)) {
    const space = /^\s+/.exec(match.contextAfter)?.[0]
    if (strip(match.text) && space && !match.contextBefore.trim())
      session.replaceTextRange(match.enclosingAnchor.id, match.text + space, "", { maxReplacements: 1 })
  }
  session.fillPlaceholders((p) => (strip(p.match.text) ? "" : null), {
    kinds: dx.PlaceholderKinds.AlternativeClause,
    coalesceWhitespaceAroundEmptyFill: true,
  })

  // An omitted component leaves a hole in the manually numbered headings ("8."
  // followed by "10."); close the sequence up.
  if ((options.omit?.length ?? 0) > missingOmits.length) renumberHeadings(session)

  // Consistency lint over the composed result (defined-term hygiene,
  // cross-reference integrity, the firm's date/number style). Advisory — the
  // draft still ships, and the problems go back to the caller. The projection's
  // anchor markers and synthetic "# Document" line are engine bookkeeping,
  // not document text, so they are stripped before linting.
  const lint = lintRenderedDraftMarkdown(
    session
      .project()
      .markdown.replace(/\{#[^}]+\} /g, "")
      .split("\n")
      .filter((line) => line !== "# Document")
      .join("\n"),
  )

  const remaining = session.findPlaceholders().map((p) => p.match.text)
  // The template's literal clause numbers become its own (or the house)
  // Word list numbering once the text is final. Headings the engine inserts
  // name built-in style ids, which a localized template keeps under its own.
  // A library template has no document form either, so its typed clause
  // references stay text; the cross-references it already has are
  // refreshed. A template reference to a part the draft no longer has (an
  // omitted component) becomes an open-item marker, reported to the caller.
  const removedReferences: string[] = []
  let saved: Uint8Array
  try {
    saved = session.save()
  } finally {
    session.close()
  }
  const bytes = withGeneratedDocxAttribution(withParagraphPointers(await linkCrossReferences(numberDocxClauses(withTemplateStyleIds(saved)),
    { refreshOnly: true, onRemovedTarget: (marker) => removedReferences.push(marker) })))

  return { bytes, remaining, missingOmits, refusedOmits, missedReplaces, lint, removedReferences }
}

/** The filled document for one row: the shared template fill, then every
 * placeholder still in it shaded so it stands out. With `markFilled` (the
 * preview only) each placeholder it fills is shaded in FILLED_VALUE_FILL first,
 * and the fill keeps its run's formatting, so the viewer can outline exactly
 * the values written in; a created document is never marked. */
export async function fillTemplateDocx(template: Uint8Array, fill: RowFill, { markFilled = false } = {}) {
  const names = (keys: string[]) => keys.flatMap((key) => {
    const name = OPTIONAL_FIELD.exec(key)?.[1]
    return name ? [name] : []
  })
  const omit = names(Object.entries(fill.include).filter(([, included]) => !included).map(([key]) => key))
  // An unclear "Include" answer keeps the component and its marker, so the
  // document shows what is still to decide.
  const source = markFilled ? shadeTextsInDocx(template, Object.keys(fill.fills), FILLED_VALUE_FILL) : template
  const filled = await fillTemplate(source, { fills: fill.fills, omit, keepMarkers: names(fill.missingKeys) })
  return { bytes: shadeOpenItemsInDocx(filled.bytes, undefined, undefined, filled.remaining), remaining: filled.remaining }
}
