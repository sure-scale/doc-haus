import type { Template } from "../api/ingest/library"
import { markerKind } from "../components/document-markers"

// Shared by the two ways to start from a template: the Templates page (pick a
// matter) and the matter's Documents drawer (pick a template).

/** The matter chat with draft mode pinned and the assembly prompt prefilled;
 * the lawyer reviews and sends, nothing runs on its own. */
export function templateDraftLink(matterId: string, template: string) {
  const prompt = `Draft a new document from the template "${template}". Gather what you can from this matter's documents and our conversation first, then interview me for the remaining terms and any optional clauses before assembling it.`
  return `/matter/${matterId}?view=chat&agent=draft&prompt=${encodeURIComponent(prompt)}`
}

/** Fields to fill and optional clauses, counted as the template viewer counts
 * them: a [REDACTED] mark is a redaction, not a placeholder. */
export function templateFieldCounts(template: Template) {
  const markers = template.placeholders.filter((p) => markerKind(p.text) !== "redaction")
  const optional = markers.filter((p) => /^\[optional:/i.test(p.text)).length
  return { fills: markers.length - optional, optional, total: markers.length }
}

export function templateFieldSummary(template: Template) {
  const { fills, optional } = templateFieldCounts(template)
  return `${fills} placeholder${fills === 1 ? "" : "s"}${optional > 0 ? `, ${optional} optional clause${optional === 1 ? "" : "s"}` : ""}`
}
