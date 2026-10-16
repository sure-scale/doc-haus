import { existsSync } from "node:fs"
import { findQuote, liveText } from "../../lib/extract"
import type { DocumentCitation } from "../../lib/citations"

// Run the exact → re-anchor → reject ladder for one citation. Returns the
// verified (possibly re-anchored) citation, or undefined to reject it.
export async function verifyCitation(citation: DocumentCitation) {
  if (!existsSync(citation.docPath)) return undefined
  const text = await liveText(citation.docPath)
  if (text.slice(citation.charStart, citation.charEnd) === citation.excerpt) return { ...citation, verified: true }

  // The indexed span does not match live extraction. This can reflect an edit
  // or an indexing/extraction alignment difference, not proof of an amendment.
  // Search for the excerpt, preferring the occurrence nearest its indexed span.
  const match = findQuote(text, citation.excerpt, citation.charStart)
  if (!match) return undefined
  return {
    ...citation,
    charStart: match.start,
    charEnd: match.end,
    excerpt: match.excerpt,
    verified: true,
    reanchored: true,
  }
}
