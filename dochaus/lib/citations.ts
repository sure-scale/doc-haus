// Shared citation shape and rendering, used by the tools that produce citations
// (search-document's vector hits, cite's anchored quotations) and the legal
// plugin (which re-renders the surviving set after span verification). One
// formatter keeps every output identical — including the untrusted-data framing,
// so no path can hand the model a document excerpt without it.

import { FLAGGED_PASSAGE_NOTE, UNTRUSTED_EXCERPT_NOTICE } from "./untrusted"

const SUPERSEDED_PASSAGE_NOTE =
  "[⚠ SUPERSEDED] This passage may be from a section a later amendment restates, amends, or deletes. Use the `current-section` tool to fetch the operative text before quoting this passage as current."

const DEFAULT_EXCERPT_DISPLAY_CHARS = 1000

export type DocumentCitation = {
  documentName: string
  docPath: string
  section: string
  // Full clause hierarchy ("Doc › §7 Indemnification › §7.2 Cap") and the
  // document's parties, attached by ingest's structure-aware chunking. Both are
  // optional — citations from sources without them (older indexes, the cite
  // tool) simply omit the line.
  structuralPath?: string
  party?: string
  excerpt: string
  charStart: number
  charEnd: number
  // Vector hits from search-document carry a similarity score; quotations the
  // cite tool anchored do not, so score is optional.
  score?: number
  reason?: string
  confidence?: number
  context?: string
  verified?: boolean
  reanchored?: boolean
  // Set when ingest flagged the chunk as containing instruction-like text
  // (see services/ingest/src/sanitize.ts), so the rendering can warn the model
  // before it reads the excerpt.
  flagged?: boolean
  // Set when this chunk belongs to a section of an amended base that a later
  // amender restates, amends, or deletes (issue #69). The renderer warns the
  // model above the excerpt; retrieval also down-ranks the chunk so the
  // amender's operative version outranks it.
  superseded_by?: {
    amenderName: string
    operation: "amend" | "restate" | "delete" | "add"
    charStart: number
    charEnd: number
  }
}

export function formatCitations(citations: DocumentCitation[]) {
  if (!citations.length) return ""
  const body = citations
    .map((c, i) => {
      const score = c.score === undefined ? "" : ` (score ${c.score.toFixed(3)})`
      const flag = c.flagged ? `\n${FLAGGED_PASSAGE_NOTE}` : ""
      const superseded = c.superseded_by
        ? `\n${SUPERSEDED_PASSAGE_NOTE}\n   superseded by: ${c.superseded_by.amenderName} (${c.superseded_by.operation}, chars ${c.superseded_by.charStart}..${c.superseded_by.charEnd})`
        : ""
      const reason =
        c.reason === undefined
          ? ""
          : `\n   reason: ${c.reason}${c.confidence === undefined ? "" : ` (confidence ${c.confidence}/5)`}`
      // Prefer the full structural path over the bare section label when ingest
      // supplied one; append the parties so the passage carries its matter context.
      const heading = c.structuralPath || `${c.documentName} § ${c.section}`
      const party = c.party ? `\n   parties: ${c.party}` : ""
      return `${i + 1}. [${heading}]${score}${flag}${superseded}${party}\n${displayExcerpt(c.excerpt)}${reason}`
    })
    .join("\n\n")
  return `${UNTRUSTED_EXCERPT_NOTICE}\n\n${body}`
}

function displayExcerpt(excerpt: string) {
  const max = citationExcerptDisplayChars()
  if (excerpt.length <= max) return excerpt
  return (
    excerpt.slice(0, max).trimEnd() +
    `\n[passage preview clipped at ${max} characters for context budget; use cite or read-document for the full verified passage before quoting beyond this preview.]`
  )
}

function citationExcerptDisplayChars() {
  const value = Number(process.env.DOCHAUS_CITATION_EXCERPT_CHARS ?? DEFAULT_EXCERPT_DISPLAY_CHARS)
  if (!Number.isFinite(value) || value < 400) return DEFAULT_EXCERPT_DISPLAY_CHARS
  return Math.floor(value)
}
