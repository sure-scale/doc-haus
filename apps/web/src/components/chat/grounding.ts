import type { Citation } from "../../api/opencode/parts"
import type { Part } from "../../api/session-view"

// The grounding state of a chat answer, read only from what the turn actually
// carries: the answer text, the citations the legal plugin verified (a
// quotation cite has no retrieval score), and the tools the turn ran.
//
// This is a structural check, never a semantic one. Model-authored inline
// citations bind a statement to a passage; code confirms that each marker names
// a quotation the plugin verified against the live document, and that every
// substantive block of the answer carries such a marker. A surviving quote is
// not proof that the passage supports the conclusion drawn from it — the badge
// says what was checked and never certifies legal support. Answers that make no
// claim (no marker, no verified quotation, and no matter evidence consulted) are
// chat and carry no badge.

export type GroundingState = "grounded" | "limited" | "abstain" | "chat"

// Answers that follow the document Q&A citation contract. A draft, redline,
// research or register answer has its own review surface; the badge would read
// as a claim check where none is being made.
export const GROUNDING_ANSWERS = new Set(["qa", "client-summary"])

export type Grounding = {
  state: GroundingState
  /** Substantive answer blocks bound to a verified quotation. */
  supported: number
  /** Substantive blocks with no bound marker, as short excerpts. */
  unlinked: string[]
  /** Distinct verified quotations the turn produced. */
  citations: number
  /** Markers naming no verified quotation. */
  unbound: number
  /** Quotations the plugin refused to verify in this turn. */
  rejected: number
  /** Quoted passages in the answer that match no verified quotation. */
  unverifiedQuotes: number
}

// Tools that read the matter's own evidence. Running one means the answer was
// working from sources, so unsupported prose is a grounding failure rather than
// ordinary chat. Browsing (glob, list) and research tools are deliberately
// absent: a research answer is judged by its source audit, not a matter span.
const EVIDENCE_TOOLS = new Set([
  "search-document",
  "read",
  "grep",
  "grep-matter",
  "get-section",
  "define",
  "matter-parties",
  "amendment-chain",
  "read-chronology",
  "read-document",
  "verify-quote",
  "cite",
  "compare-versions",
])

// The Q&A contract's inline citation: `[Document § 7.2]`, with "Section"
// accepted because agents also write it out. A marker may add a heading after
// the number (`[MSA § 7.2 – Limitation of Liability]`).
const MARKER = /\[\s*([^[\]\n]{1,200}?)\s*(?:§\s*|(?:section|sec\.)\s*)([^[\]\n]{1,120}?)\s*\]/giu
const BULLET = /^\s*(?:[-*+]|\d{1,3}[.)])\s+/
const TABLE_ROW = /^\s*\|/
const TABLE_RULE = /^\s*\|?[\s:|-]+\|?\s*$/
const HEADING = /^#{1,6}\s/
const RULE = /^([-*_]\s*){3,}$/
const QUOTED = /["“]([^"“”\n]{40,600})["”]/g
// A short block ("Here is a summary") carries no material claim on its own;
// this keeps framing lines from reading as unsupported assertions.
const MATERIAL = 25

type Marker = { document: string; section: string }

/** The grounding state for one assistant answer, from its text, verified
 * citations and the tools the turn ran. */
export function answerGrounding({ text, citations, parts }: { text: string; citations: Citation[]; parts: Part[] }): Grounding {
  // A quotation cite carries no score; a search hit is retrieval only and can
  // never stand as span-verified support for a claim.
  const quotes = distinct(citations.filter(citation => citation.score === undefined))
  const markers = [...text.matchAll(MARKER)].flatMap(match => marker(match))
  const verifiedExcerpts = quotes.map(quotation => textOf(quotation.excerpt))
  const unverifiedQuotes = [...text.matchAll(QUOTED)].filter(match => {
    const quoted = textOf(match[1] ?? "")
    return quoted.length >= 40 && !verifiedExcerpts.some(excerpt => excerpt.includes(quoted))
  }).length

  const blocks = unitsOf(text).filter(substantive)
  const supported = blocks.filter(block => markersIn(block).some(marker => bindings(marker, quotes))).length
  const unlinked = blocks
    .filter(block => !markersIn(block).some(marker => bindings(marker, quotes)))
    .map(block => excerptOf(block))
  const unbound = markers.filter(marker => !bindings(marker, quotes)).length
  const rejected = rejectedQuotes(parts)
  const consulted = consultEvidence(parts)
  const claims = markers.length > 0 || quotes.length > 0 || rejected > 0 || consulted

  const base = { supported, unlinked, citations: quotes.length, unbound, rejected, unverifiedQuotes }
  if (!claims) return { state: "chat", ...base }
  // Citations exist but no statement names one: only the mapping is missing, so
  // the answer cannot be called grounded even though the turn holds verified
  // quotations. Markers that name no verified quotation leave no support at all.
  if (supported === 0)
    return { state: markers.length === 0 && quotes.length > 0 ? "limited" : "abstain", ...base }
  if (unlinked.length || unbound || rejected || unverifiedQuotes) return { state: "limited", ...base }
  return { state: "grounded", ...base }
}

/** Whether a marker names a verified quotation. The document name is compared
 * with its extension and punctuation dropped; a section may carry a heading
 * after the number. A marker that names no document falls back to its section
 * only when one document's verified quotation has it, so an ambiguous name
 * never binds a claim to the wrong instrument. */
function bindings(marker: Marker, quotes: Citation[]) {
  const section = key(marker.section)
  if (!section) return false
  const name = documentKey(marker.document)
  const sameSection = quotes.filter(quotation => sectionMatches(section, key(quotation.section)))
  const named = sameSection.filter(quotation => documentKey(quotation.documentName) === name)
  if (named.length) return true
  const partial = name.length >= 3
    ? sameSection.filter(quotation => {
      const quoted = documentKey(quotation.documentName)
      return quoted.length >= 3 && (quoted.includes(name) || name.includes(quoted))
    })
    : []
  if (partial.length) return new Set(partial.map(quotation => quotation.docPath)).size === 1
  return name.length === 0 && new Set(sameSection.map(quotation => quotation.docPath)).size === 1
}

function sectionMatches(marker: string, cited: string) {
  if (!marker || !cited) return false
  if (marker === cited) return true
  const [long, short] = marker.length > cited.length ? [marker, cited] : [cited, marker]
  return long.startsWith(short) && !/[\p{L}\p{N}]/u.test(long.slice(short.length, short.length + 1))
}

function marker(match: RegExpMatchArray): Marker[] {
  const document = (match[1] ?? "").replace(/[\s,;:–—-]+$/u, "").trim()
  const section = (match[2] ?? "").trim()
  return section ? [{ document, section }] : []
}

function markersIn(text: string) {
  return [...text.matchAll(MARKER)].flatMap(match => marker(match))
}

function unitsOf(text: string) {
  const units: string[] = []
  for (const block of text.split(/\r?\n[ \t]*\r?\n/)) {
    const lines = block.split(/\r?\n/).filter(line => line.trim())
    let prose: string[] = []
    let fenced = false
    const flush = () => {
      if (prose.length) units.push(prose.join("\n"))
      prose = []
    }
    for (let index = 0; index < lines.length; index++) {
      const line = lines[index]!
      if (/^\s*```/.test(line)) {
        flush()
        fenced = !fenced
        continue
      }
      if (fenced) continue
      if (BULLET.test(line) || TABLE_ROW.test(line)) {
        flush()
        // A table's rule row carries no content; its header row is exempted by
        // the material-word threshold rather than guessed at here.
        if (!TABLE_RULE.test(line)) units.push(line)
      } else prose.push(line)
    }
    flush()
  }
  return units
}

function substantive(unit: string) {
  const plain = unit.replace(/^\s*(?:>+\s*)+/, "").trim()
  if (!plain || HEADING.test(plain) || RULE.test(plain) || /[:：]\s*$/.test(plain) || /\?\s*$/.test(plain)) return false
  const words = (plain.match(/[\p{L}\p{N}]/gu) ?? []).length
  // A short line that states a number, amount or date is material even when it
  // is too short to clear the prose threshold; framing lines rarely carry one.
  return words >= MATERIAL || /\d/.test(plain) || markersIn(plain).length > 0
}

function excerptOf(unit: string) {
  const plain = unit.replace(/^[\s>]*/, "").replace(/^\s*(?:[-*+]|\d{1,3}[.)])\s+/, "").replace(/\s+/g, " ").trim()
  return plain.length > 140 ? `${plain.slice(0, 139)}…` : plain
}

function distinct(citations: Citation[]) {
  const seen = new Set<string>()
  return citations.filter(citation => {
    const id = `${citation.docPath}§${citation.section}§${citation.excerpt}`
    if (seen.has(id)) return false
    seen.add(id)
    return true
  })
}

/** Whitespace-insensitive text for quotation containment. */
function textOf(value: string) {
  return value.replace(/\s+/g, " ").trim()
}

function key(value: string) {
  return value.toLowerCase().replace(/\s+/g, " ").replace(/[.,;:]+$/, "").trim()
}

function documentKey(value: string) {
  return key(value.replace(/\.(?:docx?|pdf|txt|md|eml|csv|xlsx?)$/i, "")).replace(/[^\p{L}\p{N}]+/gu, " ").trim()
}

function toolMetadata(part: Extract<Part, { type: "tool" }>) {
  return "metadata" in part.state ? part.state.metadata : undefined
}

/** Quotations the plugin refused to verify. Only `cite` is counted: a stale
 * search hit the model never quoted is retrieval noise, not a false claim. */
function rejectedQuotes(parts: Part[]) {
  return parts.filter((part): part is Extract<Part, { type: "tool" }> => part.type === "tool" && part.tool === "cite")
    .flatMap(part => {
      const rejected = toolMetadata(part)?.rejectedCitations
      return Array.isArray(rejected) ? rejected : []
    }).length
}

function consultEvidence(parts: Part[]) {
  return parts.some(part => part.type === "tool" && EVIDENCE_TOOLS.has(part.tool) && part.state.status !== "pending")
}
