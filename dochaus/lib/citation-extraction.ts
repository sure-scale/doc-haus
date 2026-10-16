import { readLeanState } from "../workflows/lean/state"
import { readFileSync } from "node:fs"
import path from "node:path"
import type { CitationCheckRow, CitationCheckStatus, CitationPinpoint, CitationSupport } from "./citation-check"
import { buildSupportIndex, resolveSupport, type SupportHaystack } from "./citation-support"

// Citation extraction: finds citation candidates in text (URLs, rules and
// statutes, cases, reporter citations), normalizes their keys, and turns each
// into a report row checked against the matter's citation support.

type CitationCandidate = {
  citation_text: string
  normalized_key: string
  kind: CitationCheckRow["kind"]
  start: number
  end: number
}

export function normalizeCitationKey(raw: string): string {
  const trimmed = trimCitation(raw)
  if (!trimmed) return ""
  const url = normalizeURL(trimmed)
  if (url) return url
  const normalized = trimmed
    .replace(/\u00a0/g, " ")
    .replace(/[""]/g, '"')
    .replace(/[']/g, "'")
    .replace(/\s*§+\s*/g, " § ")
    .replace(/\bv\.?\b/gi, " v ")
    .replace(/\s+([,.;:)])/g, "$1")
    .replace(/([(])\s+/g, "$1")
    .replace(/\s+/g, " ")
    .replace(/[),.;:]+$/g, "")
    .trim()
    .toLowerCase()
  return rebalanceParens(normalized)
}

/** Check text that does not live at the matter root (for example a working
 * draft under `.dochaus/`) against the same matter citation support. */
export function citationRowsForText(directory: string, artifact: string, text: string): CitationCheckRow[] {
  return citationRowsFromText(text, artifact, buildSupportIndex(directory))
}

/** Lean authority comes from uploaded sources or the current independently
 * reviewed research, not a historical fetch or a previous pipeline's ledger. */
export function citationRowsForLeanText(directory: string, artifact: string, text: string): CitationCheckRow[] {
  const support = buildSupportIndex(directory)
  if (!readLeanState(directory)?.stages.research) return citationRowsFromText(text, artifact, support)
  return citationRowsFromText(text, artifact, support.filter(row =>
    row.support.id.startsWith("lean-research:") || !["official-fetch", "authority-ledger"].includes(row.support.kind) && !row.support.research_row_id))
}

export function citationRowsForArtifact(
  directory: string,
  artifact: string,
  supportIndex: SupportHaystack[],
): CitationCheckRow[] {
  return citationRowsFromText(readFileSync(path.join(directory, artifact), "utf8"), artifact, supportIndex)
}

function citationRowsFromText(
  text: string,
  artifact: string,
  supportIndex: SupportHaystack[],
): CitationCheckRow[] {
  return extractCitations(text).map((citation) => {
    const position = lineColumn(text, citation.start)
    const support = resolveSupport(citation.normalized_key, supportIndex)
    let status: CitationCheckStatus
    let pinpointInfo: CitationPinpoint | undefined
    let allSupport = support
    if (support.some((item) => item.status === "verified")) {
      status = "verified"
    } else if (support.length) {
      status = "unresolved"
    } else {
      const pinpoint = extractPinpoint(citation.normalized_key, citation.kind)
      if (pinpoint) {
        const baseSupport = resolveSupport(pinpoint.base_key, supportIndex)
        if (baseSupport.length) {
          status = "unresolved"
          pinpointInfo = {
            base_key: pinpoint.base_key,
            suffix: pinpoint.suffix,
            reason: `parent authority "${pinpoint.base_key}" has source support but pinpoint "${pinpoint.suffix}" is not confirmed`,
          }
          allSupport = dedupSupport([...support, ...baseSupport])
        } else {
          status = "unsupported"
        }
      } else {
        status = "unsupported"
      }
    }
    return {
      citation_text: citation.citation_text,
      normalized_key: citation.normalized_key,
      kind: citation.kind,
      artifact: {
        file: artifact,
        line: position.line,
        column: position.column,
        character_start: citation.start,
        character_end: citation.end,
      },
      status,
      linked_source_ids: allSupport
        .filter((item) => item.kind !== "research-row")
        .map((item) => item.id)
        .sort((a, b) => a.localeCompare(b)),
      linked_research_ids: allSupport
        .map((item) => item.research_row_id)
        .filter((id): id is string => Boolean(id))
        .sort((a, b) => a.localeCompare(b)),
      support: allSupport,
      pinpoint: pinpointInfo,
    }
  })
}

function dedupSupport(items: CitationSupport[]): CitationSupport[] {
  const seen = new Map<string, CitationSupport>()
  for (const item of items) seen.set(item.id, item)
  return [...seen.values()].sort((a, b) => a.id.localeCompare(b.id))
}

function extractCitations(text: string): CitationCandidate[] {
  const candidates = [
    ...matchCitations(text, /https?:\/\/[^\s<>"')\]]+/g, "url"),
    ...matchCitations(
      text,
      /\b(?:section|sections|s\.|§+|article|art\.|rule|r\.)[ \t]+[A-Za-z0-9][A-Za-z0-9()./-]*(?:[ \t]*(?:-|to|and|,)[ \t]*[A-Za-z0-9][A-Za-z0-9()./-]*)?[ \t]+of[ \t]+(?:the[ \t]+)?[A-Z][A-Za-z0-9&.'-]*(?:[ \t]+[A-Z][A-Za-z0-9&.'-]*){0,8}[ \t]+(?:Act|Code|Rules?|Regulations?|Statutes?)\b/g,
      "rule-or-statute",
    ),
    ...matchCitations(
      text,
      /\b[A-Z][A-Za-z0-9&.'-]*(?:[ \t]+[A-Z][A-Za-z0-9&.'-]*){0,8}[ \t]+(?:Act|Code|Rules?|Regulations?|Statutes?)[ \t]*(?:§+|s\.|section|sections|article|art\.|rule|r\.)[ \t]*[A-Za-z0-9][A-Za-z0-9()./-]*(?:[ \t]*(?:-|to|and)[ \t]*[A-Za-z0-9][A-Za-z0-9()./-]*|[ \t]*,[ \t]*\d[A-Za-z0-9()./-]*)?/g,
      "rule-or-statute",
    ),
    ...matchCitations(
      text,
      /\b[A-Z][A-Za-z0-9&.'-]+(?:[ \t]+(?:of|for|and|the|[A-Z][A-Za-z0-9&.'-]+)){0,8}[ \t]+v\.?[ \t]+[A-Z][A-Za-z0-9&.'-]+(?:[ \t]+(?:of|for|and|the|[A-Z][A-Za-z0-9&.'-]+)){0,8}\b/g,
      "case",
    ),
    ...matchCitations(
      text,
      /\b[A-Z][A-Za-z0-9&.'-]+(?:[ \t]+[A-Z][A-Za-z0-9&.'-]+){0,8},[ \t]+\d{1,4}[ \t]+[A-Z][A-Za-z.]{1,12}[ \t]+\d{1,6}(?:,[ \t]*\d{1,6})?\b/g,
      "reporter",
    ),
  ].sort((a, b) => a.start - b.start || b.end - b.start - (a.end - a.start))

  const selected: CitationCandidate[] = []
  for (const candidate of candidates) {
    if (selected.some((item) => rangesOverlap(item, candidate))) continue
    selected.push(candidate)
  }
  return selected
}

function matchCitations(text: string, pattern: RegExp, kind: CitationCandidate["kind"]): CitationCandidate[] {
  return [...text.matchAll(pattern)]
    .map((match) => {
      const citation_text = trimCitation(match[0])
      const start = (match.index ?? 0) + match[0].indexOf(citation_text)
      return {
        citation_text,
        normalized_key: normalizeCitationKey(citation_text),
        kind,
        start,
        end: start + citation_text.length,
      }
    })
    .filter((candidate) => candidate.citation_text && candidate.normalized_key)
}

function rangesOverlap(a: CitationCandidate, b: CitationCandidate) {
  return a.start < b.end && b.start < a.end
}

function extractPinpoint(
  normalizedKey: string,
  kind: CitationCandidate["kind"],
): { base_key: string; suffix: string } | null {
  if (kind === "url") return null
  // Parenthetical subsection pinpoints: § 12(a)(3), § 4.2.1, Rule 5(b)(ii)
  const parenMatch = normalizedKey.match(/^(.*?§+\s*[\w.-]+)((?:\([\w.-]+\))+)$/)
  if (parenMatch?.[1] && parenMatch[2]) return { base_key: parenMatch[1], suffix: parenMatch[2] }
  // Decimal-depth pinpoints: § 12.3.4 (base = § 12.3)
  const decimalMatch = normalizedKey.match(/^(.*?§+\s*\d+\.\d+)(\.\d+.*)$/)
  if (decimalMatch?.[1] && decimalMatch[2]) return { base_key: decimalMatch[1], suffix: decimalMatch[2] }
  // "at" page pinpoints in reporter/case citations: smith v jones at 456
  const atMatch = normalizedKey.match(/^(.*?\d{1,4}\s+[a-z]{1,12}\s+\d{1,6})\s+at\s+(.*)$/)
  if (atMatch?.[1] && atMatch[2]) return { base_key: atMatch[1], suffix: `at ${atMatch[2]}` }
  return null
}

function trimCitation(raw: string) {
  const trimmed = raw.replace(/^[\s([{"']+/g, "").replace(/[\s),.;:}"']+$/g, "").trim()
  return rebalanceParens(trimmed)
}

function rebalanceParens(text: string): string {
  const opens = (text.match(/\(/g) ?? []).length
  const closes = (text.match(/\)/g) ?? []).length
  if (opens > closes) return text + ")".repeat(opens - closes)
  return text
}

function normalizeURL(raw: string): string {
  if (!/^https?:\/\//i.test(raw)) return ""
  try {
    const parsed = new URL(raw)
    parsed.hash = ""
    parsed.searchParams.sort()
    return parsed.toString().toLowerCase()
  } catch {
    return raw.trim().toLowerCase()
  }
}

function lineColumn(text: string, index: number) {
  const lines = text.slice(0, index).split(/\r?\n/)
  return { line: lines.length, column: (lines.at(-1) ?? "").length + 1 }
}
