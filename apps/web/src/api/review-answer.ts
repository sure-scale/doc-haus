import type { Citation } from "./opencode/parts"
import type { Part } from "./session-view"

export type ReviewRow = { name: string; docPath?: string }

// Whether a citation is to the row's own document. Another matter document can
// share a section label, and its passage never supports this row's answer.
export function citesRow(citation: Pick<Citation, "documentName" | "docPath">, row: ReviewRow) {
  return citation.documentName === row.name && (!row.docPath || citation.docPath === row.docPath || citation.docPath === row.name)
}

// A short source label: the cited heading itself rather than the whole path.
// Ingest paths read `Doc(file) › §7 Termination › §7.2 Notice`; an unnumbered
// heading stands alone, so it gets no §. Clause-sentence headings are cut at a
// word; the full path stays in the answer detail.
export function citationLabel(citation: Pick<Citation, "section" | "structuralPath">) {
  const crumbs = (citation.structuralPath ?? "").split("›").map(crumb => crumb.trim()).filter(Boolean)
  if (crumbs[0]?.startsWith("Doc(")) crumbs.shift()
  const heading = crumbs.at(-1)?.replace(/^§\s*/, "§")
  if (heading) return shortLabel(heading)
  const section = citation.section.replace(/^§\s*/, "").trim()
  return /^\d/.test(section) ? `§ ${section}` : shortLabel(section)
}

const LABEL_MAX = 48

function shortLabel(label: string) {
  if (label.length <= LABEL_MAX) return label
  const [, number = "", title = label] = label.match(/^(§\S+\s+)(.+)$/) ?? []
  const room = Math.max(LABEL_MAX - number.length, 16)
  const cut = title.slice(0, room + 1)
  const space = cut.lastIndexOf(" ")
  return `${number}${(space > room / 2 ? cut.slice(0, space) : cut.slice(0, room)).replace(/[\s,;:.(–-]+$/, "")}…`
}

// Keep the final answer separate from intermediate progress text. Retrieval
// citations may belong to earlier provider steps in the same extraction.
export function readReviewAnswer(messages: Array<{ parts: Part[] }>, row: ReviewRow) {
  const { text, source } = readAnswerText(messages)
  const citations = retrievedCitations(messages).filter(citation => citesRow(citation, row))
  const section = source?.replace(/^§\s*/, "")
  const matches = citations.filter(citation => section && (citation.section === section ||
    citation.structuralPath?.split("›").at(-1)?.trim().replace(/^§\s*/, "") === section))
  const identities = new Map(matches.map(citation => [JSON.stringify([
    citation.docPath, citation.section, citation.charStart, citation.charEnd,
  ]), citation]))
  return { text, citation: identities.size === 1 ? identities.values().next().value : undefined }
}

// The final answer text and its raw `Source:` label (as written, `§` kept).
export function readAnswerText(messages: Array<{ parts: Part[] }>) {
  const final = messages.findLast(message => message.parts.some(part => part.type === "text" && part.text.trim()))
  const raw = (final?.parts ?? []).filter(part => part.type === "text").map(part => part.text).join("").trim()
  const source = raw.match(/\n\s*Source:\s*(.+?)\s*$/i)
  return { text: source ? raw.slice(0, source.index).trim() : raw, source: source?.[1]?.trim() }
}

// Every passage the extraction's search-document calls returned.
export function retrievedCitations(messages: Array<{ parts: Part[] }>) {
  return messages.flatMap(message => message.parts)
    .filter((part): part is Extract<Part, { type: "tool" }> => part.type === "tool")
    .flatMap(part => part.tool === "search-document" && part.state.status === "completed"
      ? (part.state.metadata?.citations ?? []) as Citation[]
      : [])
}
