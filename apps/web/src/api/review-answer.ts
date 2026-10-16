import type { Citation } from "./opencode/parts"
import type { Part } from "./session-view"

// Keep the final answer separate from intermediate progress text. Retrieval
// citations may belong to earlier provider steps in the same extraction.
export function readReviewAnswer(messages: Array<{ parts: Part[] }>) {
  const final = messages.findLast(message => message.parts.some(part => part.type === "text" && part.text.trim()))
  const raw = (final?.parts ?? []).filter(part => part.type === "text").map(part => part.text).join("").trim()
  const citations = messages.flatMap(message => message.parts)
    .filter((part): part is Extract<Part, { type: "tool" }> => part.type === "tool")
    .flatMap(part => part.tool === "search-document" && part.state.status === "completed"
      ? (part.state.metadata?.citations ?? []) as Citation[]
      : [])
  const source = raw.match(/\n\s*Source:\s*§?\s*(.+?)\s*$/i)
  const section = source?.[1]?.trim()
  const matches = citations.filter(citation => section && (citation.section === section ||
    citation.structuralPath?.split("›").at(-1)?.trim().replace(/^§\s*/, "") === section))
  const identities = new Map(matches.map(citation => [JSON.stringify([
    citation.docPath, citation.section, citation.charStart, citation.charEnd,
  ]), citation]))
  return {
    text: source ? raw.slice(0, source.index).trim() : raw,
    citation: identities.size === 1 ? identities.values().next().value : undefined,
  }
}
