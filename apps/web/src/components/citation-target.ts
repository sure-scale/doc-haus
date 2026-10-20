import type { Citation } from "../api/opencode/parts"
import { matterDocumentKind, type Document } from "../api/ingest/documents"

// Resolve against this matter's registry, never turn tool metadata into a URL or
// an arbitrary filesystem request. The viewer receives only a registered name.
export function citationDocument(citation: Citation, documents: Document[] = []) {
  const matches = documents.filter((document) =>
    document.name === citation.documentName &&
    !/[\\/\u0000]/.test(document.name) &&
    matterDocumentKind(document.name) !== undefined &&
    (citation.docPath === document.doc_path || citation.docPath === document.name),
  )
  return matches.length === 1 ? matches[0] : undefined
}

export type CitationFocus = Pick<Citation, "section" | "excerpt">

export function citationPassage(text: string, excerpt: string) {
  const normalize = (value: string) => value.replace(/\s+/g, " ").trim()
  const needle = normalize(excerpt)
  if (!needle) return { status: "unavailable" as const }
  // Keep a map to the rendered text's positions. Ingest character offsets refer
  // to a different projection and cannot locate text in the DOCX HTML.
  const positions: number[] = []
  const characters: string[] = []
  for (let index = 0; index < text.length; index++) {
    const character = /\s/.test(text[index]) ? " " : text[index]
    if (character === " " && (!characters.length || characters.at(-1) === " ")) continue
    characters.push(character)
    positions.push(index)
  }
  const haystack = characters.join("")
  const start = haystack.indexOf(needle)
  if (start < 0) return { status: "unavailable" as const }
  if (haystack.indexOf(needle, start + 1) >= 0) return { status: "ambiguous" as const }
  return { status: "matched" as const, start: positions[start], end: positions[start + needle.length - 1] + 1 }
}

export function focusCitationPassage(root: HTMLElement, excerpt: string) {
  const walker = root.ownerDocument.createTreeWalker(root, NodeFilter.SHOW_TEXT)
  const nodes: { node: Text; start: number; end: number }[] = []
  const text: string[] = []
  let length = 0
  let block: Element | null | undefined
  while (walker.nextNode()) {
    const node = walker.currentNode as Text
    const nextBlock = node.parentElement?.closest("p,li,td,th,h1,h2,h3,h4,h5,h6")
    if (nodes.length && block !== nextBlock) {
      text.push("\n")
      length++
    }
    nodes.push({ node, start: length, end: length + node.length })
    text.push(node.data)
    length += node.length
    block = nextBlock
  }
  const match = citationPassage(text.join(""), excerpt)
  if (match.status !== "matched") return match.status
  const marks = nodes.filter(({ start, end }) => start < match.end && end > match.start).map(({ node, start, end }) => {
    const range = root.ownerDocument.createRange()
    range.setStart(node, Math.max(0, match.start - start))
    range.setEnd(node, Math.min(end, match.end) - start)
    const mark = root.ownerDocument.createElement("mark")
    mark.className = "citation-passage"
    range.surroundContents(mark)
    return mark
  })
  marks[0].tabIndex = -1
  marks[0].scrollIntoView({ block: "center" })
  marks[0].focus({ preventScroll: true })
  return match.status
}
