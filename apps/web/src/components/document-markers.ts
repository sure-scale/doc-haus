import {
  INSTRUCTION_COMMENT_PATTERN,
  OPEN_ITEM_PATTERN,
  OPTIONAL_CLAUSE_PATTERN,
  PLACEHOLDER_PATTERN,
  REDACTION_MARKER_PATTERN,
} from "../../../../dochaus/lib/marker-patterns"

// Display-only colour coding of the product's own markers in a rendered
// document. The patterns are the server's (dochaus/lib/marker-patterns.ts), so
// the viewer shows exactly what render-readiness and the DOCX highlights see.
// Each match is wrapped in place with a <span>; no text is added, removed or
// reordered, so the text offsets focusCitationPassage walks are unchanged.

export type MarkerKind = "redaction" | "placeholder"
export type MarkerCounts = Record<MarkerKind, number>

const CLASSES: { kind: MarkerKind; patterns: RegExp[] }[] = [
  { kind: "redaction", patterns: [REDACTION_MARKER_PATTERN] },
  { kind: "placeholder", patterns: [PLACEHOLDER_PATTERN, INSTRUCTION_COMMENT_PATTERN, OPEN_ITEM_PATTERN, OPTIONAL_CLAUSE_PATTERN] },
]
const LABELS: Record<MarkerKind, string> = { redaction: "Redacted text", placeholder: "Placeholder: to complete" }
const BLOCKS = "p,li,td,th,h1,h2,h3,h4,h5,h6"

type Piece = { node: Text; start: number; end: number }

export function decorateMarkers(root: HTMLElement): MarkerCounts {
  for (const span of root.querySelectorAll("span.doc-marker")) span.replaceWith(...span.childNodes)
  root.normalize()
  let sequence = 0
  const counts: MarkerCounts = { redaction: 0, placeholder: 0 }
  const groups = new Map<Element, Piece[]>()
  const walker = root.ownerDocument.createTreeWalker(root, NodeFilter.SHOW_TEXT)
  while (walker.nextNode()) {
    const node = walker.currentNode as Text
    const block = node.parentElement?.closest(BLOCKS) ?? root
    const pieces = groups.get(block) ?? []
    const start = pieces.at(-1)?.end ?? 0
    pieces.push({ node, start, end: start + node.length })
    groups.set(block, pieces)
  }
  for (const pieces of groups.values()) {
    const text = pieces.map((piece) => piece.node.data).join("")
    const found: { kind: MarkerKind; start: number; end: number }[] = []
    for (const { kind, patterns } of CLASSES)
      for (const pattern of patterns)
        for (const match of text.matchAll(pattern))
          found.push({ kind, start: match.index, end: match.index + match[0].length })
    // Longest first at a shared start; a match inside an earlier one is dropped.
    found.sort((a, b) => a.start - b.start || b.end - a.end)
    const kept: typeof found = []
    for (const marker of found) if (!kept.length || marker.start >= kept[kept.length - 1].end) kept.push(marker)
    // Last match first: splitting a text node leaves earlier offsets valid.
    for (const { kind, start, end } of kept.reverse()) {
      counts[kind]++
      const id = String(++sequence)
      const wrapped: HTMLSpanElement[] = []
      for (const piece of pieces) {
        if (piece.start >= end || piece.end <= start) continue
        const range = root.ownerDocument.createRange()
        range.setStart(piece.node, Math.max(0, start - piece.start))
        range.setEnd(piece.node, Math.min(piece.end, end) - piece.start)
        const span = root.ownerDocument.createElement("span")
        span.className = `doc-marker doc-marker-${kind}`
        range.surroundContents(span)
        wrapped.push(span)
      }
      // One marker across several runs is one chip: shared id, rounded outer
      // ends only, and one accessible name on the first piece.
      wrapped.forEach((span, index) => {
        span.dataset.markerId = id
        span.dataset.markerPart = wrapped.length === 1 ? "only" : index === 0 ? "first" : index === wrapped.length - 1 ? "last" : "middle"
        if (index > 0) return span.setAttribute("aria-hidden", "true")
        span.title = LABELS[kind]
        span.setAttribute("role", "mark")
        span.setAttribute("aria-roledescription", LABELS[kind])
        if (wrapped.length > 1) span.setAttribute("aria-label", text.slice(start, end).replace(/\u00a0/g, " "))
      })
    }
  }
  return counts
}

export function markerSummary({ placeholder, redaction }: MarkerCounts) {
  return [
    placeholder && `${placeholder} placeholder${placeholder === 1 ? "" : "s"}`,
    redaction && `${redaction} redaction${redaction === 1 ? "" : "s"}`,
  ].filter(Boolean).join(" · ")
}
