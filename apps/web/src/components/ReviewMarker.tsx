import type { CellAttention, MarkerState } from "../api/review-attention"
import { Tooltip } from "./Tooltip"

// One marker per review cell: a shape, a token colour and a word. Shape and
// word carry the state; colour is a second cue, never the only one. The legend
// in Needs attention draws the same shapes.

type Shape = "cross" | "triangle" | "dot" | "ring" | "diamond"
export const MARKER_LOOK: Record<MarkerState, { shape: Shape; tone: "danger" | "warn" | "muted" | "info" }> = {
  other_document: { shape: "cross", tone: "danger" },
  source_missing: { shape: "triangle", tone: "danger" },
  source_changed: { shape: "triangle", tone: "warn" },
  question_changed: { shape: "dot", tone: "warn" },
  no_citation: { shape: "ring", tone: "muted" },
  differs: { shape: "diamond", tone: "info" },
}
const SHAPES: Record<Shape, React.ReactNode> = {
  cross: <path d="M2.2 2.2 7.8 7.8M7.8 2.2 2.2 7.8" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" />,
  triangle: <path d="M5 1 9.3 8.7H.7Z" fill="currentColor" />,
  dot: <circle cx="5" cy="5" r="3.6" fill="currentColor" />,
  ring: <circle cx="5" cy="5" r="3.3" fill="none" stroke="currentColor" strokeWidth="1.5" />,
  diamond: <path d="M5 .7 9.3 5 5 9.3.7 5Z" fill="currentColor" />,
}

export function MarkerGlyph({ state }: { state: MarkerState }) {
  const look = MARKER_LOOK[state]
  return <svg className={`review-glyph is-${look.tone}`} width="10" height="10" viewBox="0 0 10 10" aria-hidden="true">{SHAPES[look.shape]}</svg>
}

/**
 * The cell's highest-priority state, named by its word. The tooltip gives the
 * reason, the cell's other states and a re-check's outcome; the cell's detail
 * (`onOpen`) has the full account.
 */
export function ReviewMarker({ entry, recheck, onOpen }: { entry: CellAttention; recheck?: string; onOpen: () => void }) {
  const marker = entry.marker
  if (!marker) return null
  const also = entry.further.map((note) => note.word)
  const reason = [marker.reason, also.length ? `Also: ${also.join(", ")}.` : "", recheck ?? ""].filter(Boolean).join(" ")
  return (
    <Tooltip label={reason}>
      <button type="button" className={`review-marker is-${MARKER_LOOK[marker.state].tone}`} onClick={onOpen}>
        <MarkerGlyph state={marker.state} />
        {marker.word}
      </button>
    </Tooltip>
  )
}
