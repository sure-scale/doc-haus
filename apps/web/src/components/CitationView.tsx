import type { Citation } from "../api/opencode/parts"
import type { Document } from "../api/ingest/documents"
import { useRef, useState } from "react"
import { createPortal } from "react-dom"
import DocumentViewer from "./DocumentViewer"
import { citationDocument } from "./citation-target"

export default function CitationView({ citations, matterId, documents }: { citations: Citation[]; matterId?: string; documents?: Document[] }) {
  const [viewing, setViewing] = useState<Citation>()
  const trigger = useRef<HTMLButtonElement>(null)
  if (citations.length === 0) return null
  // A turn's searches return the same section many times over; the reader needs
  // each distinct source once, not the raw retrieval dump. Keep the best-scoring
  // hit per document+section and tuck them behind a collapsed "Sources" line so
  // the answer stays the focus — the compact source list Harvey and Legora show.
  const best = new Map<string, Citation>()
  for (const c of [...citations].sort((a, b) => a.score - b.score)) best.set(`${c.docPath}§${c.section}`, c)
  const unique = [...best.values()].sort((a, b) => b.score - a.score)
  const document = viewing && citationDocument(viewing, documents)
  return (
    <>
    <details className="sources">
      <summary>
        Sources ({unique.length})
      </summary>
      {unique.map((c, i) => (
        <div className="citation" key={`${c.docPath}-${c.charStart}-${i}`}>
          {matterId && citationDocument(c, documents) ? (
            <button type="button" className="ref" aria-label={`Open ${c.documentName}, section ${c.section}`} onClick={(event) => {
              trigger.current = event.currentTarget
              setViewing(c)
            }}>[{c.documentName} § {c.section}]</button>
          ) : <div className="ref">[{c.documentName} § {c.section}]<span className="muted"> — Source unavailable in this matter</span></div>}
          <div className="excerpt">{c.excerpt}</div>
        </div>
      ))}
    </details>
    {viewing && document && matterId && typeof window !== "undefined" && createPortal(<DocumentViewer
      matterId={matterId}
      name={document.name}
      citation={viewing}
      onClose={() => {
        setViewing(undefined)
        trigger.current?.focus()
      }}
    />, window.document.body)}
    </>
  )
}
