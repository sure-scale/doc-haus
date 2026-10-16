import { useEffect, useState } from "react"
import { createPortal } from "react-dom"
import type { Document } from "../../api/ingest/documents"
import type { LeanDraftingStatus } from "../../api/ingest/lean"
import DocumentUpload from "../DocumentUpload"
import { setDocumentsDrawer } from "./documents-drawer"

// The matter's documents as a panel that slides out beside the rail. It is
// portaled into the shell's slot so it sits between the rail and the canvas:
// wide windows push the canvas over, narrower ones lay it over the canvas with
// a scrim, phones take the full screen. The document viewer opens above it.
export default function DocumentsDrawer({
  matterId,
  documents,
  drafting,
  incoming,
  onIncoming,
  onUploaded,
  onView,
}: {
  matterId: string
  documents: Document[]
  drafting?: LeanDraftingStatus
  // Files dropped on the composer, handed here so every upload shows its
  // progress in one list.
  incoming?: File[]
  onIncoming: () => void
  onUploaded: () => void
  onView: (name: string) => void
}) {
  const [slot, setSlot] = useState<HTMLElement | null>(null)

  useEffect(() => {
    setSlot(document.getElementById("matter-drawer-slot"))
  }, [])

  if (!slot) return null
  const close = () => setDocumentsDrawer(false)
  return createPortal(
    <>
      <div className="docs-drawer-scrim" onClick={close} aria-hidden />
      <aside
        id="documents-drawer"
        className="docs-drawer"
        aria-label="Documents"
        onKeyDown={(e) => {
          if (e.key !== "Escape" || e.defaultPrevented) return
          e.preventDefault()
          close()
        }}
      >
        <DocumentUpload
          matterId={matterId}
          documents={documents}
          onUploaded={onUploaded}
          onView={onView}
          drafting={drafting}
          onClose={close}
          incoming={incoming}
          onIncoming={onIncoming}
        />
      </aside>
    </>,
    slot,
  )
}
