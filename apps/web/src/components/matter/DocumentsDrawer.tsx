import { useEffect, useLayoutEffect, useRef, useState } from "react"
import { createPortal } from "react-dom"
import type { Document } from "../../api/ingest/documents"
import type { LeanDraftingStatus } from "../../api/ingest/lean"
import DocumentUpload from "../DocumentUpload"
import { setDocumentsDrawer } from "./documents-drawer"
import useModalFocus from "../useModalFocus"

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
  onCompare,
  closing = false,
  restored = false,
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
  onCompare?: (name: string) => void
  // Set while the drawer plays its exit (see useExit).
  closing?: boolean
  // Open since the matter loaded: shown in place, without the slide-in.
  restored?: boolean
}) {
  const [slot, setSlot] = useState<HTMLElement | null>(null)
  const panel = useRef<HTMLElement>(null)
  const [compact, setCompact] = useState(() => window.matchMedia("(max-width: 700px)").matches)
  // The full-screen phone drawer is a modal. Wider drawers intentionally keep
  // the matter rail available, so focus remains free to move between panels.
  // It lets go as the exit starts.
  useModalFocus(panel, Boolean(slot) && compact && !closing)
  // Wider drawers are not modal, but closing one with focus inside (or lost to
  // the page) still hands focus back as the exit starts: to what opened it,
  // else to a Documents toggle.
  const opener = useRef<Element | null>(null)
  useLayoutEffect(() => {
    const drawer = panel.current
    if (!closing) {
      if (!drawer?.contains(document.activeElement)) opener.current = document.activeElement
      return
    }
    if (compact || !drawer) return
    const active = document.activeElement
    if (active && active !== document.body && !drawer.contains(active)) return
    const target = [opener.current, ...document.querySelectorAll('[aria-controls="documents-drawer"]')].find(
      (el): el is HTMLElement => el instanceof HTMLElement && el !== document.body && el.isConnected && !drawer.contains(el) && el.getClientRects().length > 0,
    )
    target?.focus()
  }, [closing, compact])

  useEffect(() => {
    const query = window.matchMedia("(max-width: 700px)")
    const update = () => setCompact(query.matches)
    query.addEventListener("change", update)
    return () => query.removeEventListener("change", update)
  }, [])

  useEffect(() => {
    setSlot(document.getElementById("matter-drawer-slot"))
  }, [])

  if (!slot) return null
  const close = () => setDocumentsDrawer(false)
  const motion = closing ? " is-closing" : restored ? " is-restored" : ""
  return createPortal(
    <>
      <div className={`docs-drawer-scrim${motion}`} onClick={close} aria-hidden />
      <aside
        ref={panel}
        id="documents-drawer"
        className={`docs-drawer${motion}`}
        inert={closing}
        role={compact ? "dialog" : undefined}
        aria-modal={compact || undefined}
        aria-label="Documents"
        tabIndex={compact ? -1 : undefined}
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
          onCompare={onCompare}
        />
      </aside>
    </>,
    slot,
  )
}
