import { useId, useLayoutEffect, useRef, type ReactNode } from "react"
import { createPortal } from "react-dom"
import backdropClose from "../backdropClose"
import useModalFocus from "../useModalFocus"

// Settings and focused setup stages share one native modal frame and dismissal
// policy. Children provide their own content and optional save footer.
export default function SettingsDialog({ title, children, onClose, closing = false, busy = false, locked = false, focused = false }: {
  title: string
  children: ReactNode
  onClose: () => void
  closing?: boolean
  busy?: boolean
  locked?: boolean
  focused?: boolean
}) {
  const dialog = useRef<HTMLDialogElement>(null)
  const titleId = useId()
  // Release the native modal and return focus as the exit starts; the closing
  // picture can remain mounted for its existing transition.
  useLayoutEffect(() => {
    if (closing) return
    const opener = document.activeElement instanceof HTMLElement ? document.activeElement : null
    const modal = dialog.current
    modal?.showModal()
    const initial = modal?.querySelector<HTMLElement>(".modal-close") ?? modal?.querySelector<HTMLElement>(".viewer-title")
    initial?.focus({ preventScroll: true })
    return () => {
      modal?.close()
      if (opener?.isConnected) opener.focus({ preventScroll: true })
    }
  }, [closing])
  useModalFocus(dialog, focused && !closing)
  const close = () => { if (!busy && !locked) onClose() }
  return createPortal(
    <dialog
      ref={dialog}
      className={`viewer-overlay settings-overlay${focused ? " settings-focused-overlay" : ""}${closing ? " is-closing" : ""}`}
      inert={closing}
      aria-labelledby={titleId}
      onCancel={(event) => { event.preventDefault(); close() }}
      {...backdropClose(close)}
    >
      <div className={`picker-panel settings-panel${focused ? " settings-focused-panel" : ""}`} onClick={(event) => event.stopPropagation()}>
        <div className="viewer-bar">
          <span className="viewer-title" id={titleId} tabIndex={-1}>{title}</span>
          {!locked && <button type="button" className="modal-close" aria-label="Close" disabled={busy} onClick={close}>
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeLinecap="round" aria-hidden="true"><path d="M18 6 6 18M6 6l12 12" /></svg>
          </button>}
        </div>
        {children}
      </div>
    </dialog>,
    document.body,
  )
}
