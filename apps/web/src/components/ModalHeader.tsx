import { useEffect, useId, useLayoutEffect, useRef, type ReactNode } from "react"
import useModalFocus from "./useModalFocus"

// The one modal title bar: title on the left, optional actions, and an icon ×
// that also answers Escape, so every dialog closes the same way.
export default function ModalHeader({ title, actions, onClose }: { title: ReactNode; actions?: ReactNode; onClose: () => void }) {
  const header = useRef<HTMLDivElement>(null)
  const panel = useRef<HTMLElement>(null)
  const titleId = useId()
  useLayoutEffect(() => {
    const element = header.current?.closest<HTMLElement>(".picker-panel")
    if (!element || element.closest("dialog")) return
    panel.current = element
    element.setAttribute("role", "dialog")
    element.setAttribute("aria-modal", "true")
    if (!element.hasAttribute("aria-label") && !element.hasAttribute("aria-labelledby")) element.setAttribute("aria-labelledby", titleId)
  }, [titleId])
  useModalFocus(panel)
  return (
    <div className="viewer-bar" ref={header}>
      <span className="viewer-title" id={titleId}>{title}</span>
      <div className="viewer-bar-actions">
        {actions}
        <ModalClose onClose={onClose} />
      </div>
    </div>
  )
}

export function ModalClose({ onClose }: { onClose: () => void }) {
  const control = useRef<HTMLButtonElement>(null)
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key !== "Escape" || e.defaultPrevented) return
      const selector = "dialog, [role='dialog'][aria-modal='true'], .picker-panel, .viewer-panel"
      const owner = control.current?.closest<HTMLElement>(selector)
      if (owner && owner.getClientRects().length === 0) return
      const target = e.target instanceof Element ? e.target : document.activeElement
      // A nested native dialog or another viewer owns this key, even when its
      // DOM is inside this panel. Only the active surface may close itself.
      if (owner && target?.closest(selector) !== owner) return
      e.preventDefault()
      onClose()
    }
    window.addEventListener("keydown", onKey)
    return () => window.removeEventListener("keydown", onKey)
  }, [onClose])

  return (
    <button ref={control} type="button" className="modal-close" aria-label="Close" title="Close (Esc)" onClick={onClose}>
      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" aria-hidden="true">
        <path d="M18 6 6 18M6 6l12 12" />
      </svg>
    </button>
  )
}
