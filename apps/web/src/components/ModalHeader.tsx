import { useEffect, type ReactNode } from "react"

// The one modal title bar: title on the left, optional actions, and an icon ×
// that also answers Escape, so every dialog closes the same way.
export default function ModalHeader({ title, actions, onClose }: { title: ReactNode; actions?: ReactNode; onClose: () => void }) {
  return (
    <div className="viewer-bar">
      <span className="viewer-title">{title}</span>
      <div className="viewer-bar-actions">
        {actions}
        <ModalClose onClose={onClose} />
      </div>
    </div>
  )
}

export function ModalClose({ onClose }: { onClose: () => void }) {
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape" && !e.defaultPrevented) onClose()
    }
    window.addEventListener("keydown", onKey)
    return () => window.removeEventListener("keydown", onKey)
  }, [onClose])

  return (
    <button type="button" className="modal-close" aria-label="Close" title="Close (Esc)" onClick={onClose}>
      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" aria-hidden="true">
        <path d="M18 6 6 18M6 6l12 12" />
      </svg>
    </button>
  )
}
