import { useRef, type ReactNode } from "react"
import ModalHeader from "../ModalHeader"
import useModalFocus from "../useModalFocus"
import backdropClose from "../backdropClose"

/** A focused saved-data surface that allows source viewers and native forms above it. */
export default function ChronologySurface({
  title,
  onClose,
  children,
}: {
  title: string
  onClose: () => void
  children: ReactNode
}) {
  const panel = useRef<HTMLDivElement>(null)
  useModalFocus(panel)
  return (
    <div className="viewer-overlay chronology-surface-overlay" {...backdropClose(() => onClose())}>
      <div ref={panel} className="chronology-surface-panel" role="dialog" aria-modal="true" aria-label={title}>
        <ModalHeader title={title} onClose={onClose} />
        <div className="chronology-surface-body">{children}</div>
      </div>
    </div>
  )
}
