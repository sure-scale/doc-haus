import { useLayoutEffect, useRef, type ReactNode } from "react"
import ModalHeader from "./ModalHeader"
import backdropClose from "./backdropClose"

// A native nested dialog keeps keyboard focus within the current review step.
export default function ReviewSetupDialog({ title, onClose, children }: { title: string; onClose: () => void; children: ReactNode }) {
  const dialog = useRef<HTMLDialogElement>(null)
  useLayoutEffect(() => {
    const opener = document.activeElement instanceof HTMLElement ? document.activeElement : null
    dialog.current?.showModal()
    return () => {
      dialog.current?.close()
      if (opener?.isConnected) opener.focus({ preventScroll: true })
    }
  }, [])
  return (
    <dialog ref={dialog} className="viewer-overlay review-detail-modal" aria-label={title} onCancel={(event) => { event.preventDefault(); event.stopPropagation(); onClose() }} {...backdropClose(onClose)} onKeyDown={(event) => {
      if (event.key !== "Tab") return
      const controls = [...event.currentTarget.querySelectorAll<HTMLElement>('button:not(:disabled), textarea:not(:disabled), input:not(:disabled), [tabindex="0"]')]
      const first = controls[0]
      const last = controls.at(-1)
      if (event.shiftKey && document.activeElement === first) { event.preventDefault(); last?.focus() }
      else if (!event.shiftKey && document.activeElement === last) { event.preventDefault(); first?.focus() }
    }}>
      <div className="picker-panel review-starter" onClick={(event) => event.stopPropagation()}>
        <ModalHeader title={title} onClose={onClose} />
        <div className="picker-body review-starter-body">{children}</div>
      </div>
    </dialog>
  )
}
