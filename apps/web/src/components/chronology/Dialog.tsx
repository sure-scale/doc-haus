import { useLayoutEffect, useRef, type ReactNode } from "react"
import ModalHeader from "../ModalHeader"
import useModalFocus from "../useModalFocus"
export default function ChronologyDialog({
  title,
  onClose,
  children,
}: {
  title: string
  onClose: () => void
  children: ReactNode
}) {
  const dialog = useRef<HTMLDialogElement>(null)
  const panel = useRef<HTMLDivElement>(null)
  useModalFocus(panel)
  useLayoutEffect(() => {
    dialog.current?.showModal()
    return () => dialog.current?.close()
  }, [])
  return (
    <dialog
      ref={dialog}
      className="chronology-dialog"
      aria-label={title}
      onCancel={(event) => {
        event.preventDefault()
        event.stopPropagation()
        onClose()
      }}
    >
      <div ref={panel}>
        <ModalHeader title={title} onClose={onClose} />
        <div className="chronology-dialog-body">{children}</div>
      </div>
    </dialog>
  )
}
