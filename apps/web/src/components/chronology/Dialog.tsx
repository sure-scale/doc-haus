import { useLayoutEffect, useRef, type ReactNode } from "react"
import ModalHeader from "../ModalHeader"
import useModalFocus from "../useModalFocus"
export default function ChronologyDialog({
  title,
  onClose,
  children,
  footer,
}: {
  title: string
  onClose: () => void
  children: ReactNode
  footer?: ReactNode
}) {
  const dialog = useRef<HTMLDialogElement>(null)
  useModalFocus(dialog)
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
      <div>
        <ModalHeader title={title} onClose={onClose} />
        <div className="chronology-dialog-body">{children}</div>
        {footer && <div className="chronology-dialog-footer">{footer}</div>}
      </div>
    </dialog>
  )
}
