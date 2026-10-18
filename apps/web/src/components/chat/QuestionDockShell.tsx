import { type ReactNode, useEffect, useId, useLayoutEffect, useRef, useState } from "react"

export default function QuestionDockShell({
  count,
  kept = 0,
  keptTitle = "",
  preview,
  subject,
  renderQuestions,
}: {
  /** Questions asked; the dock counts them. */
  count: number
  /** Settled questions the dialog still offers; with none asked, the dock
   * names them with keptTitle and offers Review. */
  kept?: number
  keptTitle?: string
  preview: string
  subject: string
  renderQuestions: (titleId: string, onClose: () => void) => ReactNode
}) {
  const [open, setOpen] = useState(false)
  const dialog = useRef<HTMLDialogElement>(null)
  const titleId = useId()

  // Opened before the questions' own effects run, so their answer field can take focus.
  useLayoutEffect(() => {
    const node = dialog.current
    if (!node) return
    if (open && !node.open) node.showModal()
    if (!open && node.open) node.close()
  }, [open])

  // The dialog closes once the last asked question is settled, and once it holds nothing.
  useEffect(() => {
    if (count === 0) setOpen(false)
  }, [count])
  useEffect(() => {
    if (count + kept === 0) setOpen(false)
  }, [count, kept])

  if (count + kept === 0) return null
  return (
    <>
      <div className="question-dock">
        <span className="question-dock-icon" aria-hidden>
          <svg
            width="14"
            height="14"
            viewBox="0 0 24 24"
            fill="none"
            stroke="currentColor"
            strokeWidth="2.6"
            strokeLinecap="round"
            strokeLinejoin="round"
          >
            <path d="M9.1 9a3 3 0 0 1 5.8 1c0 2-3 3-3 3" />
            <line x1="12" y1="17" x2="12.01" y2="17" />
          </svg>
        </span>
        <div className="question-dock-text">
          <span className="question-dock-title">
            {count === 0 ? keptTitle : count === 1 ? `${subject} has a question for you` : `${subject} has ${count} questions for you`}
          </span>
          <span className="question-dock-preview">{preview}</span>
        </div>
        <button type="button" className="question-dock-cta" aria-haspopup="dialog" onClick={() => setOpen(true)}>
          {count === 0 ? "Review" : count === 1 ? "Answer" : `Answer ${count}`}
        </button>
      </div>
      <dialog
        ref={dialog}
        className="question-dialog"
        aria-labelledby={titleId}
        onClose={() => setOpen(false)}
        onClick={(e) => e.target === e.currentTarget && setOpen(false)}
      >
        {open && renderQuestions(titleId, () => setOpen(false))}
      </dialog>
    </>
  )
}
