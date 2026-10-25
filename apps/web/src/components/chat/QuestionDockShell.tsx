import { type ReactNode, useEffect, useId, useLayoutEffect, useRef, useState } from "react"
import backdropClose from "../backdropClose"
import useExit, { useClosed } from "../useExit"

export default function QuestionDockShell({
  count,
  preview,
  subject,
  answeredCount = 0,
  renderQuestions,
}: {
  /** Questions asked; the dock counts them. */
  count: number
  preview: string
  subject: string
  /** Saved answers remain reviewable after the last open question is settled. */
  answeredCount?: number
  renderQuestions: (titleId: string, onClose: () => void) => ReactNode
}) {
  const [open, setOpen] = useState(false)
  // The dialog stays mounted and closes at once; `is-closing` plays the exit
  // while its questions stay for it.
  const [shown] = useExit(open)
  const closed = useClosed(open)
  const dialog = useRef<HTMLDialogElement>(null)
  const titleId = useId()

  // Opened before the questions' own effects run, so their answer field can take focus.
  useLayoutEffect(() => {
    const node = dialog.current
    if (!node) return
    if (open && !node.open) node.showModal()
    if (!open && node.open) node.close()
  }, [open])

  // A newly asked question opens the dialog, so the lawyer need not find the
  // dock. While they type it waits there; asked while another dialog is open,
  // it opens once that dialog has gone, unless they are typing by then. One
  // they closed stays closed until something new is asked. The dialog closes
  // once the last asked question is settled.
  const asked = useRef(0)
  const [waiting, setWaiting] = useState(false)
  useEffect(() => {
    const more = count > asked.current
    asked.current = count
    if (count === 0) {
      setOpen(false)
      setWaiting(false)
    } else if (more && dialogOpen()) setWaiting(true)
    else if (more && !typing()) setOpen(true)
  }, [count])
  useEffect(() => {
    if (!waiting) return
    if (open) {
      setWaiting(false)
      return
    }
    // A closing dialog removes its open attribute or its own elements.
    const observer = new MutationObserver(() => {
      if (dialogOpen()) return
      setWaiting(false)
      if (!typing()) setOpen(true)
    })
    observer.observe(document.body, { subtree: true, childList: true, attributes: true, attributeFilter: ["open", "aria-modal"] })
    return () => observer.disconnect()
  }, [waiting, open])

  if (count === 0 && answeredCount === 0) return null
  const historyOnly = count === 0
  return (
    <>
      <div className={`question-dock${historyOnly ? " is-history" : ""}`}>
        <span className="question-dock-icon" aria-hidden>
          <svg
            width="14"
            height="14"
            viewBox="0 0 24 24"
            fill="none"
            stroke="currentColor"
            strokeWidth="2.5"
            strokeLinecap="round"
            strokeLinejoin="round"
          >
            <path d="M9.1 9a3 3 0 0 1 5.8 1c0 2-3 3-3 3" />
            <line x1="12" y1="17" x2="12.01" y2="17" />
          </svg>
        </span>
        <div className="question-dock-text">
          <span className="question-dock-title">
            {historyOnly ? "Your drafting answers" : count === 1 ? `${subject} has a question for you` : `${subject} has ${count} questions for you`}
          </span>
          <span className="question-dock-preview">{historyOnly ? `${answeredCount} saved ${answeredCount === 1 ? "answer" : "answers"}` : preview}</span>
        </div>
        <button type="button" className="question-dock-cta" aria-haspopup="dialog" onClick={() => setOpen(true)}>
          {historyOnly ? "Review answers" : count === 1 ? "Answer" : `Answer ${count}`}
        </button>
      </div>
      <dialog
        ref={dialog}
        className={`question-dialog${closed ? " is-closing" : ""}`}
        inert={closed}
        aria-labelledby={titleId}
        onClose={() => setOpen(false)}
        {...backdropClose(() => setOpen(false))}
      >
        {shown && renderQuestions(titleId, () => setOpen(false))}
      </dialog>
    </>
  )
}

// Whether the lawyer is typing: the focused text field holds their words. An
// empty one, such as the composer just sent from, does not count, and neither
// does a box or button, whose value is never empty.
const TEXT_ENTRY = new Set(["text", "search", "email", "url", "tel", "password", "number"])
function typing() {
  const active = document.activeElement
  const words =
    (active instanceof HTMLInputElement && TEXT_ENTRY.has(active.type)) || active instanceof HTMLTextAreaElement
      ? active.value
      : active instanceof HTMLElement && active.isContentEditable
        ? active.textContent
        : ""
  return Boolean(words?.trim())
}

// Whether another dialog holds the page.
function dialogOpen() {
  return document.querySelector("dialog[open], [aria-modal='true']") !== null
}
