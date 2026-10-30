import { useEffect, useId, useRef, useState } from "react"
import ReviewSetupDialog from "./ReviewSetupDialog"
import "./ReviewStarter.css"

const plural = (n: number, word: string) => `${n} ${word}${n === 1 ? "" : "s"}`

// Adds a review question, or edits or removes one. Removing a column that holds
// answers asks first, inside the dialog, and says what goes with it.
export default function ReviewQuestion({ question, answers, reviewed, comments, disabled, onSave, onRemove, onClose }: {
  question?: string // absent when adding
  answers: number
  reviewed: number
  comments: number
  disabled: boolean
  onSave: (question: string) => void
  onRemove: () => void
  onClose: () => void
}) {
  const [text, setText] = useState(question ?? "")
  const [confirming, setConfirming] = useState(false)
  const ids = useId()
  const field = useRef<HTMLTextAreaElement>(null)
  const keep = useRef<HTMLButtonElement>(null)
  const removeButton = useRef<HTMLButtonElement>(null)
  const returning = useRef(false)
  const adding = question === undefined
  const next = text.replace(/\s+/g, " ").trim()
  // The dialog opens before effects run, so focus moves here rather than by autoFocus.
  useEffect(() => {
    if (confirming) { keep.current?.focus(); return }
    if (returning.current) { returning.current = false; removeButton.current?.focus(); return }
    const node = field.current
    node?.focus()
    node?.setSelectionRange(node.value.length, node.value.length)
  }, [confirming])
  function save() {
    if (!next || disabled) return
    if (next === question) onClose()
    else onSave(next)
  }
  function back() {
    returning.current = true
    setConfirming(false)
  }
  return (
    <ReviewSetupDialog title={confirming ? "Remove column" : adding ? "Add question" : "Edit question"} variant="narrow" onClose={onClose}>
      {confirming ? (
        <>
          <p id={`${ids}-confirm`}>
            Remove “{question}” and its {plural(answers, "answer")}{reviewed ? `, ${reviewed} of them reviewed` : ""}{comments ? `, and ${plural(comments, "comment")}` : ""}? This cannot be undone.
          </p>
          <div className="compare-actions" onKeyDown={(event) => {
            // Escape steps back to the question rather than closing the dialog.
            if (event.key !== "Escape") return
            event.preventDefault()
            event.stopPropagation()
            back()
          }}>
            <button type="button" ref={keep} className="ghost" onClick={back}>Keep column</button>
            <button type="button" className="danger" aria-describedby={`${ids}-confirm`} disabled={disabled} onClick={onRemove}>Remove column</button>
          </div>
        </>
      ) : (
        <>
          <label className="review-question-field" htmlFor={`${ids}-question`}>Question</label>
          <textarea
            id={`${ids}-question`}
            ref={field}
            className="review-question-input"
            rows={3}
            value={text}
            placeholder="e.g. What is the governing law?"
            aria-describedby={adding ? `${ids}-hint` : undefined}
            onChange={(event) => setText(event.target.value)}
            onKeyDown={(event) => {
              // A question is one line: Enter saves it.
              if (event.key !== "Enter" || event.shiftKey || event.nativeEvent.isComposing) return
              event.preventDefault()
              save()
            }}
          />
          {adding && <p id={`${ids}-hint`} className="review-question-hint">Asked of every document in the table.</p>}
          <div className="compare-actions">
            {!adding && <button type="button" ref={removeButton} className="ghost review-question-remove" disabled={disabled} onClick={() => (answers ? setConfirming(true) : onRemove())}>Remove column</button>}
            <button type="button" className="ghost" onClick={onClose}>Cancel</button>
            <button type="button" className="primary" disabled={!next || disabled} onClick={save}>{adding ? "Add" : "Save"}</button>
          </div>
        </>
      )}
    </ReviewSetupDialog>
  )
}
