import { useAnswerDrafts } from "../../use-answer-drafts"
import AnswerDraftStatus from "../AnswerDraftStatus"
import { useEffect, useId, useRef, useState } from "react"
import type { LeanQuestion } from "../../api/ingest/lean"
import { ModalClose } from "../ModalHeader"
import { delegableQuestion } from "./model"

// The run's open questions, one at a time with an "N of M" counter and how many
// are answered, inside the question dialog. Answers in progress are kept per
// run on the server. Once saved, answers survive dialog closure and reload.
// Each answer is the lawyer's own: there is no
// default or suggested answer. A question the run's model asked can instead be
// handed back with Decide for me; the count calls it left to the drafter. The
// cursor sits in the answer field of whichever question is showing.
export default function Questions({
  matterId,
  ownerSessionID,
  questions,
  answered,
  handedBack = 0,
  titles,
  busy,
  error,
  titleId,
  onClose,
  onSend,
  onDecide,
}: {
  matterId: string
  ownerSessionID: string
  questions: LeanQuestion[]
  answered: number
  /** How many of the answered questions the lawyer handed back. */
  handedBack?: number
  titles: Map<string, string>
  busy: boolean
  error: string
  /** Id for the heading, which names the dialog. */
  titleId: string
  onClose: () => void
  onSend: (question: LeanQuestion, answer: string) => Promise<boolean>
  onDecide: (question: LeanQuestion) => Promise<boolean>
}) {
  const hintId = useId()
  const drafts = useAnswerDrafts(matterId, `lean.${ownerSessionID}`, `dh.lean.answers.${matterId}`)
  const [index, setIndex] = useState(0)
  const field = useRef<HTMLTextAreaElement>(null)
  const at = Math.min(index, Math.max(0, questions.length - 1))
  const shown = questions[at]?.id

  // A new question (the next one once an answer is sent) gets the cursor.
  useEffect(() => {
    field.current?.focus({ preventScroll: true })
  }, [shown, drafts.loaded])

  if (questions.length === 0) return null
  const question = questions[at]!
  const draft = typeof drafts.answers[question.id] === "string" ? drafts.answers[question.id] as string : ""
  const about = question.document && titles.size > 1 ? titles.get(question.document) : undefined
  const fieldId = `lean-answer-${question.id}`
  const total = answered + questions.length
  const delegated = Math.min(handedBack, answered)
  const delegable = delegableQuestion(question)

  async function send() {
    if (!(await onSend(question, draft))) return
    drafts.set(question.id, undefined)
  }

  async function decide() {
    if (!(await onDecide(question))) return
    drafts.set(question.id, undefined)
  }

  return (
    <section className="question-card">
      <div className="question-head">
        <h2 id={titleId}>{questions.length === 1 ? "Question for you" : `Question ${at + 1} of ${questions.length}`}</h2>
        {questions.length > 1 && (
          <span className="question-nav">
            <button type="button" className="icon-btn question-nav-btn" disabled={at === 0} onClick={() => setIndex(at - 1)} aria-label="Previous question">
              <Chevron points="15 18 9 12 15 6" />
            </button>
            <button
              type="button"
              className="icon-btn question-nav-btn"
              disabled={at === questions.length - 1}
              onClick={() => setIndex(at + 1)}
              aria-label="Next question"
            >
              <Chevron points="9 18 15 12 9 6" />
            </button>
          </span>
        )}
        <ModalClose onClose={onClose} />
      </div>
      {total > 1 && (
        <div className="question-progress">
          <span className="question-dots" aria-hidden>
            {Array.from({ length: total }, (_, i) => (
              <span
                key={i}
                className={`question-dot${i < answered ? " is-answered" : ""}${i >= answered - delegated && i < answered ? " is-delegated" : ""}${i === answered + at ? " is-current" : ""}`}
              />
            ))}
          </span>
          <span>
            {delegated
              ? `${answered - delegated} answered, ${delegated} left to the drafter, ${questions.length} to answer`
              : `${answered} answered, ${questions.length} left`}
          </span>
        </div>
      )}
      {about && <p className="question-about">About {about}</p>}
      <p className="question-text">{question.question}</p>
      {question.context && <p className="question-context">{question.context}</p>}
      <label className="question-label" htmlFor={fieldId}>
        Your answer
      </label>
      <textarea
        ref={field}
        id={fieldId}
        className="question-answer"
        value={draft}
        disabled={busy || !drafts.loaded}
        onChange={(e) => drafts.set(question.id, e.target.value)}
        onKeyDown={(e) => {
          if (e.key === "Enter" && (e.metaKey || e.ctrlKey) && !busy && drafts.loaded && draft.trim()) send()
        }}
        rows={4}
        placeholder="In your own words."
      />
      <AnswerDraftStatus {...drafts} />
      {error && (
        <div className="msg-error" role="alert">
          {error}
        </div>
      )}
      <div className="question-actions">
        <span className="question-keys" aria-hidden>
          ⌘/Ctrl+Enter to send
        </span>
        {delegable && (
          <button type="button" aria-describedby={hintId} disabled={busy} onClick={decide}>
            Decide for me
          </button>
        )}
        <button className="primary" type="button" disabled={busy || !drafts.loaded || !draft.trim()} onClick={send}>
          {busy ? "Sending…" : "Send answer"}
        </button>
      </div>
      {delegable && (
        <p id={hintId} className="question-delegate-hint">
          You'll get a review draft that says what was chosen and why.
        </p>
      )}
    </section>
  )
}

function Chevron({ points }: { points: string }) {
  return (
    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.25" strokeLinecap="round" strokeLinejoin="round" aria-hidden>
      <polyline points={points} />
    </svg>
  )
}
