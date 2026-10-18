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
// handed back with Decide for me while its answer box is empty; the count calls
// it left to the drafter, and it stays in its place so the lawyer can still
// answer it instead. The questions keep the run's order, so a hand-back moves
// no question and each keeps its number and dot. The cursor sits in the answer
// field of whichever question is showing.
export default function Questions({
  matterId,
  ownerSessionID,
  questions,
  answered,
  handedBack = [],
  order,
  done = false,
  titles,
  busy,
  error,
  errorFor,
  titleId,
  onClose,
  onSend,
  onDecide,
}: {
  matterId: string
  ownerSessionID: string
  questions: LeanQuestion[]
  answered: number
  /** The answered questions left to the drafter, offered in their place. */
  handedBack?: LeanQuestion[]
  /** The run's question ids in the order it asked them. Without it, the open
   * questions come first. */
  order?: string[]
  /** The run finished: answering a question left to the drafter runs the
   * drafting again. */
  done?: boolean
  titles: Map<string, string>
  busy: boolean
  error: string
  /** The question the error is about; it shows only with that question. */
  errorFor?: string
  /** Id for the heading, which names the dialog. */
  titleId: string
  onClose: () => void
  onSend: (question: LeanQuestion, answer: string) => Promise<boolean>
  onDecide: (question: LeanQuestion) => Promise<boolean>
}) {
  const hintId = useId()
  const rerunId = useId()
  const drafts = useAnswerDrafts(matterId, `lean.${ownerSessionID}`, `dh.lean.answers.${matterId}`)
  // The question chosen to show; until one is (or once it leaves the card),
  // the first one still asked.
  const [chosen, setChosen] = useState<string>()
  // Which button's request is in flight, so each reads as its own.
  const [pending, setPending] = useState<"send" | "decide">()
  // Counts refused hand-backs: the cursor goes back to Decide for me.
  const [refused, setRefused] = useState(0)
  const field = useRef<HTMLTextAreaElement>(null)
  const decideButton = useRef<HTMLButtonElement>(null)
  const place = new Map(order?.map((id, index) => [id, index]))
  const cards = [...questions, ...handedBack].sort((a, b) => (place.get(a.id) ?? place.size) - (place.get(b.id) ?? place.size))
  const handed = new Set(handedBack.map((question) => question.id))
  const found = cards.findIndex((question) => question.id === chosen)
  const at = found === -1 ? Math.max(0, cards.findIndex((question) => !handed.has(question.id))) : found
  const shown = cards[at]?.id

  // A new question (the next one once an answer is sent) gets the cursor.
  useEffect(() => {
    field.current?.focus({ preventScroll: true })
  }, [shown, drafts.loaded])

  useEffect(() => {
    if (refused) decideButton.current?.focus()
  }, [refused])

  if (cards.length === 0) return null
  const question = cards[at]!
  const left = handed.has(question.id)
  const draft = typeof drafts.answers[question.id] === "string" ? drafts.answers[question.id] as string : ""
  const about = question.document && titles.size > 1 ? titles.get(question.document) : undefined
  const fieldId = `lean-answer-${question.id}`
  const delegated = Math.min(handedBack.length, answered)
  // Answered dots first, then one per question in the card, in its order.
  const settled = answered - delegated
  const total = settled + cards.length
  const current = settled + at
  const dot = (i: number) => {
    const handedDot = i >= settled && handed.has(cards[i - settled]!.id)
    return `question-dot${i < settled || handedDot ? " is-answered" : ""}${handedDot ? " is-delegated" : ""}${i === current ? " is-current" : ""}`
  }
  const delegable = !left && delegableQuestion(question)
  const rerun = left && done

  // Once this question is answered or handed back, the next one still asked.
  function moveOn() {
    const asked = cards.filter((card) => card.id !== question.id && !handed.has(card.id))
    setChosen((asked.find((card) => cards.indexOf(card) > at) ?? asked[0])?.id)
  }

  async function send() {
    setPending("send")
    const sent = await onSend(question, draft)
    setPending(undefined)
    if (!sent) return
    drafts.set(question.id, undefined)
    moveOn()
  }

  // Offered only while the answer box is empty, so no typed answer is lost.
  async function decide() {
    setPending("decide")
    const accepted = await onDecide(question)
    setPending(undefined)
    if (accepted) moveOn()
    else setRefused((count) => count + 1)
  }

  return (
    <section className="question-card">
      <div className="question-head">
        <h2 id={titleId}>{cards.length === 1 ? "Question for you" : `Question ${at + 1} of ${cards.length}`}</h2>
        {cards.length > 1 && (
          <span className="question-nav">
            <button type="button" className="icon-btn question-nav-btn" disabled={at === 0} onClick={() => setChosen(cards[at - 1]!.id)} aria-label="Previous question">
              <Chevron points="15 18 9 12 15 6" />
            </button>
            <button
              type="button"
              className="icon-btn question-nav-btn"
              disabled={at === cards.length - 1}
              onClick={() => setChosen(cards[at + 1]!.id)}
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
              <span key={i} className={dot(i)} />
            ))}
          </span>
          <span>
            {delegated
              ? `${settled} answered, ${delegated} left to the drafter, ${questions.length} to answer`
              : `${answered} answered, ${questions.length} left`}
          </span>
        </div>
      )}
      {about && <p className="question-about">About {about}</p>}
      {left && (
        <p className="question-left">
          <span className="question-dot is-answered is-delegated" aria-hidden />
          Left to the drafter: the review draft says what was chosen and why, or what is still open.
        </p>
      )}
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
      {error && errorFor === question.id && (
        <div className="msg-error" role="alert">
          {error}
        </div>
      )}
      <div className="question-actions">
        <span className="question-keys" aria-hidden>
          ⌘/Ctrl+Enter to send
        </span>
        <span className="question-send">
          {delegable && (
            <button ref={decideButton} type="button" aria-describedby={hintId} disabled={busy || !drafts.loaded || Boolean(draft.trim())} onClick={decide}>
              {busy && pending === "decide" ? "Handing back…" : "Decide for me"}
            </button>
          )}
          <button
            className="primary"
            type="button"
            aria-describedby={rerun ? rerunId : undefined}
            disabled={busy || !drafts.loaded || !draft.trim()}
            onClick={send}
          >
            {busy && pending !== "decide" ? "Sending…" : left ? "Answer instead" : "Send answer"}
          </button>
        </span>
      </div>
      {delegable && (
        <p id={hintId} className="question-delegate-hint">
          You'll get a review draft that says what was chosen and why, or what is still open.
        </p>
      )}
      {rerun && (
        <p id={rerunId} className="question-delegate-hint">
          This draft is finished. Answering now runs the drafting again with your answer.
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
