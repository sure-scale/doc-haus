import { useAnswerDrafts } from "../../use-answer-drafts"
import AnswerDraftStatus from "../AnswerDraftStatus"
import { useEffect, useId, useRef, useState } from "react"
import type { LeanQuestion } from "../../api/ingest/lean"
import { ModalClose } from "../ModalHeader"
import { delegableQuestion, questionOptions, recommendedOption } from "./model"

// The run's open questions, one at a time with an "N of M" counter and how many
// are answered, inside the question dialog. A question's options are buttons in
// the run's order, none chosen for the lawyer; the recommended one is marked,
// with why and the passages behind it. Other, the lawyer's own words, is always
// offered, and answers in progress there are kept per run on the server. Decide
// for me answers with the recommended option where the question allows it.
// Choosing an option or Decide for me waits while Other holds words, so no
// typed answer is lost. The cursor sits in Other.
export default function Questions({
  matterId,
  ownerSessionID,
  questions,
  answered,
  titles,
  busy,
  error,
  errorFor,
  titleId,
  onClose,
  onSend,
  onChoose,
  onDecide,
}: {
  matterId: string
  ownerSessionID: string
  questions: LeanQuestion[]
  answered: number
  titles: Map<string, string>
  busy: boolean
  error: string
  /** The question the error is about; it shows only with that question. */
  errorFor?: string
  /** Id for the heading, which names the dialog. */
  titleId: string
  onClose: () => void
  onSend: (question: LeanQuestion, answer: string) => Promise<boolean>
  onChoose: (question: LeanQuestion, optionID: string) => Promise<boolean>
  onDecide: (question: LeanQuestion) => Promise<boolean>
}) {
  const hintId = useId()
  const choicesId = useId()
  const whyId = useId()
  const drafts = useAnswerDrafts(matterId, `lean.${ownerSessionID}`, `dh.lean.answers.${matterId}`)
  // The question chosen to show; until one is (or once it leaves the card), the first.
  const [chosen, setChosen] = useState<string>()
  // Which button's request is in flight, so each reads as its own.
  const [pending, setPending] = useState<"send" | "decide" | { option: string }>()
  // Counts refused Decide for me requests: the cursor goes back to it.
  const [refused, setRefused] = useState(0)
  const field = useRef<HTMLTextAreaElement>(null)
  const decideButton = useRef<HTMLButtonElement>(null)
  const found = questions.findIndex((question) => question.id === chosen)
  const at = found === -1 ? 0 : found
  const shown = questions[at]?.id

  // A new question (the next one once an answer is sent) gets the cursor.
  useEffect(() => {
    field.current?.focus({ preventScroll: true })
  }, [shown, drafts.loaded])

  useEffect(() => {
    if (refused) decideButton.current?.focus()
  }, [refused])

  if (questions.length === 0) return null
  const question = questions[at]!
  const draft = typeof drafts.answers[question.id] === "string" ? drafts.answers[question.id] as string : ""
  const about = question.document && titles.size > 1 ? titles.get(question.document) : undefined
  const fieldId = `lean-answer-${question.id}`
  const total = answered + questions.length
  const current = answered + at
  const options = questionOptions(question)
  const recommended = recommendedOption(question)
  const advice = question.advice
  const recommendation = recommended && advice && "recommendation" in advice ? advice.recommendation : undefined
  const unavailable = advice && "unavailable" in advice ? advice.unavailable : undefined
  const delegable = delegableQuestion(question)
  // An option or Decide for me never replaces words typed in Other.
  const picking = busy || !drafts.loaded || Boolean(draft.trim())

  // Once this question is answered, the next one still asked.
  function moveOn() {
    const rest = questions.filter((card) => card.id !== question.id)
    setChosen((rest.find((card) => questions.indexOf(card) > at) ?? rest[0])?.id)
  }

  async function settle(kind: NonNullable<typeof pending>, work: Promise<boolean>) {
    setPending(kind)
    const accepted = await work
    setPending(undefined)
    if (accepted) {
      drafts.set(question.id, undefined)
      moveOn()
    }
    return accepted
  }

  return (
    <section className="question-card">
      <div className="question-head">
        <h2 id={titleId}>{questions.length === 1 ? "Question for you" : `Question ${at + 1} of ${questions.length}`}</h2>
        {questions.length > 1 && (
          <span className="question-nav">
            <button type="button" className="icon-btn question-nav-btn" disabled={at === 0} onClick={() => setChosen(questions[at - 1]!.id)} aria-label="Previous question">
              <Chevron points="15 18 9 12 15 6" />
            </button>
            <button
              type="button"
              className="icon-btn question-nav-btn"
              disabled={at === questions.length - 1}
              onClick={() => setChosen(questions[at + 1]!.id)}
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
              <span key={i} className={`question-dot${i < answered ? " is-answered" : ""}${i === current ? " is-current" : ""}`} />
            ))}
          </span>
          <span>{`${answered} answered, ${questions.length} left`}</span>
        </div>
      )}
      {about && <p className="question-about">About {about}</p>}
      <p className="question-text">{question.question}</p>
      {question.context && <p className="question-context">{question.context}</p>}
      {options.length > 0 && (
        <>
          <span id={choicesId} className="question-label">
            Your answer
          </span>
          <div role="group" aria-labelledby={choicesId} className="question-options">
            {options.map((option) => {
              const marked = option === recommended
              return (
                <button
                  key={option.id}
                  type="button"
                  className={marked ? "question-option is-recommended" : "question-option"}
                  aria-describedby={marked && recommendation ? whyId : undefined}
                  disabled={picking}
                  onClick={() => settle({ option: option.id }, onChoose(question, option.id))}
                >
                  <span className="question-option-text">
                    {busy && typeof pending === "object" && pending.option === option.id ? "Sending…" : option.text}
                  </span>
                  {marked && <span className="question-option-tag">Recommended</span>}
                </button>
              )
            })}
          </div>
        </>
      )}
      {recommendation && (
        <div id={whyId} className="question-advice">
          <p className="question-advice-why">
            <span className="question-advice-label">Why</span> {recommendation.rationale}
          </p>
          {recommendation.basis.length > 0 && (
            <ul className="question-advice-basis">
              {recommendation.basis.map((passage, index) => (
                <li key={index}>
                  <blockquote>{passage.quote}</blockquote>
                  <cite>{passage.source_id === "request" ? "Your instructions" : passage.source_id}</cite>
                </li>
              ))}
            </ul>
          )}
        </div>
      )}
      {unavailable && <p className="question-advice is-unavailable">No recommendation: {unavailable.reason}</p>}
      <label className="question-label" htmlFor={fieldId}>
        {options.length > 0 ? "Other" : "Your answer"}
      </label>
      <textarea
        ref={field}
        id={fieldId}
        className="question-answer"
        value={draft}
        disabled={busy || !drafts.loaded}
        onChange={(e) => drafts.set(question.id, e.target.value)}
        onKeyDown={(e) => {
          if (e.key === "Enter" && (e.metaKey || e.ctrlKey) && !busy && drafts.loaded && draft.trim()) settle("send", onSend(question, draft))
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
            <button
              ref={decideButton}
              type="button"
              aria-describedby={hintId}
              disabled={picking}
              onClick={async () => {
                if (!(await settle("decide", onDecide(question)))) setRefused((count) => count + 1)
              }}
            >
              {busy && pending === "decide" ? "Sending…" : "Decide for me"}
            </button>
          )}
          <button className="primary" type="button" disabled={busy || !drafts.loaded || !draft.trim()} onClick={() => settle("send", onSend(question, draft))}>
            {busy && (pending ?? "send") === "send" ? "Sending…" : "Send answer"}
          </button>
        </span>
      </div>
      {delegable && (
        <p id={hintId} className="question-delegate-hint">
          Decide for me answers with the recommended option.
        </p>
      )}
    </section>
  )
}

function Chevron({ points }: { points: string }) {
  return (
    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round" aria-hidden>
      <polyline points={points} />
    </svg>
  )
}
