import { useAnswerDrafts } from "../../use-answer-drafts"
import AnswerDraftStatus from "../AnswerDraftStatus"
import { useEffect, useId, useRef, useState } from "react"
import type { LeanQuestion } from "../../api/ingest/lean"
import { ModalClose } from "../ModalHeader"
import { delegableQuestion, questionGroups, questionOptions, recommendedOption } from "./model"

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
  questions: asked,
  answered,
  answeredQuestions = [],
  queuedAnswerIDs = [],
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
  answeredQuestions?: LeanQuestion[]
  queuedAnswerIDs?: string[]
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
  const [view, setView] = useState<"open" | "answered">("open")
  const [editing, setEditing] = useState<string>()
  // Which button's request is in flight, so each reads as its own.
  const [pending, setPending] = useState<"send" | "decide" | { option: string }>()
  // Counts refused Decide for me requests: the cursor goes back to it.
  const [refused, setRefused] = useState(0)
  const field = useRef<HTMLTextAreaElement>(null)
  const decideButton = useRef<HTMLButtonElement>(null)
  const changeButton = useRef<HTMLButtonElement>(null)
  const queuedNotice = useRef<HTMLParagraphElement>(null)
  const groups = questionGroups(asked)
  const questions = groups.flatMap(group => group.questions)
  const reviewing = view === "answered" || questions.length === 0
  const cards = reviewing ? answeredQuestions : questions
  const found = cards.findIndex((question) => question.id === chosen)
  const at = found === -1 ? 0 : found
  const shown = cards[at]?.id

  // A new question (the next one once an answer is sent) gets the cursor.
  useEffect(() => {
    if (busy) return
    if (reviewing && queuedAnswerIDs.includes(shown ?? "")) queuedNotice.current?.focus({ preventScroll: true })
    else if (reviewing && editing !== shown) changeButton.current?.focus({ preventScroll: true })
    else field.current?.focus({ preventScroll: true })
  }, [shown, drafts.loaded, reviewing, editing, busy, queuedAnswerIDs])

  useEffect(() => {
    if (refused) decideButton.current?.focus()
  }, [refused])

  if (cards.length === 0) return null
  const question = cards[at]!
  const draftKey = reviewing ? `correction:${question.id}` : question.id
  const draft = typeof drafts.answers[draftKey] === "string" ? drafts.answers[draftKey] as string : ""
  const about = question.document && titles.size > 1 ? titles.get(question.document) : undefined
  const fieldId = `lean-answer-${question.id}`
  const total = answered + questions.length
  const current = answered + at
  const options = questionOptions(question)
  const recommended = recommendedOption(question)
  const advice = question.advice
  const recommendation = recommended && advice && "recommendation" in advice ? advice.recommendation : undefined
  const unavailable = advice && "unavailable" in advice ? advice.unavailable : undefined
  const delegable = !reviewing && delegableQuestion(question)
  const group = groups.find(row => row.questions.some(card => card.id === question.id))
  const grouped = groups.some(row => row.questions.length > 1)
  const correctionQueued = reviewing && queuedAnswerIDs.includes(question.id)
  const correcting = reviewing && editing === question.id && !correctionQueued
  const canSend = !busy && drafts.loaded && Boolean(draft.trim()) && (!reviewing || draft.trim() !== question.answer?.trim())
  // An option or Decide for me never replaces words typed in Other.
  const picking = busy || !drafts.loaded || Boolean(draft.trim())

  // Once this question is answered, the next one still asked.
  function moveOn() {
    const rest = questions.filter((card) => card.id !== question.id)
    setChosen((rest.find((card) => questions.indexOf(card) > at) ?? rest[0])?.id)
  }

  async function settle(kind: NonNullable<typeof pending>, work: Promise<boolean>) {
    setPending(kind)
    let accepted = false
    try { accepted = await work } finally { setPending(undefined) }
    if (accepted) {
      drafts.set(draftKey, undefined)
      if (reviewing) setEditing(undefined)
      else moveOn()
    } else if (reviewing) {
      field.current?.focus({ preventScroll: true })
    }
    return accepted
  }

  return (
    <section className="question-card">
      <div className="question-head">
        <h2 id={titleId}>{reviewing ? `Answer ${at + 1} of ${cards.length}` : questions.length === 1 ? "Question for you" : `Question ${at + 1} of ${questions.length}`}</h2>
        {cards.length > 1 && (
          <span className="question-nav">
            <button type="button" className="icon-btn question-nav-btn" disabled={at === 0 || busy} onClick={() => setChosen(cards[at - 1]!.id)} aria-label={reviewing ? "Previous answer" : "Previous question"}>
              <Chevron points="15 18 9 12 15 6" />
            </button>
            <button
              type="button"
              className="icon-btn question-nav-btn"
              disabled={at === cards.length - 1 || busy}
              onClick={() => setChosen(cards[at + 1]!.id)}
              aria-label={reviewing ? "Next answer" : "Next question"}
            >
              <Chevron points="9 18 15 12 9 6" />
            </button>
          </span>
        )}
        <ModalClose onClose={onClose} />
      </div>
      {answeredQuestions.length > 0 && questions.length > 0 && (
        <div className="question-views" role="group" aria-label="Question view">
          <button type="button" aria-pressed={!reviewing} disabled={busy} onClick={() => { setView("open"); setChosen(undefined) }}>To answer ({questions.length})</button>
          <button type="button" aria-pressed={reviewing} disabled={busy} onClick={() => { setView("answered"); setChosen(undefined) }}>Answered ({answeredQuestions.length})</button>
        </div>
      )}
      {!reviewing && total > 1 && (
        <div className="question-progress">
          <span className="question-dots" aria-hidden>
            {Array.from({ length: total }, (_, i) => (
              <span key={i} className={`question-dot${i < answered ? " is-answered" : ""}${i === current ? " is-current" : ""}`} />
            ))}
          </span>
          <span>{`${answered} answered, ${questions.length} left`}</span>
        </div>
      )}
      {!reviewing && grouped && group && (
        <div className="question-topics">
          <label className="question-label" htmlFor={`${fieldId}-topic`}>Topic</label>
          <select id={`${fieldId}-topic`} value={group.id} disabled={busy} onChange={event => setChosen(groups.find(row => row.id === event.target.value)?.questions[0]?.id)}>
            {groups.map(row => <option key={row.id} value={row.id}>{row.label} ({row.questions.length})</option>)}
          </select>
          {group.questions.length > 1 && (
            <div className="question-related" role="group" aria-label={`Questions about ${group.label}`}>
              {group.questions.map((row, index) => <button type="button" key={row.id} aria-current={row.id === question.id ? "step" : undefined} disabled={busy} onClick={() => setChosen(row.id)}>{index + 1}. {row.question}</button>)}
            </div>
          )}
        </div>
      )}
      {about && <p className="question-about">About {about}</p>}
      <p className="question-text">{question.question}</p>
      {question.context && <p className="question-context">{question.context}</p>}
      {!reviewing && options.length > 0 && (
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
                  {passage.relevance && <p className="question-basis-relevance">{passage.relevance}</p>}
                </li>
              ))}
            </ul>
          )}
        </div>
      )}
      {unavailable && <p className="question-advice is-unavailable">No recommendation: {unavailable.reason}</p>}
      {reviewing && (
        <div className="question-saved-answer">
          <p className="question-label">{question.answer_kind === "adopted" ? "Adopted answer" : "Saved answer"}</p>
          <p className="question-saved-text">{question.answer}</p>
          {correctionQueued ? <p ref={queuedNotice} role="status" tabIndex={-1}>Correction queued. Your saved answer will update when drafting applies it.</p> : !correcting && (
            <button ref={changeButton} type="button" disabled={busy || !drafts.loaded} onClick={() => {
              if (typeof drafts.answers[draftKey] !== "string") drafts.set(draftKey, question.answer ?? "")
              setEditing(question.id)
            }}>Change answer</button>
          )}
        </div>
      )}
      {(!reviewing || correcting) && <>
      <label className="question-label" htmlFor={fieldId}>
        {reviewing ? "Corrected answer" : options.length > 0 ? "Other" : "Your answer"}
      </label>
      <textarea
        ref={field}
        id={fieldId}
        className="question-answer"
        value={draft}
        disabled={busy || !drafts.loaded}
        onChange={(e) => drafts.set(draftKey, e.target.value)}
        onKeyDown={(e) => {
          if (e.key === "Enter" && (e.metaKey || e.ctrlKey) && canSend) settle("send", onSend(question, draft))
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
          {reviewing && <button type="button" disabled={busy} onClick={() => { drafts.set(draftKey, undefined); setEditing(undefined) }}>Cancel</button>}
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
          <button className="primary" type="button" disabled={!canSend} onClick={() => settle("send", onSend(question, draft))}>
            {busy && (pending ?? "send") === "send" ? reviewing ? "Saving correction…" : "Sending…" : reviewing ? "Save correction" : "Send answer"}
          </button>
        </span>
      </div>
      </>}
      {reviewing && error && errorFor === question.id && !correcting && <div className="msg-error" role="alert">{error}</div>}
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
