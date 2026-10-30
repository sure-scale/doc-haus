import { useEffect, useId, useRef, useState } from "react"
import type { Document } from "../api/ingest/documents"
import { normaliseStarterQuestion, recommendStarter, REVIEW_STARTERS } from "../review-starters"
import ReviewSetupDialog from "./ReviewSetupDialog"
import ReviewExample from "./ReviewExample"
import "./ReviewStarter.css"

type Step = "choose" | "example" | "questions"

// Choose a starter (its example is optional), then review its questions and
// append them to the table. Nothing is written and no model runs before Add
// questions; answers wait for Fill.
export default function ReviewStarter({ matterName, documents, existingQuestions, problem, initialStarter, initialStep = "choose", onCancel, onConfirm }: {
  matterName: string
  documents: readonly Pick<Document, "name" | "role">[]
  existingQuestions: string[]
  problem?: string
  // A starter chosen before the table opened (from the chat's Workflows dialog
  // or the Workflows page). A table with questions opens at its questions step.
  initialStarter?: string
  initialStep?: "choose" | "questions"
  onCancel: () => void
  onConfirm: (questions: string[]) => Promise<boolean>
}) {
  const titleId = useId()
  const [recommended] = useState(() => recommendStarter(documents))
  const [chosenId, setChosenId] = useState(initialStarter ?? recommended.id)
  const [step, setStep] = useState<Step>(initialStep)
  const starter = REVIEW_STARTERS.find((item) => item.id === chosenId)!
  // The questions being edited, kept while the lawyer steps back to the same starter.
  const [draft, setDraft] = useState<{ starter: string; questions: { id: number; question: string }[] } | undefined>(() =>
    initialStep === "questions" ? { starter: starter.id, questions: starter.questions.map((question, id) => ({ id, question })) } : undefined)
  const questions = draft?.questions ?? []
  const setQuestions = (update: (current: { id: number; question: string }[]) => { id: number; question: string }[]) =>
    setDraft((current) => current && { ...current, questions: update(current.questions) })
  const [saving, setSaving] = useState(false)
  const [error, setError] = useState<string>()
  const inputs = useRef(new Map<number, HTMLTextAreaElement>())
  const cancel = useRef<HTMLButtonElement>(null)
  const stepView = useRef<HTMLDivElement>(null)
  const names = documents.map((doc) => doc.name)
  const seen = new Set(existingQuestions.map(normaliseStarterQuestion))
  const entries = questions.map((item) => {
    const question = item.question.trim()
    const key = normaliseStarterQuestion(question)
    const duplicate = key !== "" && seen.has(key)
    if (key) seen.add(key)
    return { ...item, question, duplicate }
  })
  const additions = entries.filter((item) => item.question && !item.duplicate).map((item) => item.question)
  const close = () => { if (!saving) onCancel() }

  // Each step starts where the lawyer acts: the chosen starter, or the top of the new view.
  useEffect(() => {
    stepView.current?.parentElement?.scrollTo({ top: 0 })
    if (step === "choose") stepView.current?.querySelector<HTMLInputElement>("input[type=radio]:checked")?.focus({ preventScroll: true })
    else stepView.current?.focus({ preventScroll: true })
  }, [step])

  function toQuestions() {
    if (draft?.starter !== starter.id) setDraft({ starter: starter.id, questions: starter.questions.map((question, id) => ({ id, question })) })
    setError(undefined)
    setStep("questions")
  }

  async function confirm() {
    if (saving || !documents.length || !additions.length) return
    setSaving(true)
    setError(undefined)
    try {
      const saved = await onConfirm(additions)
      if (saved) onCancel()
      else setError("The questions could not be confirmed. See the details below and close this preview to recover.")
    } catch {
      setError("Questions could not be added. Check the table's save message and try again.")
    } finally {
      setSaving(false)
    }
  }

  const title = step === "choose" ? "Choose a starter" : step === "example" ? `Example: ${starter.name}` : starter.name
  return (
    <ReviewSetupDialog title={title} wide={step === "example"} onClose={close}>
      <div ref={stepView} className="review-starter-step" tabIndex={-1}>
        {step === "choose" && (
          <>
            <div className="review-starter-choose">
              <fieldset className="review-starter-choices">
                <legend className="sr-only">Starters</legend>
                {REVIEW_STARTERS.map((item) => (
                  <label key={item.id} className={`review-starter-choice${item.id === chosenId ? " is-selected" : ""}`}>
                    <input type="radio" name={`${titleId}-starter`} value={item.id} checked={item.id === chosenId} aria-describedby={item.id === chosenId ? `${titleId}-purpose` : undefined} onChange={() => setChosenId(item.id)} />
                    <span>
                      <span className="review-starter-choice-name">{item.name}</span>
                      {item.id === recommended.id && <span className="review-starter-choice-reason"><strong>Recommended</strong> · {recommended.reason}</span>}
                    </span>
                  </label>
                ))}
              </fieldset>
              <section className="review-starter-preview" aria-label={`About ${starter.name}`}>
                <p id={`${titleId}-purpose`}>{starter.purpose}</p>
                <dl>
                  <dt>Inputs</dt><dd>{starter.inputs}</dd>
                  <dt>Output</dt><dd>{starter.output}</dd>
                </dl>
                <ol className="review-starter-preview-questions" aria-label="Questions">
                  {starter.questions.map((question) => <li key={question}>{question}</li>)}
                </ol>
                <p className="review-starter-limits">{starter.limits}</p>
                {starter.example && <button type="button" className="linklike review-starter-example-link" onClick={() => setStep("example")}>See example</button>}
              </section>
            </div>
            <div className="compare-actions">
              <button type="button" className="ghost" onClick={close}>Cancel</button>
              <button type="button" className="primary" onClick={toQuestions}>Continue</button>
            </div>
          </>
        )}

        {step === "example" && starter.example && (
          <>
            <ReviewExample questions={starter.questions} example={starter.example} />
            <div className="compare-actions">
              <button type="button" className="ghost review-starter-back" onClick={() => setStep("choose")}>Back</button>
              <button type="button" className="primary" onClick={toQuestions}>Use with my matter</button>
            </div>
          </>
        )}

        {step === "questions" && (
          <>
            <p><strong>Matter:</strong> {matterName}</p>
            <section aria-label="Documents in scope">
              <h3>Documents in scope ({names.length})</h3>
              {names.length ? <ul>{names.map((name, index) => <li key={`${name}-${index}`}>{name}</li>)}</ul> : <p>Add documents to this matter before using the starter.</p>}
            </section>
            <section aria-label="Questions to add">
              <h3>Questions to add</h3>
              <p>Edit or remove these questions before adding them. Existing table questions and answers stay as they are.</p>
              {entries.map((item, index) => (
                <div className="review-starter-question" key={item.id}>
                  <div className="review-starter-question-head">
                    <label htmlFor={`${titleId}-question-${item.id}`}>Question {index + 1}</label>
                    <button type="button" className="ghost" disabled={saving} onClick={() => {
                      const adjacent = entries[index + 1]?.id ?? entries[index - 1]?.id
                      setQuestions((current) => current.filter((question) => question.id !== item.id))
                      requestAnimationFrame(() => { (adjacent === undefined ? cancel.current : inputs.current.get(adjacent))?.focus() })
                    }} aria-label={`Remove question ${index + 1}`}>Remove</button>
                  </div>
                  <textarea ref={(element) => { if (element) inputs.current.set(item.id, element); else inputs.current.delete(item.id) }} id={`${titleId}-question-${item.id}`} rows={2} value={questions.find((question) => question.id === item.id)?.question ?? ""} disabled={saving} onChange={(event) => setQuestions((current) => current.map((question) => question.id === item.id ? { ...question, question: event.target.value } : question))} />
                  {item.duplicate && <p role="status">Already in the table or repeated above; this question will not be added. Edit or remove it to continue.</p>}
                  {!item.question && <p>Empty questions will not be added.</p>}
                </div>
              ))}
              {!additions.length && <p>No new questions to add. Edit a question or cancel.</p>}
            </section>
            <p><strong>{names.length * additions.length} pending {names.length * additions.length === 1 ? "cell" : "cells"}</strong> for the new questions across these documents.</p>
            <p>Adding questions does not use AI or fill answers. When you are ready, fill the cells from the table; the fill may include other eligible cells.</p>
            {error && <p className="compare-error" role="alert">{error}</p>}
            {error && problem && <p role="alert">{problem}</p>}
            <div className="compare-actions">
              <button type="button" className="ghost review-starter-back" onClick={() => setStep("choose")} disabled={saving}>Back</button>
              <button ref={cancel} type="button" className="ghost" onClick={close} disabled={saving}>Cancel</button>
              <button type="button" className="primary" disabled={saving || !names.length || !additions.length} onClick={() => void confirm()}>{saving ? "Adding…" : "Add questions"}</button>
            </div>
          </>
        )}
      </div>
    </ReviewSetupDialog>
  )
}
