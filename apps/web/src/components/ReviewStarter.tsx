import { useId, useRef, useState } from "react"
import ReviewSetupDialog from "./ReviewSetupDialog"
import "./ReviewStarter.css"

// A definition only: the selected matter supplies all documents and the grid
// remains the sole store for questions and answers.
export const REVIEW_STARTER = {
  id: "document-overview-v1",
  name: "Document overview",
  purpose: "Start a review table with questions about each document in this matter.",
  inputs: "All documents in this matter and the questions below.",
  output: "A review table for answers and sources for each document and question.",
  limits: "Answers may be missing or incomplete. Check the sources and tailor the questions to your matter; this is not a complete legal review.",
  questions: [
    "What is the purpose of this document?",
    "What obligations are described in this document?",
    "What dates or deadlines are stated in this document?",
  ],
} as const

export function normaliseStarterQuestion(question: string) {
  return question.trim().replace(/\s+/g, " ").toLocaleLowerCase()
}

export default function ReviewStarter({ matterName, documents, existingQuestions, problem, onCancel, onConfirm }: {
  matterName: string
  documents: string[]
  existingQuestions: string[]
  problem?: string
  onCancel: () => void
  onConfirm: (questions: string[]) => Promise<boolean>
}) {
  const titleId = useId()
  const [questions, setQuestions] = useState(() => REVIEW_STARTER.questions.map((question, id) => ({ id, question: String(question) })))
  const [saving, setSaving] = useState(false)
  const [error, setError] = useState<string>()
  const inputs = useRef(new Map<number, HTMLTextAreaElement>())
  const cancel = useRef<HTMLButtonElement>(null)
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

  return (
    <ReviewSetupDialog title={REVIEW_STARTER.name} onClose={close}>
          <p>{REVIEW_STARTER.purpose}</p>
          <p><strong>Matter:</strong> {matterName}</p>
          <p><strong>Inputs:</strong> {REVIEW_STARTER.inputs}</p>
          <p><strong>Output:</strong> {REVIEW_STARTER.output}</p>
          <p>{REVIEW_STARTER.limits}</p>
          <section aria-label="Documents in scope">
            <h3>Documents in scope ({documents.length})</h3>
            {documents.length ? <ul>{documents.map((name, index) => <li key={`${name}-${index}`}>{name}</li>)}</ul> : <p>Add documents to this matter before using the starter.</p>}
          </section>
          <section aria-label="Questions to add">
            <h3>Questions to add</h3>
            <p>Edit or remove these questions before adding them. Existing table questions and answers stay as they are.</p>
            {entries.map((item, index) => (
              <div className="review-starter-question" key={item.id}>
                <label htmlFor={`${titleId}-question-${item.id}`}>Question {index + 1}</label>
                <textarea ref={(element) => { if (element) inputs.current.set(item.id, element); else inputs.current.delete(item.id) }} id={`${titleId}-question-${item.id}`} rows={2} value={questions.find((question) => question.id === item.id)?.question ?? ""} disabled={saving} onChange={(event) => setQuestions((current) => current.map((question) => question.id === item.id ? { ...question, question: event.target.value } : question))} />
                {item.duplicate && <p role="status">Already in the table or repeated above; this question will not be added. Edit or remove it to continue.</p>}
                {!item.question && <p>Empty questions will not be added.</p>}
                <button type="button" disabled={saving} onClick={() => {
                  const adjacent = entries[index + 1]?.id ?? entries[index - 1]?.id
                  setQuestions((current) => current.filter((question) => question.id !== item.id))
                  requestAnimationFrame(() => { (adjacent === undefined ? cancel.current : inputs.current.get(adjacent))?.focus() })
                }} aria-label={`Remove question ${index + 1}`}>Remove</button>
              </div>
            ))}
            {!additions.length && <p>No new questions to add. Edit a question or cancel.</p>}
          </section>
          <p><strong>{documents.length * additions.length} pending {documents.length * additions.length === 1 ? "cell" : "cells"}</strong> for the new questions across these documents.</p>
          <p>Adding questions does not use AI or fill answers. When you are ready, fill the cells from the table; the fill may include other eligible cells.</p>
          {error && <p className="compare-error" role="alert">{error}</p>}
          {error && problem && <p role="alert">{problem}</p>}
          <div className="compare-actions">
            <button ref={cancel} type="button" onClick={close} disabled={saving}>Cancel</button>
            <button type="button" className="primary" disabled={saving || !documents.length || !additions.length} onClick={() => void confirm()}>{saving ? "Adding…" : "Add questions"}</button>
          </div>
    </ReviewSetupDialog>
  )
}
