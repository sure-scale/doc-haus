import { useEffect, useId, useRef, useState } from "react"
import type { Document } from "../api/ingest/documents"
import { DESCRIPTION_CHARS, MAX_SUGGESTIONS } from "../api/review-suggest"
import { normaliseStarterQuestion, recommendStarter, REVIEW_STARTERS, type ReviewStarterDefinition, type StarterQuestion } from "../review-starters"
import ReviewSetupDialog from "./ReviewSetupDialog"
import ReviewExample from "./ReviewExample"
import StarterSummary from "./StarterSummary"
import "./ReviewStarter.css"

type Step = "choose" | "example" | "describe" | "questions"
// A question being edited keeps its starter's compare flag through any edit;
// `suggested` marks one the AI suggested, until it is edited or added.
type Draft = { id: number; question: string; compare?: false; suggested?: true }
const draftOf = (starter: ReviewStarterDefinition): Draft[] => starter.questions.map(({ text, compare }, id) => ({ id, question: text, compare }))
// The chooser's entry for AI suggestions, beside the starters.
const AI = "ai"
const AI_NAME = "Suggest questions with AI"
const SUGGEST_TIMEOUT = 120_000

export type SuggestQuestions = (description: string, existing: string[], signal: AbortSignal) => Promise<StarterQuestion[] | undefined>

function suggestFailure(error: unknown) {
  if (error instanceof Error && error.name === "TimeoutError") return "The suggestions took too long."
  if (error instanceof Error && error.message.startsWith("Choose a model in Settings")) return error.message
  return "Questions could not be suggested."
}

// Choose a starter (its example is optional) or describe what to find out and
// have AI suggest questions, then review the questions and append them to the
// table. Suggesting is one disclosed AI call on the description, the document
// names and the chosen questions; nothing is written before Add questions, and
// nothing reads the documents' text until Fill.
export default function ReviewStarter({ documents, existingQuestions, problem, initialStarter, initialStep = "choose", suggest, description: remembered = "", onDescription, onCancel, onConfirm }: {
  documents: readonly Pick<Document, "name" | "role">[]
  existingQuestions: string[]
  problem?: string
  // A starter chosen before the table opened (from the chat's Workflows dialog
  // or the Workflows page), which the chooser opens with.
  initialStarter?: string
  initialStep?: "choose" | "describe"
  // Without it, the chooser offers the starters only.
  suggest?: SuggestQuestions
  // The last description, offered again; each one used is passed back.
  description?: string
  onDescription?: (description: string) => void
  onCancel: () => void
  onConfirm: (questions: { question: string; compare?: false }[]) => Promise<boolean>
}) {
  const titleId = useId()
  const [recommended] = useState(() => recommendStarter(documents))
  const [chosenId, setChosenId] = useState(initialStep === "describe" && suggest ? AI : initialStarter ?? recommended.id)
  // A starter the lawyer already chose is not second-guessed with a Recommended beside it.
  const recommendedId = !initialStarter || initialStarter === recommended.id ? recommended.id : undefined
  const [step, setStep] = useState<Step>(initialStep === "describe" && !suggest ? "choose" : initialStep)
  const ai = chosenId === AI
  const starter = REVIEW_STARTERS.find((item) => item.id === chosenId)
  // The questions being edited, kept while the lawyer steps back to the same choice.
  const [draft, setDraft] = useState<{ source: string; questions: Draft[] }>()
  const questions = draft?.questions ?? []
  const setQuestions = (update: (current: Draft[]) => Draft[]) =>
    setDraft((current) => current && { ...current, questions: update(current.questions) })
  const [saving, setSaving] = useState(false)
  const [error, setError] = useState<string>()
  // Describe: the lawyer's words, where Back returns to, the call in flight and its quiet failure.
  const [description, setDescription] = useState(remembered)
  const [describeFrom, setDescribeFrom] = useState<"choose" | "questions">("choose")
  const [suggesting, setSuggesting] = useState<AbortController>()
  const [suggestError, setSuggestError] = useState<string>()
  const inFlight = useRef<AbortController>(undefined)
  useEffect(() => () => inFlight.current?.abort(), [])
  const firstNew = useRef<number>(undefined)
  const inputs = useRef(new Map<number, HTMLTextAreaElement>())
  const cancel = useRef<HTMLButtonElement>(null)
  const stepView = useRef<HTMLDivElement>(null)
  const describeField = useRef<HTMLTextAreaElement>(null)
  const names = documents.map((doc) => doc.name)
  const documentCount = `${names.length} ${names.length === 1 ? "document" : "documents"}`
  const seen = new Set(existingQuestions.map(normaliseStarterQuestion))
  const entries = questions.map((item) => {
    const question = item.question.trim()
    const key = normaliseStarterQuestion(question)
    const duplicate = key !== "" && seen.has(key)
    if (key) seen.add(key)
    return { ...item, question, duplicate }
  })
  const additions = entries.filter((item) => item.question && !item.duplicate).map(({ question, compare }) => ({ question, compare }))
  const close = () => { if (!saving) onCancel() }

  // Each step starts where the lawyer acts: the chosen entry, the description,
  // the first new suggestion or the top of the new view.
  useEffect(() => {
    stepView.current?.parentElement?.scrollTo({ top: 0 })
    if (step === "choose") stepView.current?.querySelector<HTMLInputElement>("input[type=radio]:checked")?.focus({ preventScroll: true })
    else if (step === "describe") {
      const field = describeField.current
      field?.focus({ preventScroll: true })
      field?.setSelectionRange(field.value.length, field.value.length)
    } else if (step === "questions" && firstNew.current !== undefined) {
      const field = inputs.current.get(firstNew.current)
      firstNew.current = undefined
      field?.focus({ preventScroll: true })
      field?.scrollIntoView({ block: "center" })
    } else stepView.current?.focus({ preventScroll: true })
  }, [step])

  function toQuestions() {
    setError(undefined)
    if (ai) {
      // A list already suggested is kept; otherwise describe what to find out.
      if (draft?.source === AI) setStep("questions")
      else describe("choose")
      return
    }
    if (starter && draft?.source !== starter.id) setDraft({ source: starter.id, questions: draftOf(starter) })
    setStep("questions")
  }

  function describe(from: "choose" | "questions") {
    setDescribeFrom(from)
    setSuggestError(undefined)
    setStep("describe")
  }

  async function runSuggest() {
    const text = description.trim()
    if (!suggest || !text || inFlight.current) return
    onDescription?.(text)
    // A list being added to keeps its questions; the AI sees them as chosen.
    const appending = describeFrom === "questions" && draft !== undefined
    const listed = appending ? questions.map((item) => item.question.trim()).filter(Boolean) : []
    const controller = new AbortController()
    inFlight.current = controller
    setSuggesting(controller)
    setSuggestError(undefined)
    try {
      const suggested = await suggest(text, [...existingQuestions, ...listed], AbortSignal.any([controller.signal, AbortSignal.timeout(SUGGEST_TIMEOUT)]))
      if (controller.signal.aborted) return
      if (!suggested) { setSuggestError("The suggestions could not be read."); return }
      if (!suggested.length) { setSuggestError("No new questions came back. Try describing it differently."); return }
      const base = appending ? questions : []
      let next = base.reduce((max, item) => Math.max(max, item.id + 1), 0)
      const added = suggested.map(({ text: question, compare }): Draft => ({ id: next++, question, compare, suggested: true }))
      firstNew.current = appending ? added[0]?.id : undefined
      setDraft({ source: appending ? draft.source : AI, questions: [...base, ...added] })
      setError(undefined)
      setStep("questions")
    } catch (e) {
      if (!controller.signal.aborted) setSuggestError(suggestFailure(e))
    } finally {
      if (inFlight.current === controller) inFlight.current = undefined
      setSuggesting((current) => current === controller ? undefined : current)
    }
  }

  function stopSuggesting() {
    suggesting?.abort()
    describeField.current?.focus()
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

  const listName = draft?.source === AI ? "Suggested questions" : REVIEW_STARTERS.find((item) => item.id === draft?.source)?.name ?? "Questions"
  const title = step === "choose" ? "Add questions" : step === "example" && starter ? `Example: ${starter.name}` : step === "describe" ? AI_NAME : listName
  return (
    <ReviewSetupDialog title={title} variant={step === "example" ? "wide" : "medium"} onClose={close}>
      <div ref={stepView} className="review-starter-step" tabIndex={-1}>
        {step === "choose" && (
          <>
            <div className="review-starter-choose">
              <fieldset className="review-starter-choices">
                <legend className="sr-only">Ways to add questions</legend>
                {REVIEW_STARTERS.map((item) => (
                  <label key={item.id} className={`review-starter-choice${item.id === chosenId ? " is-selected" : ""}`}>
                    <input type="radio" name={`${titleId}-starter`} value={item.id} checked={item.id === chosenId} aria-describedby={item.id === chosenId ? `${titleId}-purpose` : undefined} onChange={() => setChosenId(item.id)} />
                    <span>
                      <span className="review-starter-choice-name">{item.name}</span>
                      {item.id === recommendedId && <span className="review-starter-choice-reason"><strong>Recommended</strong> · {recommended.reason}</span>}
                    </span>
                  </label>
                ))}
                {suggest && (
                  <label className={`review-starter-choice is-ai${ai ? " is-selected" : ""}`}>
                    <input type="radio" name={`${titleId}-starter`} value={AI} checked={ai} aria-describedby={ai ? `${titleId}-purpose` : undefined} onChange={() => setChosenId(AI)} />
                    <span>
                      <span className="review-starter-choice-name">{AI_NAME}</span>
                      <span className="review-starter-choice-reason">Describe what you want to find out</span>
                    </span>
                  </label>
                )}
              </fieldset>
              {starter ? (
                <section className="review-starter-preview" aria-label={`About ${starter.name}`}>
                  <StarterSummary starter={starter} questions purposeId={`${titleId}-purpose`} />
                  {starter.example && <button type="button" className="linklike review-starter-example-link" onClick={() => setStep("example")}>See example</button>}
                </section>
              ) : (
                <section className="review-starter-preview" aria-label={`About ${AI_NAME}`}>
                  <div className="starter-summary">
                    <p id={`${titleId}-purpose`}>Describe what you want to find out from this matter's documents, and AI suggests questions to ask of each one.</p>
                    <dl>
                      <dt>AI sees</dt><dd>Your description, the names of the {documentCount} in this matter and the questions already in the table. Not the documents' text.</dd>
                      <dt>Output</dt><dd>Up to {MAX_SUGGESTIONS} suggested questions, which you edit or remove before adding them.</dd>
                    </dl>
                    <p className="starter-summary-limits">One AI call when you choose Suggest questions. Answers wait for Fill. Suggestions may miss what matters in your matter; check and tailor them.</p>
                  </div>
                </section>
              )}
            </div>
            <div className="compare-actions">
              <button type="button" className="ghost" onClick={close}>Cancel</button>
              <button type="button" className="primary" onClick={toQuestions}>Continue</button>
            </div>
          </>
        )}

        {step === "example" && starter?.example && (
          <>
            <ReviewExample questions={starter.questions} example={starter.example} />
            <div className="compare-actions">
              <button type="button" className="ghost review-starter-back" onClick={() => setStep("choose")}>Back</button>
              <button type="button" className="primary" onClick={toQuestions}>Use with my matter</button>
            </div>
          </>
        )}

        {step === "describe" && (
          <>
            <label className="review-question-field" htmlFor={`${titleId}-describe`}>What do you want to find out from these documents?</label>
            <textarea ref={describeField} id={`${titleId}-describe`} className="review-question-input review-suggest-input" rows={4} maxLength={DESCRIPTION_CHARS} value={description} readOnly={suggesting !== undefined}
              aria-describedby={`${titleId}-sees${suggestError ? ` ${titleId}-suggest-error` : ""}`}
              onChange={(event) => { setDescription(event.target.value); setSuggestError(undefined) }}
              onKeyDown={(event) => {
                if (event.key !== "Enter" || !(event.metaKey || event.ctrlKey) || event.nativeEvent.isComposing) return
                event.preventDefault()
                void runSuggest()
              }} />
            <p id={`${titleId}-sees`} className="review-question-hint">The AI reads this, the names of the {documentCount} and the questions already chosen, not the documents' text.</p>
            {suggestError && (
              <p id={`${titleId}-suggest-error`} className="review-suggest-error" role="status">
                {suggestError} <button type="button" className="linklike" onClick={() => void runSuggest()}>Try again</button>
              </p>
            )}
            <div className="compare-actions">
              {suggesting ? (
                <>
                  <p className="review-suggest-running" role="status">Suggesting questions…</p>
                  <button type="button" className="ghost" onClick={stopSuggesting}>Cancel</button>
                </>
              ) : (
                <>
                  <button type="button" className="ghost review-starter-back" onClick={() => setStep(describeFrom)}>Back</button>
                  <button type="button" className="ghost" onClick={close}>Cancel</button>
                  <button type="button" className="primary" disabled={!description.trim() || !names.length} onClick={() => void runSuggest()}>Suggest questions</button>
                </>
              )}
            </div>
          </>
        )}

        {step === "questions" && (
          <>
            <p>{draft?.source === AI ? "AI suggested these questions from your description. " : ""}Edit or remove these questions before adding them; existing questions and answers stay as they are.</p>
            <ol className="review-starter-questions" aria-label="Questions to add">
              {entries.map((item, index) => {
                const field = `${titleId}-question-${item.id}`
                const note = item.duplicate ? "Already in the table or listed above, so it will not be added."
                  : !item.question ? "Empty questions will not be added." : undefined
                // The AI's own list is titled as suggested; a starter's tags what the AI added.
                const tagged = item.suggested && draft?.source !== AI
                const described = [tagged && `${field}-tag`, note && `${field}-note`].filter(Boolean).join(" ")
                return (
                  <li className={`review-starter-question${tagged ? " is-suggested" : ""}`} key={item.id}>
                    <label htmlFor={field}><span className="sr-only">Question </span>{index + 1}</label>
                    <textarea ref={(element) => { if (element) inputs.current.set(item.id, element); else inputs.current.delete(item.id) }} id={field} rows={1} placeholder="Write a question" aria-describedby={described || undefined} value={questions.find((question) => question.id === item.id)?.question ?? ""} disabled={saving} onChange={(event) => setQuestions((current) => current.map((question) => question.id === item.id ? { ...question, question: event.target.value, suggested: undefined } : question))} />
                    {tagged && <span id={`${field}-tag`} className="review-starter-tag">Suggested<span className="sr-only"> by AI</span></span>}
                    <button type="button" className="ghost" disabled={saving} onClick={() => {
                      const adjacent = entries[index + 1]?.id ?? entries[index - 1]?.id
                      setQuestions((current) => current.filter((question) => question.id !== item.id))
                      requestAnimationFrame(() => { (adjacent === undefined ? cancel.current : inputs.current.get(adjacent))?.focus() })
                    }} aria-label={`Remove question ${index + 1}`}>Remove</button>
                    {note && <p id={`${field}-note`} role={item.duplicate ? "status" : undefined}>{note}</p>}
                  </li>
                )
              })}
            </ol>
            {suggest && <button type="button" className="linklike review-starter-more" disabled={saving} onClick={() => describe("questions")}>Suggest more with AI</button>}
            <p className="review-starter-scope">{names.length ? `${documentCount}: ${names.join(", ")}` : "Add documents to this matter before adding questions."}</p>
            {!additions.length ? <p className="review-starter-summary">No new questions to add. Edit a question or cancel.</p>
              : names.length > 0 && <p className="review-starter-summary">Adds {additions.length} {additions.length === 1 ? "question" : "questions"} · {names.length * additions.length} {names.length * additions.length === 1 ? "cell" : "cells"} to fill. No AI runs until you choose Fill.</p>}
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
