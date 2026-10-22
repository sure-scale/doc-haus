import { useAnswerDrafts } from "../../use-answer-drafts"
import AnswerDraftStatus from "../AnswerDraftStatus"
import { useEffect, useRef, useState, type MutableRefObject } from "react"
import { boundedRequest, type Client } from "../../api/opencode/client"
import QuestionDockShell from "./QuestionDockShell"
import { ModalClose } from "../ModalHeader"
import { routeChoice, routeRequest, type DraftChoice } from "./quick-draft"

type PendingForm = Awaited<ReturnType<Client["session"]["form"]["list"]>>[number]
type Answer = Parameters<Client["session"]["form"]["reply"]>[0]["answer"]

// A field's own default stands until the lawyer changes it.
function defaults(form: PendingForm): Answer {
  return Object.fromEntries(
    form.fields.flatMap((field) => ("default" in field && field.default !== undefined ? [[field.key, field.default]] : [])),
  )
}

export default function SessionQuestions({
  client,
  sessionRef,
  pendingRef,
  refreshRef,
  directory,
  matterId,
  onRoute,
  onRouteDismissed,
  onRoutePending,
}: {
  client: Client
  sessionRef: MutableRefObject<string>
  pendingRef: MutableRefObject<number>
  refreshRef: MutableRefObject<() => void>
  directory: string
  matterId?: string
  // A route form (the quick-draft offer) carries a request, not a question for
  // a model. The engine only records its answer; the chat dispatches it here,
  // once the reply has settled, and takes the request back on Cancel.
  onRoute?: (request: string, choice: DraftChoice) => void
  onRouteDismissed?: (request: string) => void
  onRoutePending?: (request: string | undefined) => void
}) {
  const [forms, setForms] = useState<PendingForm[]>([])
  const [error, setError] = useState("")
  const [replyError, setReplyError] = useState("")
  const [busy, setBusy] = useState(false)
  const [draftSession, setDraftSession] = useState(sessionRef.current)
  const drafts = useAnswerDrafts(matterId, draftSession ? `session.${draftSession}` : "", `dh.session.answers.${directory}.${draftSession}`)
  const answers = drafts.answers as Record<string, Answer>
  const owner = useRef("")
  const syncing = useRef(false)
  async function sync() {
    const sessionID = sessionRef.current
    if (!sessionID || syncing.current) return
    syncing.current = true
    try {
      const next = await boundedRequest(client.session.form.list({ sessionID }), 5000, "Questions load")
      if (owner.current !== sessionID) return
      const questions = next.filter((form) => form.metadata?.kind === "question")
      const routes = next.filter((form) => routeRequest(form) !== undefined)
      // Only a model's question holds a running turn open.
      pendingRef.current = questions.length
      setForms([...routes, ...questions])
      onRoutePending?.(routes[0] && routeRequest(routes[0]))
      setError("")
    } catch (error) {
      if (owner.current === sessionID) setError(error instanceof Error ? error.message : "Could not load questions.")
    } finally {
      syncing.current = false
    }
  }
  // Native session forms remain pending on the server. Polling hydrates reloads
  // and reconciles missed events without creating a matter drafting run.
  useEffect(() => {
    const reconcile = () => {
      if (owner.current !== sessionRef.current) {
        owner.current = sessionRef.current
        pendingRef.current = 0
        setForms([])
        onRoutePending?.(undefined)
        setError("")
        setReplyError("")
        setBusy(false)
        setDraftSession(sessionRef.current)
      }
      void sync()
    }
    reconcile()
    refreshRef.current = reconcile
    const timer = setInterval(reconcile, 15000)
    window.addEventListener("focus", reconcile)
    return () => {
      owner.current = ""
      refreshRef.current = () => {}
      clearInterval(timer)
      window.removeEventListener("focus", reconcile)
      pendingRef.current = 0
    }
  }, [client, sessionRef.current])
  async function settle(form: PendingForm, dismiss = false) {
    setBusy(true)
    setReplyError("")
    const request = routeRequest(form)
    const answer = Object.fromEntries(
      Object.entries({ ...defaults(form), ...answers[form.id] }).map(([key, value]) => [
        key,
        Array.isArray(value) ? value.filter((item) => item.trim()) : value,
      ]),
    )
    try {
      if (dismiss) await client.session.form.cancel({ sessionID: form.sessionID, formID: form.id })
      if (!dismiss) await client.session.form.reply({ sessionID: form.sessionID, formID: form.id, answer })
      if (owner.current !== form.sessionID) return
      drafts.set(form.id, undefined)
      setForms((previous) => previous.filter((row) => row.id !== form.id))
      if (request === undefined) pendingRef.current = Math.max(0, pendingRef.current - 1)
      else {
        onRoutePending?.(undefined)
        // Dispatched only after the engine settled this answer: a form that
        // was already answered elsewhere fails above and starts nothing.
        const choice = routeChoice(answer)
        if (dismiss) onRouteDismissed?.(request)
        else if (choice) onRoute?.(request, choice)
      }
      await sync()
    } catch {
      if (owner.current !== form.sessionID) return
      setReplyError(
        dismiss
          ? "Could not dismiss the questions. Please try again."
          : "Could not send your answers. Your entries are still here; please try again.",
      )
    } finally {
      setBusy(false)
    }
  }
  const form = forms[0]
  if (!form)
    return error ? (
      <div className="msg-error" role="alert">
        {error}
      </div>
    ) : null
  const fields = form.fields.filter((field) => field.type === "string" || field.type === "multiselect")
  const answer = { ...defaults(form), ...answers[form.id] }
  const update = (key: string, value: string | string[]) =>
    drafts.set(form.id, { ...answers[form.id], [key]: value })
  const values = (key: string) => (Array.isArray(answer[key]) ? (answer[key] as string[]) : [])
  const complete = fields.every((field) =>
    field.type === "multiselect"
      ? values(field.key).some((value) => value.trim())
      : String(answer[field.key] ?? "").trim().length > 0,
  )
  const request = routeRequest(form)
  if (request !== undefined) {
    const field = fields[0]
    return (
      <section className="question-card route-offer" key={form.id} aria-labelledby={`${form.id}-title`}>
        <h2 id={`${form.id}-title`} className="question-label">
          {field?.title ?? form.title}
        </h2>
        <div role="radiogroup" aria-labelledby={`${form.id}-title`}>
          {field?.options?.map((option) => (
            <label className="question-label" key={option.value}>
              <input
                type="radio"
                name={`${form.id}-${field.key}`}
                checked={answer[field.key] === option.value}
                disabled={busy || !drafts.loaded}
                onChange={() => update(field.key, option.value)}
              />{" "}
              {option.label}
              {option.description && <span className="question-context"> — {option.description}</span>}
            </label>
          ))}
        </div>
        {(replyError || error) && (
          <div className="msg-error" role="alert">
            {replyError || error}
          </div>
        )}
        <div className="question-actions">
          <button type="button" className="icon-btn" disabled={busy || !drafts.loaded} onClick={() => void settle(form, true)}>
            Cancel
          </button>
          <button type="button" className="primary" disabled={busy || !drafts.loaded || !complete} onClick={() => void settle(form)}>
            {busy ? "Starting…" : "Start"}
          </button>
        </div>
      </section>
    )
  }
  return (
    <QuestionDockShell
      key={form.id}
      count={fields.length}
      preview={fields[0]?.description ?? fields[0]?.title ?? ""}
      subject="The assistant"
      renderQuestions={(titleId, onClose) => (
        <section className="question-card">
          <div className="question-head">
            <h2 id={titleId}>Questions for you</h2>
            <ModalClose onClose={onClose} />
          </div>
          <AnswerDraftStatus {...drafts} />
          {fields.map((field) => (
            <fieldset key={field.key} style={{ border: 0, padding: 0, margin: 0, minWidth: 0 }}>
              <legend className="question-label">{field.title}</legend>
              <p className="question-text">{field.description}</p>
              {field.options?.map((option) => (
                <label className="question-label" key={option.value}>
                  <input
                    type={field.type === "multiselect" ? "checkbox" : "radio"}
                    name={`${form.id}-${field.key}`}
                    checked={
                      field.type === "multiselect"
                        ? values(field.key).includes(option.value)
                        : answer[field.key] === option.value
                    }
                    disabled={busy || !drafts.loaded}
                    onChange={(event) =>
                      update(
                        field.key,
                        field.type === "string"
                          ? option.value
                          : event.target.checked
                            ? [...values(field.key), option.value]
                            : values(field.key).filter((value) => value !== option.value),
                      )
                    }
                  />{" "}
                  {option.label}
                  {option.description && <span className="question-context"> — {option.description}</span>}
                </label>
              ))}
              {field.custom && (
                <textarea
                  id={`${form.id}-${field.key}`}
                  aria-label={`${field.title ?? "Your answer"} — other answer`}
                  className="question-answer"
                  rows={3}
                  disabled={busy || !drafts.loaded}
                  placeholder={
                    field.type === "multiselect" ? "Other answers, one per line." : "Or answer in your own words."
                  }
                  value={
                    field.type === "multiselect"
                      ? values(field.key)
                          .filter((value) => !field.options?.some((option) => option.value === value))
                          .join("\n")
                      : typeof answer[field.key] === "string" &&
                          !field.options?.some((option) => option.value === answer[field.key])
                        ? String(answer[field.key])
                        : ""
                  }
                  onChange={(event) =>
                    update(
                      field.key,
                      field.type === "multiselect"
                        ? [
                            ...values(field.key).filter((value) =>
                              field.options?.some((option) => option.value === value),
                            ),
                            ...event.target.value.split("\n"),
                          ]
                        : event.target.value,
                    )
                  }
                />
              )}
            </fieldset>
          ))}
          {(replyError || error) && (
            <div className="msg-error" role="alert">
              {replyError || error}
            </div>
          )}
          <div className="question-actions">
            <button type="button" className="icon-btn" disabled={busy || !drafts.loaded} onClick={() => void settle(form, true)}>
              Dismiss questions
            </button>
            <button type="button" className="primary" disabled={busy || !drafts.loaded || !complete} onClick={() => void settle(form)}>
              {busy ? "Sending…" : "Send answers"}
            </button>
          </div>
        </section>
      )}
    />
  )
}
