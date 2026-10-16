import { useAnswerDrafts } from "../../use-answer-drafts"
import AnswerDraftStatus from "../AnswerDraftStatus"
import { useEffect, useRef, useState, type MutableRefObject } from "react"
import { boundedRequest, type Client } from "../../api/opencode/client"
import QuestionDockShell from "./QuestionDockShell"
import { ModalClose } from "../ModalHeader"

type PendingForm = Awaited<ReturnType<Client["session"]["form"]["list"]>>[number]
type Answer = Parameters<Client["session"]["form"]["reply"]>[0]["answer"]

export default function SessionQuestions({
  client,
  sessionRef,
  pendingRef,
  refreshRef,
  directory,
  matterId,
}: {
  client: Client
  sessionRef: MutableRefObject<string>
  pendingRef: MutableRefObject<number>
  refreshRef: MutableRefObject<() => void>
  directory: string
  matterId: string
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
      pendingRef.current = questions.length
      setForms(questions)
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
    try {
      if (dismiss) await client.session.form.cancel({ sessionID: form.sessionID, formID: form.id })
      if (!dismiss)
        await client.session.form.reply({
          sessionID: form.sessionID,
          formID: form.id,
          answer: Object.fromEntries(
            Object.entries(answers[form.id] ?? {}).map(([key, value]) => [
              key,
              Array.isArray(value) ? value.filter((item) => item.trim()) : value,
            ]),
          ),
        })
      if (owner.current !== form.sessionID) return
      drafts.set(form.id, undefined)
      setForms((previous) => previous.filter((row) => row.id !== form.id))
      pendingRef.current = Math.max(0, pendingRef.current - 1)
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
  const answer = answers[form.id] ?? {}
  const update = (key: string, value: string | string[]) =>
    drafts.set(form.id, { ...answers[form.id], [key]: value })
  const values = (key: string) => (Array.isArray(answer[key]) ? (answer[key] as string[]) : [])
  const complete = fields.every((field) =>
    field.type === "multiselect"
      ? values(field.key).some((value) => value.trim())
      : String(answer[field.key] ?? "").trim().length > 0,
  )
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
