import { useEffect, useId, useRef, useState } from "react"
import { Link } from "react-router-dom"
import type { LeanDraftingStatus } from "../../api/ingest/lean"
import { formatAgo, formatElapsed } from "../../lib/format"
import { activeStep, stepActivityLabel, stoppedByTechnicalProblem } from "../narration/activity"
import { draftedIn } from "../narration/timeline"
import {
  afterCheckChange,
  attentionItems,
  checklistFindingsCount,
  cleanTitle,
  currentStep,
  delegableQuestion,
  draftedGroups,
  finalCheckNote,
  findingsTitle,
  groupOpenItems,
  leanContinueAvailable,
  leanRunStopped,
  leanSettledItems,
  leanSettledSummary,
  leanStopAvailable,
  notesWording,
  LONG_STEP_MS,
  openItemsSentence,
  openQuestions,
  runPosition,
  runState,
  runStepName,
  type RunState,
  stageElapsed,
  stepCause,
  stepGroups,
} from "./model"
import OpenItems from "./OpenItems"
import ChecklistRefinement from "./ChecklistRefinement"
import Stepper from "./Stepper"
import type { RunActions } from "./useRunActions"

// The drafting run at the top of its conversation: one line on where it is, the
// one thing to do next, and a segmented bar of its steps. The step list, open
// items and other detail stay behind Details. Questions are answered from the
// dock above the composer, so the header only counts them. Stop asks once, in
// place, before it ends a run that nothing is working on. In any other conversation it is a
// single line that leads back to the run's own conversation.
export default function RunHeader({
  status,
  actions,
  own,
  statusError,
  checkedAt,
  onOpen,
  onShowDrafts,
}: {
  status: LeanDraftingStatus
  actions: RunActions
  /** This is the run's own conversation (or the one that started it). */
  own: boolean
  statusError?: string
  checkedAt?: number
  onOpen: (name: string) => void
  /** Brings the result card, with every drafted file, into view. */
  onShowDrafts: () => void
}) {
  const [now, setNow] = useState(() => Date.now())
  const [details, setDetails] = useState(false)
  const detailsId = useId()
  // The run state the lawyer was asked about: any change to the run withdraws the question.
  const [confirming, setConfirming] = useState<string>()
  const confirmId = useId()
  const refocusStop = useRef(false)

  useEffect(() => {
    const timer = window.setInterval(() => setNow(Date.now()), status.in_flight ? 1000 : 30_000)
    return () => window.clearInterval(timer)
  }, [status.in_flight])

  const head = runHeadline(status, now, actions.resuming)
  // A review draft delivered with unresolved checklist findings is not a clean completion.
  const findings = head.tone === "done" && checklistFindingsCount(status) !== undefined
  const stopped = leanRunStopped(status)
  const asks = openQuestions(status).length
  const drafts = draftedGroups(status).map((group) => group.main).filter((output) => output.role === "draft")

  // Elsewhere the run is one line; questions still owed lead it, since they are
  // answered in the run's own conversation.
  if (!own)
    return (
      <section className={`run-head is-compact is-${head.tone}`} aria-label="Drafting run">
        <RunIcon tone={head.tone} stopped={stopped} findings={findings} />
        <div className="run-head-text">
          <p className="run-head-title" role="status" aria-live="polite">
            {head.title}
          </p>
          <p className="run-head-meta">
            {asks > 0 && <span className="run-head-asks">{asks === 1 ? "1 question for you" : `${asks} questions for you`}</span>}
            <span>{asks > 0 ? "Answer them in the run's own conversation" : stopped ? "A new request starts a new draft" : "Running in its own conversation"}</span>
          </p>
        </div>
        <Link className="run-head-cta button" to={{ search: `?view=chat&session=${encodeURIComponent(status.owner_session_id)}` }}>
          Open run
        </Link>
      </section>
    )

  const titles = new Map((status.plan?.documents ?? []).map((row) => [row.id, cleanTitle(row.title)]))
  const attention = attentionItems(status)
  // Once delivered, what is left to look at; before then, how the final check
  // settled the open items once it ran, else the run's own count.
  const settled = leanSettledItems(status)
  const note = status.status === "completed" || !settled ? openItemsSentence(status) : settled.length > 0 ? leanSettledSummary(settled, status.memo) : undefined
  const verdict = finalCheckNote(status)
  // The combined file holds the memo and a draft already counted, so it is not counted again.
  const held = status.documents?.status === "blocked"
    ? status.documents.documents.filter((document) => document.kind !== "drafting_notes_markdown" && document.kind !== "combined").length
    : 0
  const canContinue = leanContinueAvailable(status) && !actions.resuming
  const canPause = runState(status) === "running" && !actions.resuming
  const canStop = leanStopAvailable(status) && !actions.resuming
  const shown = `${status.owner_session_id}@${status.revision}`
  const confirmStop = canStop && confirming === shown
  const openDraft = () => {
    if (drafts.length === 1) return onOpen(drafts[0]!.name)
    onShowDrafts()
  }
  const draftLabel = drafts.length === 1 ? "Open draft" : "View drafts"
  const meta = [
    ...head.meta,
    asks > 0 && head.tone !== "input" ? <span className="run-head-asks">{asks === 1 ? "1 question for you" : `${asks} questions for you`}</span> : null,
    status.pending_requests > 0 && !stopped ? "Request queued" : null,
    statusError ? `Reconnecting… as of ${checkedAt ? formatAgo(checkedAt, now) : "earlier"}` : null,
  ].filter((part) => part !== null && part !== "")

  return (
    <section className={`run-head is-${head.tone}`} aria-label="Drafting run">
      <div className="run-head-main">
        <RunIcon tone={head.tone} stopped={stopped} findings={findings} />
        <div className="run-head-text">
          <h2 className="run-head-title" aria-live="polite">
            {head.title}
          </h2>
          {meta.length > 0 && (
            <p className="run-head-meta">
              {meta.map((part, index) => (
                <span key={index}>{part}</span>
              ))}
            </p>
          )}
        </div>
        <div className="run-head-actions">
          {canPause && (
            <button type="button" className="run-head-cta" disabled={actions.busy} onClick={() => actions.pause()}>
              {actions.pausing ? "Pausing…" : "Pause drafting"}
            </button>
          )}
          {canStop && !confirmStop && (
            <button
              type="button"
              className="run-head-cta"
              disabled={actions.busy}
              ref={(button) => {
                if (button && refocusStop.current) button.focus()
                refocusStop.current = false
              }}
              onClick={() => setConfirming(shown)}
            >
              Stop draft
            </button>
          )}
          {asks === 0 && canContinue && !confirmStop && (
            <button type="button" className="primary run-head-cta" disabled={actions.busy} onClick={() => actions.resume()}>
              Continue drafting
            </button>
          )}
          {head.tone === "done" && drafts.length > 0 && (
            <button type="button" className={`${asks > 0 ? "" : "primary "}run-head-cta`} onClick={openDraft}>
              {draftLabel}
            </button>
          )}
        </div>
      </div>
      {confirmStop && (
        <div
          className="run-head-confirm"
          role="group"
          aria-labelledby={confirmId}
          onKeyDown={(event) => {
            if (event.key !== "Escape" || actions.busy) return
            refocusStop.current = true
            setConfirming(undefined)
          }}
        >
          <p id={confirmId}>
            Stop this draft? Drafting ends here and cannot be continued. Its work so far stays in the matter, and your next
            drafting request starts a new draft.
          </p>
          <div className="run-head-confirm-actions">
            <button
              type="button"
              autoFocus
              disabled={actions.busy}
              onClick={() => {
                refocusStop.current = true
                setConfirming(undefined)
              }}
            >
              Keep draft
            </button>
            <button type="button" className="danger" disabled={actions.busy} onClick={() => actions.stop()}>
              {actions.stopping ? "Stopping…" : "Stop draft"}
            </button>
          </div>
        </div>
      )}
      {actions.error && (
        <div className="msg-error" role="alert">
          {actions.error}
        </div>
      )}
      <ChecklistRefinement key={status.owner_session_id} status={status} actions={actions} />
      <div className="run-head-progress">
        <Segments status={status} />
        <button
          type="button"
          className="run-head-toggle"
          aria-expanded={details}
          aria-controls={detailsId}
          onClick={() => setDetails(!details)}
        >
          Details
        </button>
      </div>
      <div id={detailsId} className={`run-head-details${details ? " is-open" : ""}`} inert={!details} aria-hidden={!details}>
        <div className="run-head-details-body">
          <Stepper status={status} now={now} />
          <p className="run-note">{status.stages["research"] ? status.stages["research"].status === "done" ? status.research?.status === "not-needed" ? `External authority was not needed for this request; the reason is recorded with ${notesWording(status.memo).phrase}.` : status.research?.status === "limited" ? "Legal research is incomplete. Source and research limitations must be disclosed with this review draft." : "Retrieved legal research findings are independently checked before drafting." : "External legal research runs after material questions are answered. Unavailable sources and research gaps are disclosed." : "This earlier drafting run checked matter documents; external legal research was not included."}</p>
          {held > 0 && !stopped && <p className="run-note">{held} prepared document{held === 1 ? " is" : "s are"} held until the final check passes.</p>}
          {note && <p className="run-note">{note}</p>}
          {head.tone === "done" && verdict && <p className="run-note">{verdict}</p>}
          <OpenItems items={attention.filter((item) => item.kind !== "question")} titles={titles} />
        </div>
      </div>
    </section>
  )
}

/** The header's state, title and meta, from the run's status alone. Questions
 * the lawyer owes lead whenever the run is not working, except once it has
 * delivered: the draft is then ready and the questions refine it. */
export function runHeadline(status: LeanDraftingStatus, now: number, resuming = false): { tone: RunState; title: string; meta: string[] } {
  const state = runState(status)
  const position = runPosition(status)
  const step = currentStep(status)
  const steps = step ? `Step ${step.position} of ${step.total}` : ""
  if (resuming) return { tone: "running", title: "Resuming…", meta: [position?.document ?? "", steps] }
  if (state === "done") {
    const took = draftedIn(status)
    const drafts = draftedGroups(status).filter((group) => group.main.role === "draft").length
    const findings = checklistFindingsCount(status)
    return {
      tone: "done",
      title: findings === undefined ? "Ready for review" : findingsTitle(findings),
      meta: [
        took ? `Drafted in ${took}` : "",
        drafts ? `${drafts} document${drafts === 1 ? "" : "s"}` : "",
        afterCheckChange(status)?.label ?? "",
      ],
    }
  }
  if (state === "running") {
    const key = activeStep(status) ?? step?.key
    const row = key ? status.stages[key] : undefined
    const time = row ? stageElapsed(row, now) : undefined
    return {
      tone: "running",
      title: key ? stepActivityLabel(status, key) : "Drafting",
      meta: [
        position?.document ?? "",
        steps,
        time !== undefined ? formatElapsed(time) : "",
        time !== undefined && time > LONG_STEP_MS ? "long steps are normal" : "",
      ],
    }
  }
  if (leanRunStopped(status))
    return {
      tone: "paused",
      title: "Draft stopped",
      meta: ["You stopped this draft", "Its work so far is kept", "Send a new request to start a new draft"],
    }
  if (state === "input") {
    const questions = openQuestions(status)
    const asks = questions.length
    // Where Decide for me is offered, leaving a question to the drafter also lets the run resume.
    const delegable = questions.filter(delegableQuestion).length
    const resumes = !delegable
      ? "Drafting resumes after all answers are received"
      : delegable < asks
        ? "Drafting resumes once each is answered or, where Decide for me is offered, left to the drafter"
        : `Drafting resumes once ${asks === 1 ? "it is" : "each is"} answered or left to the drafter`
    return {
      tone: "input",
      title: "Waiting for your answers",
      meta: [`${asks === 1 ? "1 question" : `${asks} questions`} remaining · ${resumes}`],
    }
  }
  if (status.status === "running")
    return { tone: "paused", title: "Paused", meta: ["The drafting service is not working on this run", steps] }
  if (status.paused_at)
    return {
      tone: "paused",
      title: step ? `Paused at ${runStepName(status, step.key).toLowerCase()}` : "Paused",
      meta: [position?.document ?? "", "You paused drafting", "Continue picks up from this step"],
    }
  if (stoppedByTechnicalProblem(status))
    return { tone: "paused", title: "Paused after a technical problem", meta: ["The work so far is kept", steps] }
  if (step?.key === "gate") {
    const blocking = status.gate?.status === "blocked" ? status.gate.blocking : []
    const first = groupOpenItems(blocking)[0]
    return {
      tone: "paused",
      title: "Paused at the final check",
      meta: [
        first ? `${blocking.length} item${blocking.length === 1 ? "" : "s"} to resolve, starting with ${first.label.toLowerCase()}` : "",
        "Continue runs a repair round",
      ],
    }
  }
  // The checker's own wording stays on the step's row under Details.
  const failed = step ? stepCause(status, step.key) !== undefined : false
  return {
    tone: "paused",
    title: step ? `Paused at ${runStepName(status, step.key).toLowerCase()}` : "Paused",
    meta: [position?.document ?? "", failed ? "Its check did not pass" : steps, failed ? "Continue runs the step again" : ""],
  }
}

/** One segment per step, grouped per document; the step in progress pulses. */
function Segments({ status }: { status: LeanDraftingStatus }) {
  const current = currentStep(status)
  const groups = stepGroups(status)
  const done = Object.values(status.stages).filter((row) => row.status === "done").length
  const total = Object.keys(status.stages).length
  return (
    <div className="run-segments" role="img" aria-label={`${done} of ${total} steps done`}>
      {groups.map((group) => (
        <span key={group.id} className="run-segment-group" title={group.title}>
          {group.steps.map((step) => (
            <span
              key={step.key}
              className={`run-segment is-${segmentState(status, step.key, current?.key)}`}
              title={`${group.title} · ${step.name}`}
            />
          ))}
        </span>
      ))}
    </div>
  )
}

function segmentState(status: LeanDraftingStatus, key: string, current: string | undefined) {
  const row = status.stages[key]!
  if (row.status === "done") return "done"
  if (key !== current) return "pending"
  return status.in_flight && row.status === "running" ? "live" : "stopped"
}

function RunIcon({ tone, stopped = false, findings = false }: { tone: RunState; stopped?: boolean; findings?: boolean }) {
  return (
    <span className={`run-head-icon is-${tone}${findings ? " is-findings" : ""}`} aria-hidden>
      {tone === "done" && findings && (
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="3" strokeLinecap="round">
          <line x1="12" y1="6" x2="12" y2="13" />
          <line x1="12" y1="18" x2="12.01" y2="18" />
        </svg>
      )}
      {tone === "done" && !findings && (
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="3" strokeLinecap="round" strokeLinejoin="round">
          <polyline points="20 6 9 17 4 12" />
        </svg>
      )}
      {tone === "input" && (
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.6" strokeLinecap="round" strokeLinejoin="round">
          <path d="M9.1 9a3 3 0 0 1 5.8 1c0 2-3 3-3 3" />
          <line x1="12" y1="17" x2="12.01" y2="17" />
        </svg>
      )}
      {tone === "paused" && stopped && (
        <svg width="14" height="14" viewBox="0 0 24 24" fill="currentColor">
          <rect x="7" y="7" width="10" height="10" rx="1.5" />
        </svg>
      )}
      {tone === "paused" && !stopped && (
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="3" strokeLinecap="round">
          <line x1="10" y1="7" x2="10" y2="17" />
          <line x1="15" y1="7" x2="15" y2="17" />
        </svg>
      )}
    </span>
  )
}
