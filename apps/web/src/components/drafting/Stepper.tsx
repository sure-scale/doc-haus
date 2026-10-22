import { useRef } from "react"
import type { LeanDraftingStatus } from "../../api/ingest/lean"
import { formatElapsed } from "../../lib/format"
import { currentStep, stageElapsed, stepCause, stepGroups, type Step } from "./model"

const STATUS_TEXT: Record<Step["row"]["status"], string> = {
  pending: "not started",
  // A running step whose run is not in flight is waiting for Continue.
  running: "paused",
  done: "done",
  failed: "stopped",
  blocked: "waiting",
}

// The run's steps as a vertical stepper, one group per planned document. The
// step in progress carries the accent and a calm pulse; a step the run stopped
// at shows why on its own row. A step that finishes while the panel is open
// draws its check in; steps already done when it opened just show theirs.
export default function Stepper({ status, now }: { status: LeanDraftingStatus; now: number }) {
  const current = currentStep(status)
  const groups = stepGroups(status)
  const keys = Object.keys(status.stages)
  const seen = useRef<Set<string>>(undefined)
  seen.current ??= new Set(keys.filter((key) => status.stages[key]?.status === "done"))
  return (
    <div className="stepper">
      {groups.map((group) => (
        <section key={group.id} className="stepper-group" aria-label={group.title}>
          {groups.length > 1 && <h3 className="stepper-group-title">{group.title}</h3>}
          <ol className="stepper-list" role="list">
            {group.steps.map((step) => {
              const isCurrent = current?.key === step.key
              const live = step.row.status === "running" && status.in_flight
              const elapsed = formatElapsed(stageElapsed(step.row, live ? now : Date.parse(status.updated_at)))
              const cause = stepCause(status, step.key)
              return (
                <li
                  key={step.key}
                  className={`stepper-step is-${step.row.status}${isCurrent ? " is-current" : ""}${live ? " is-live" : ""}`}
                  aria-current={isCurrent ? "step" : undefined}
                  aria-label={`Step ${keys.indexOf(step.key) + 1} of ${keys.length}: ${step.name}, ${live ? "running" : STATUS_TEXT[step.row.status]}${cause ? `. ${cause}` : ""}`}
                >
                  <span
                    className={`step-icon is-${live || step.row.status !== "running" ? step.row.status : "paused"}${step.row.status === "done" && !seen.current?.has(step.key) ? " is-fresh" : ""}`}
                    aria-hidden
                  >
                    <StepGlyph state={step.row.status} />
                  </span>
                  <span className="step-text">
                    <span className="step-name">{step.name}</span>
                    {step.row.attempts > 1 && <span className="step-tries">try {step.row.attempts}</span>}
                    {elapsed && <span className="step-time">{elapsed}</span>}
                    {cause && <span className="step-cause">{cause}</span>}
                  </span>
                </li>
              )
            })}
          </ol>
        </section>
      ))}
    </div>
  )
}

function StepGlyph({ state }: { state: Step["row"]["status"] }) {
  if (state === "done")
    return (
      <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="3" strokeLinecap="round" strokeLinejoin="round">
        <polyline points="20 6 9 17 4 12" pathLength={1} />
      </svg>
    )
  if (state === "failed" || state === "blocked")
    return (
      <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="3" strokeLinecap="round">
        <line x1="10" y1="7" x2="10" y2="17" />
        <line x1="15" y1="7" x2="15" y2="17" />
      </svg>
    )
  return null
}
