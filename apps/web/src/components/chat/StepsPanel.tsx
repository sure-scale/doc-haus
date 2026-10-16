import { useEffect, useState } from "react"
import Markdown from "../markdown/Markdown"
import { ToolOutput } from "../markdown/tool-output"
import type { Step } from "./turn-model"

// The reasoning timeline: a collapsible panel of thinking blocks and tool calls,
// each on a dotted timeline. It is the focus only while the model is still
// working with nothing to read yet; the instant the answer streams (or the turn
// settles) it collapses to a one-line affordance so the prose owns the view —
// the handoff from live work to the settled answer. A manual toggle
// still wins until the next phase change.
export default function StepsPanel({ steps, busy, answered }: { steps: Step[]; busy: boolean; answered: boolean }) {
  const thinking = busy && !answered
  const [open, setOpen] = useState(thinking)
  useEffect(() => setOpen(thinking), [thinking])
  if (steps.length === 0) return null
  return (
    <details className="steps" open={open} onToggle={(e) => setOpen(e.currentTarget.open)}>
      <summary>{thinking ? "Working..." : `${steps.length} step${steps.length === 1 ? "" : "s"}`}</summary>
      <StepList steps={steps} busy={busy} />
    </details>
  )
}

// One level of the timeline. Each tool step that spawned a subagent nests that
// subagent's own steps as a deeper list, so a consulted reviewer's thinking and
// tool calls stream live beneath "Consulted the Reviewer" rather than the step
// sitting opaque while the child works. Recurses to any depth.
function StepList({ steps, busy }: { steps: Step[]; busy: boolean }) {
  // When the list is still running but its last step has settled (tool finished,
  // thinking block closed), the next step hasn't arrived yet — show a pending
  // row at the bottom so the timeline never looks stalled between steps.
  const lastStep = steps[steps.length - 1]
  const settled = lastStep && (lastStep.kind === "tool" ? lastStep.status !== "running" : lastStep.done)
  return (
    <ol className="step-list">
      {steps.map((s, i) => (
        <StepRow key={i} step={s} busy={busy} last={i === steps.length - 1} />
      ))}
      {busy && settled && (
        <li className="step step-pending">
          <span className="dot running" />
          <span className="sk-line" />
        </li>
      )}
    </ol>
  )
}

function StepRow({ step, busy, last }: { step: Step; busy: boolean; last: boolean }) {
  // The last step in a still-running list is the one in flight: a tool not yet
  // done, or a thinking block still streaming. That one stays open and live.
  const active = busy && last
  if (step.kind === "tool")
    return (
      <li className="step">
        <span className={`dot ${step.status}`} />
        {step.detail ? (
          <details className="step-detail">
            <summary className="step-label">{step.label}</summary>
            {step.templateResult ? (
              <div className="step-template-result">
                <p>{step.templateResult.summary}</p>
                {step.templateResult.placeholders.length > 0 && (
                  <details className="step-template-fields">
                    <summary>{step.templateResult.placeholders.length} placeholders</summary>
                    <ul>
                      {step.templateResult.placeholders.map((placeholder, index) => (
                        <li key={index}>{placeholder}</li>
                      ))}
                    </ul>
                  </details>
                )}
              </div>
            ) : (
              <ToolOutput>{step.detail}</ToolOutput>
            )}
          </details>
        ) : (
          <span className="step-label">{step.label}</span>
        )}
        {step.children && step.children.length > 0 && (
          <StepList steps={step.children} busy={busy && step.status === "running"} />
        )}
      </li>
    )
  // The block still streaming reads as "Analyzing..." and stays open; settled
  // blocks collapse to "Thought process" so the timeline is calm.
  const live = active && !step.done
  return (
    <li className="step">
      <span className={`dot reasoning${live ? " running" : ""}`} />
      <details className="thinking" open={live}>
        <summary>{live ? "Analyzing..." : "Thought process"}</summary>
        <Markdown>{step.text}</Markdown>
      </details>
    </li>
  )
}
