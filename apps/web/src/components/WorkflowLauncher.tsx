import { useId, useState } from "react"
import type { WorkflowMeta } from "../agents"
import { recommendStarter, REVIEW_STARTERS } from "../review-starters"
import ModalHeader from "./ModalHeader"
import StarterSummary from "./StarterSummary"
import backdropClose from "./backdropClose"
import { Exit } from "./useExit"
import "./WorkflowLauncher.css"

type Props = {
  workflows: WorkflowMeta[]
  available: Set<string>
  documents?: readonly { role?: "source" | "deliverable" }[]
  onLaunch: (name: string) => void
  // Present in a matter chat: review-table starters are offered too, and Start
  // opens the matter's table with the starter chosen.
  onStartReview?: (starter: string) => void
}

// "Workflows" control by the composer. Opens a list of the multi-step routines
// available for this matter, and the review-table starters, beside a preview of
// the highlighted one. Only Start runs anything; closing leaves the composer as
// it was.
export default function WorkflowLauncher({ workflows, available, ...props }: Props) {
  const options = workflows.filter((w) => available.has(w.name))
  const [open, setOpen] = useState(false)
  if (options.length === 0 && !props.onStartReview) return null
  return (<>
    <button className="workflow-btn" onClick={() => setOpen(true)} title={props.onStartReview ? "Run a workflow or start a review table" : "Run a multi-step workflow"}>
      <svg className="workflow-icon" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
        <rect x="3" y="3" width="6" height="6" rx="1" /><rect x="15" y="15" width="6" height="6" rx="1" /><path d="M9 6h6a2 2 0 0 1 2 2v7" />
      </svg>
      <span>Workflows</span>
    </button>
    <Exit when={open}>{(_, closing) => <Launcher {...props} options={options} closing={closing} onClose={() => setOpen(false)} />}</Exit>
  </>)
}

function Launcher({ options, documents = [], onLaunch, onStartReview, closing, onClose }: Omit<Props, "workflows" | "available"> & { options: WorkflowMeta[]; closing: boolean; onClose: () => void }) {
  const ids = useId()
  const [recommended] = useState(() => recommendStarter(documents))
  const starters = onStartReview ? [...REVIEW_STARTERS].sort((a, b) => Number(b.id === recommended.id) - Number(a.id === recommended.id)) : []
  // Several of the lawyer's own documents make the recommended starter the
  // likely next step, so it leads and is highlighted; otherwise chat leads.
  const startersFirst = starters.length > 0 && (options.length === 0 || documents.filter((doc) => doc.role !== "deliverable").length >= 2)
  const [selected, setSelected] = useState(() => startersFirst ? `starter:${starters[0].id}` : `workflow:${options[0].name}`)
  const workflow = options.find((w) => selected === `workflow:${w.name}`)
  const starter = starters.find((s) => selected === `starter:${s.id}`)

  function start() {
    onClose()
    if (starter) onStartReview?.(starter.id)
    else if (workflow) onLaunch(workflow.name)
  }

  const choice = (key: string, name: string, reason?: string) => (
    <label key={key} className={`workflow-launcher-choice${selected === key ? " is-selected" : ""}`}>
      <input type="radio" name={`${ids}-choice`} value={key} checked={selected === key} autoFocus={selected === key} onChange={() => setSelected(key)} />
      <span>
        <span className="workflow-launcher-choice-name">{name}</span>
        {reason && <span className="workflow-launcher-choice-reason"><strong>Recommended</strong> · {reason}</span>}
      </span>
    </label>
  )
  const groups = [
    options.length > 0 && (
      <fieldset key="chat" className="workflow-launcher-group">
        <legend>In chat</legend>
        {options.map((w) => choice(`workflow:${w.name}`, w.label))}
      </fieldset>
    ),
    starters.length > 0 && (
      <fieldset key="starters" className="workflow-launcher-group">
        <legend>Review table</legend>
        {/* Recommended only where it leads; beside a preselected chat workflow it would compete. */}
        {starters.map((s) => choice(`starter:${s.id}`, s.name, startersFirst && s.id === recommended.id ? recommended.reason : undefined))}
      </fieldset>
    ),
  ]

  return (
    <div className={`viewer-overlay${closing ? " is-closing" : ""}`} inert={closing} {...backdropClose(onClose)}>
      <div className="picker-panel workflow-launcher" onClick={(e) => e.stopPropagation()}>
        <ModalHeader title="Workflows" onClose={onClose} />
        <div className="workflow-launcher-body">
          <div className="workflow-launcher-list">{startersFirst ? groups.reverse() : groups}</div>
          <section className="workflow-launcher-preview" aria-label={`About ${starter?.name ?? workflow?.label}`}>
            {starter ? (
              <>
                <StarterSummary starter={starter} />
                <p className="workflow-launcher-note">Start opens the review table at this starter&apos;s {starter.questions.length} questions, which you check before they are added. No answers are filled until you choose Fill.</p>
              </>
            ) : workflow && (
              <>
                <p>{workflow.description}</p>
                {workflow.steps?.length ? (
                  <section aria-label="Steps">
                    <h3>Steps</h3>
                    <ol>{workflow.steps.map((step, index) => <li key={index}>{step.instructions}</li>)}</ol>
                  </section>
                ) : null}
                <p className="workflow-launcher-note">
                  {workflow.scope === "matter"
                    ? "Start runs this workflow with AI across every document in this matter, as a new conversation."
                    : "Start runs this workflow with AI in this conversation."}
                </p>
              </>
            )}
          </section>
        </div>
        <div className="workflow-launcher-actions compare-actions">
          <button type="button" className="ghost" onClick={onClose}>Cancel</button>
          <button type="button" className="primary" onClick={start}>Start</button>
        </div>
      </div>
    </div>
  )
}
