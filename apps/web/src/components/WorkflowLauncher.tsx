import { useState } from "react"
import type { WorkflowMeta } from "../agents"
import ModalHeader from "./ModalHeader"
import backdropClose from "./backdropClose"
import { Exit } from "./useExit"

// "Workflows" control by the composer. Opens a modal of the multi-step routines
// available for this matter; picking one launches it into the artifact panel.
export default function WorkflowLauncher({
  workflows,
  available,
  onLaunch,
}: {
  workflows: WorkflowMeta[]
  available: Set<string>
  onLaunch: (name: string) => void
}) {
  const options = workflows.filter((w) => available.has(w.name))
  const [open, setOpen] = useState(false)

  if (options.length === 0) return null

  return (
    <>
      <button className="workflow-btn" onClick={() => setOpen(true)} title="Run a multi-step workflow">
        <svg className="workflow-icon" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
          <rect x="3" y="3" width="6" height="6" rx="1" />
          <rect x="15" y="15" width="6" height="6" rx="1" />
          <path d="M9 6h6a2 2 0 0 1 2 2v7" />
        </svg>
        <span>Workflows</span>
      </button>
      <Exit when={open}>{(_, closing) => (
        <div className={`viewer-overlay${closing ? " is-closing" : ""}`} inert={closing} {...backdropClose(() => setOpen(false))}>
          <div className="picker-panel" onClick={(e) => e.stopPropagation()}>
            <ModalHeader title="Workflows" onClose={() => setOpen(false)} />
            <div className="picker-body">
              {options.map((w) => (
                <button
                  key={w.name}
                  className="assistant-option"
                  onClick={() => {
                    onLaunch(w.name)
                    setOpen(false)
                  }}
                >
                  <span className="assistant-name">{w.label}</span>
                  <span className="assistant-desc muted">{w.description}</span>
                </button>
              ))}
            </div>
          </div>
        </div>
      )}</Exit>
    </>
  )
}
