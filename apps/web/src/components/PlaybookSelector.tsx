import { useState } from "react"
import { Link } from "react-router-dom"

// "Playbook" control by the composer. Binds one firm playbook to this matter:
// the review checks documents against its positions and approved clause text.
// Opens a modal of the firm playbooks available; picking one binds it (or
// unbinds via "No playbook"). The trigger shows the bound playbook so the
// binding reads at a glance.
export default function PlaybookSelector({
  playbooks,
  value,
  onChange,
}: {
  playbooks: { name: string; description: string }[]
  value?: string
  onChange: (name?: string) => void
}) {
  const [open, setOpen] = useState(false)

  // The bound playbook's readable name (the "playbook-" prefix stripped and
  // upper-cased) — surfaced on the trigger so the binding is visible at a glance.
  const boundLabel = value ? value.replace(/^playbook-/, "").toUpperCase() : undefined
  return (
    <>
      <button
        className="workflow-btn"
        onClick={() => setOpen(true)}
        title="Firm playbook for this matter — review checks clauses against its positions"
      >
        <svg className="workflow-icon" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
          <path d="M4 19.5A2.5 2.5 0 0 1 6.5 17H20" />
          <path d="M6.5 2H20v20H6.5A2.5 2.5 0 0 1 4 19.5v-15A2.5 2.5 0 0 1 6.5 2z" />
        </svg>
        <span>{boundLabel ? `Playbook: ${boundLabel}` : "Playbook"}</span>
      </button>
      {open && (
        <div className="viewer-overlay" onClick={() => setOpen(false)}>
          <div className="picker-panel picker-wide" onClick={(e) => e.stopPropagation()}>
            <div className="viewer-bar">
              <span className="viewer-title">Playbook</span>
              <button className="modal-close" aria-label="Close" onClick={() => setOpen(false)}>
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" aria-hidden="true"><path d="M18 6 6 18M6 6l12 12" /></svg>
              </button>
            </div>
            <div className="picker-body">
              <p className="muted" style={{ margin: "0 0 4px" }}>
                A firm playbook is your standard positions and approved clause text. The review checks this matter's
                documents against it.
              </p>
              {playbooks.length === 0 && <p className="muted">No playbooks yet. Import or create one in the library.</p>}
              <Link to="/playbooks" className="assistant-option" onClick={() => setOpen(false)}>
                <span className="assistant-name">Manage playbooks</span>
                <span className="assistant-desc muted">View, import and edit your firm playbooks</span>
              </Link>
              {playbooks.map((p) => (
                <button
                  key={p.name}
                  className={`assistant-option${p.name === value ? " selected" : ""}`}
                  onClick={() => {
                    onChange(p.name)
                    setOpen(false)
                  }}
                >
                  <span className="assistant-name">{p.name.replace(/^playbook-/, "").toUpperCase()}</span>
                  <span className="assistant-desc muted">{p.description}</span>
                </button>
              ))}
              <button
                className={`assistant-option${value ? "" : " selected"}`}
                onClick={() => {
                  onChange(undefined)
                  setOpen(false)
                }}
              >
                <span className="assistant-name">No playbook</span>
                <span className="assistant-desc muted">Skip playbook checks for this matter</span>
              </button>
            </div>
          </div>
        </div>
      )}
    </>
  )
}
