import { useState } from "react"
import { AUTO_ASSISTANT, CHAT_ASSISTANTS, isAuto } from "../agents"
import backdropClose from "./backdropClose"
import { Exit } from "./useExit"

// Picks who answers in the chat — a stateful mode selector (Auto, Q&A, Redline,
// Research), like a model picker. The choice persists across messages and is
// reflected in the chip. "Auto" is a router mode, not a real agent: while it is
// selected, `resolvedLabel` carries which real assistant the last message routed
// to so the chip reads "Auto · Redline". Workflows are one-shot actions, not
// modes, so they live in their own launcher rather than this list. `available` is
// the set of assistant ids the engine actually has, so we never offer one missing.
export default function ModelSelector({
  available,
  value,
  resolvedLabel,
  onChange,
}: {
  available: Set<string>
  value: string
  resolvedLabel?: string
  onChange: (name: string) => void
}) {
  // Auto leads the list and is always offered — it is a router over the real
  // assistants, not an engine agent, so it does not depend on `available`.
  const assistants = [AUTO_ASSISTANT, ...CHAT_ASSISTANTS.filter((a) => available.has(a.name))]
  const [open, setOpen] = useState(false)
  if (assistants.length === 1) return null

  const current = assistants.find((a) => a.name === value) ?? assistants[0]
  const triggerLabel = isAuto(current.name) && resolvedLabel ? `${current.label} · ${resolvedLabel}` : current.label
  return (
    <>
      <button className="assistant-trigger" onClick={() => setOpen(true)} title="Choose assistant or workflow">
        <svg className="assistant-icon" width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
          <path d="M9.937 15.5A2 2 0 0 0 8.5 14.063l-6.135-1.582a.5.5 0 0 1 0-.962L8.5 9.936A2 2 0 0 0 9.937 8.5l1.582-6.135a.5.5 0 0 1 .963 0L14.063 8.5A2 2 0 0 0 15.5 9.937l6.135 1.581a.5.5 0 0 1 0 .964L15.5 14.063a2 2 0 0 0-1.437 1.437l-1.582 6.135a.5.5 0 0 1-.963 0z" />
          <path d="M20 3v4M22 5h-4M4 17v2M5 18H3" />
        </svg>
        <span className="assistant-trigger-label">{triggerLabel}</span>
        <svg className="assistant-caret" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
          <path d="m6 9 6 6 6-6" />
        </svg>
      </button>
      <Exit when={open}>{(_, closing) => (
        <div className={`viewer-overlay${closing ? " is-closing" : ""}`} inert={closing} {...backdropClose(() => setOpen(false))}>
          <div className="picker-panel picker-wide" onClick={(e) => e.stopPropagation()}>
            <div className="viewer-bar">
              <span className="viewer-title">Choose an assistant</span>
              <button className="modal-close" aria-label="Close" onClick={() => setOpen(false)}>
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" aria-hidden="true"><path d="M18 6 6 18M6 6l12 12" /></svg>
              </button>
            </div>
            <div className="picker-body">
              {assistants.map((a) => (
                <button
                  key={a.name}
                  className={`assistant-option${a.name === value ? " selected" : ""}`}
                  onClick={() => {
                    onChange(a.name)
                    setOpen(false)
                  }}
                >
                  <span className="assistant-name">{a.label}</span>
                  <span className="assistant-desc muted">{a.description}</span>
                </button>
              ))}
            </div>
          </div>
        </div>
      )}</Exit>
    </>
  )
}
