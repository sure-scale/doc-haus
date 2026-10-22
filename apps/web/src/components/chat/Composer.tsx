import { useId, type ReactNode } from "react"
import { isAuto, type AssistantMeta, type WorkflowMeta } from "../../agents"
import ModelSelector from "../ModelSelector"
import PlaybookSelector from "../PlaybookSelector"
import WorkflowLauncher from "../WorkflowLauncher"
import { agentLabel } from "./step-labels"
import type { ChatState } from "./use-chat-state"

export default function Composer({
  pinned,
  available,
  agent,
  resolvedAgent,
  setResolvedAgent,
  onAgentChange,
  inputRef,
  workflows,
  runWorkflow,
  playbooks,
  playbook,
  onPlaybookChange,
  toolsLoading,
  composerPlaceholder,
  input,
  setInput,
  onSend,
  onStop,
  stopping = false,
  busy,
  draftingUpdate = false,
  draftingClarification = false,
  documentsControl,
}: {
  pinned?: AssistantMeta
  available: Set<string>
  agent: string
  resolvedAgent?: string
  setResolvedAgent: ChatState["setResolvedAgent"]
  onAgentChange: (name: string) => void
  inputRef: ChatState["inputRef"]
  workflows: WorkflowMeta[]
  runWorkflow: (name: string) => Promise<void>
  playbooks?: { name: string; description: string }[]
  playbook?: string
  onPlaybookChange?: (name?: string) => void
  toolsLoading?: boolean
  composerPlaceholder: string
  input: string
  setInput: ChatState["setInput"]
  onSend: () => Promise<void>
  // Present while a chat answer runs: Stop takes Send's place.
  onStop?: () => void
  stopping?: boolean
  busy: boolean
  draftingUpdate?: boolean
  draftingClarification?: boolean
  // The matter's Documents control (see DocumentChips), seated after the playbook.
  documentsControl?: ReactNode
}) {
  const hintId = useId()
  // A pinned surface (the template library) owns its one assistant — no
  // picker, no workflow launcher, nothing to choose. Drop the whole tools row so
  // the composer reads clean. Otherwise the tools sit in the composer's footer,
  // beside Send, so the field and its controls read as one form.
  const tools =
    !pinned &&
    !draftingUpdate &&
    (toolsLoading ? (
      <div className="composer-tools" aria-busy="true" aria-label="Loading assistant choices">
        <span className="composer-tool-skeleton is-assistant" />
        <span className="composer-tool-skeleton is-workflows" />
        {onPlaybookChange && <span className="composer-tool-skeleton is-playbook" />}
      </div>
    ) : (
      <div className="composer-tools">
        <ModelSelector
          available={available}
          value={agent}
          resolvedLabel={isAuto(agent) && resolvedAgent ? agentLabel(resolvedAgent) : undefined}
          onChange={(a) => {
            onAgentChange(a)
            // Leaving Auto drops the last routing so the chip does not keep showing
            // a resolved label under a manually-picked assistant.
            if (!isAuto(a)) setResolvedAgent(undefined)
            inputRef.current?.focus()
          }}
        />
        <WorkflowLauncher available={available} onLaunch={runWorkflow} workflows={workflows} />
        {onPlaybookChange && (
          <PlaybookSelector playbooks={playbooks ?? []} value={playbook} onChange={onPlaybookChange} />
        )}
        {documentsControl}
      </div>
    ))
  return (
    <div className="composer">
      <textarea
        ref={inputRef}
        rows={1}
        aria-label="Message"
        aria-describedby={hintId}
        placeholder={composerPlaceholder}
        value={input}
        onChange={(e) => {
          setInput(e.target.value)
          const el = e.currentTarget
          el.style.height = "auto"
          el.style.height = Math.min(el.scrollHeight, 200) + "px"
        }}
        onKeyDown={(e) => {
          if (e.nativeEvent.isComposing || e.keyCode === 229) return
          if (e.key === "Enter" && (e.metaKey || e.ctrlKey)) {
            e.preventDefault()
            onSend()
            return
          }
          if (e.key === "Enter" && !e.shiftKey) {
            e.preventDefault()
            onSend()
          }
        }}
      />
      <div className="composer-footer">
        {tools}
        <span id={hintId} className="sr-only">Enter to send. Shift and Enter for a new line.</span>
        <span className="composer-hint" aria-hidden>
          <span>
            <kbd>Enter</kbd> send
          </span>
          <span className="composer-hint-sep">·</span>
          <span>
            <kbd>Shift + Enter</kbd> new line
          </span>
        </span>
        {onStop ? (
          <button
            type="button"
            className="send-btn stop-btn btn-lg"
            onClick={onStop}
            disabled={stopping}
            aria-label={stopping ? "Stopping answer" : "Stop answer"}
          >
            <span>{stopping ? "Stopping…" : "Stop"}</span>
            <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true">
              <rect x="6" y="6" width="12" height="12" rx="2" />
            </svg>
          </button>
        ) : (
          <button
            type="button"
            className="send-btn btn-lg"
            onClick={onSend}
            disabled={busy || !input.trim()}
            aria-label={
              draftingClarification
                ? "Send checklist clarification"
                : draftingUpdate
                  ? "Send drafting update"
                  : "Send message"
            }
          >
            <span>{draftingClarification ? "Clarify checklist" : draftingUpdate ? "Update draft" : "Send"}</span>
            <svg
              width="16"
              height="16"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeLinecap="round"
              strokeLinejoin="round"
              aria-hidden="true"
            >
              <path d="M12 19V5" />
              <path d="m5 12 7-7 7 7" />
            </svg>
          </button>
        )}
      </div>
    </div>
  )
}
