import type { LeanDraftingStatus } from "../../api/ingest/lean"
import type { MatterDetail as Detail } from "../../api/ingest/matters"
import TelemetryPanel from "../TelemetryPanel"
import { checklistFindingsCount, findingsTitle } from "../drafting/model"

// The matter's status bar: a chat answer or drafting state on the left, usage on the right.
export default function MatterRail({
  matter,
  session,
  drafting,
  chat,
  showStatusbar = true,
}: {
  matter: Detail
  session?: string
  drafting?: LeanDraftingStatus
  // A chat answer in progress, and what it is doing now when that has a plain name.
  chat?: { activity?: string }
  showStatusbar?: boolean
}) {
  const label = statusbarLabel(drafting, chat)

  return (
    <>
      {showStatusbar && <footer className="matter-statusbar" aria-label="Matter activity and usage">
        <span className={`matter-statusbar-label${label.status ? ` is-${label.status}` : ""}${label.findings ? " is-findings" : ""}`}>
          {label.status && <span className="matter-status-dot" aria-hidden />}
          {label.text}
        </span>
        <div className="matter-statusbar-usage">
          <TelemetryPanel
            matterId={matter.id}
            sessionID={session}
          />
        </div>
      </footer>}
    </>
  )
}

/** The status bar's label. A chat answer in progress is what the matter is
 * doing while it runs; otherwise the drafting run's label, or Ready. */
export function statusbarLabel(drafting?: LeanDraftingStatus, chat?: { activity?: string }) {
  if (chat)
    return {
      text: chat.activity ? `Answer in progress · ${chat.activity}` : "Answer in progress",
      status: "running",
      findings: false,
    }
  if (drafting) return { ...draftingStatusLabel(drafting), status: drafting.status }
  return { text: "Ready", findings: false }
}

/** The status bar's drafting label. A review draft delivered with unresolved
 * checklist findings is not a clean completion, so it is named and flagged. */
export function draftingStatusLabel(drafting: LeanDraftingStatus) {
  const findings = drafting.status === "completed" ? checklistFindingsCount(drafting) : undefined
  const text = {
    running: drafting.in_flight ? "Drafting in progress" : "Drafting paused",
    blocked: "Drafting needs your attention",
    failed: "Drafting interrupted",
    completed: findings === undefined ? "Drafting complete" : findingsTitle(findings),
  }[drafting.status]
  return { text, findings: findings !== undefined }
}
