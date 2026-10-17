import type { LeanDraftingStatus } from "../../api/ingest/lean"
import type { MatterDetail as Detail } from "../../api/ingest/matters"
import TelemetryPanel from "../TelemetryPanel"
import { checklistFindingsCount, findingsTitle } from "../drafting/model"

// The matter's status bar: drafting state on the left, usage on the right.
export default function MatterRail({
  matter,
  session,
  drafting,
  showStatusbar = true,
}: {
  matter: Detail
  session?: string
  drafting?: LeanDraftingStatus
  showStatusbar?: boolean
}) {
  const label = drafting && draftingStatusLabel(drafting)

  return (
    <>
      {showStatusbar && <footer className="matter-statusbar" aria-label="Matter activity and usage">
        <span className={`matter-statusbar-label${drafting ? ` is-${drafting.status}` : ""}${label?.findings ? " is-findings" : ""}`}>
          {drafting && <span className="matter-status-dot" aria-hidden />}
          {label?.text || "Ready"}
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
