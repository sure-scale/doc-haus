import type { LeanDraftingStatus } from "../../api/ingest/lean"
import type { MatterDetail as Detail } from "../../api/ingest/matters"
import TelemetryPanel from "../TelemetryPanel"

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
  const status = drafting && {
    running: drafting.in_flight ? "Drafting in progress" : "Drafting paused",
    blocked: "Drafting needs your attention",
    failed: "Drafting interrupted",
    completed: "Drafting complete",
  }[drafting.status]

  return (
    <>
      {showStatusbar && <footer className="matter-statusbar" aria-label="Matter activity and usage">
        <span className={`matter-statusbar-label${drafting ? ` is-${drafting.status}` : ""}`}>
          {drafting && <span className="matter-status-dot" aria-hidden />}
          {status || "Ready"}
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
