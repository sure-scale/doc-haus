import type { RedlineProposal } from "./turn-model"
import { basename } from "./step-labels"

// In-chat preview of the redlines a turn proposed: each shows the prior wording
// struck red above the new wording in green — the same red/green the document
// viewer paints — so the lawyer sees the change without leaving the chat, with a
// link straight into the full tracked-changes review for that proposal.
export default function RedlineView({
  redlines,
  onView,
}: {
  redlines: RedlineProposal[]
  onView: (name: string, redlineId?: number) => void
}) {
  if (redlines.length === 0) return null
  return (
    <div className="redlines">
      {redlines.map((r, i) => (
        <div className="redline-card" key={i}>
          <div className="redline-doc">{basename(r.document)}</div>
          {r.oldText && <div className="redline-old">{r.oldText}</div>}
          <div className="redline-new">{r.newText}</div>
          {r.comment && <div className="redline-comment">{r.comment}</div>}
          <button className="linklike redline-view" onClick={() => onView(basename(r.document), r.id)}>
            View in document
          </button>
        </div>
      ))}
    </div>
  )
}
