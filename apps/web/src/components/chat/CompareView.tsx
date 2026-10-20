import type { CompareParams } from "../../api/ingest/compare"
import { CompareIcon } from "../CompareDialog"
import type { VersionCompare } from "./turn-model"

// In-chat card for each comparison a turn ran: both files, how many changes,
// and how many of them were made without tracked changes, with a link into the
// blackline view of the same comparison.
export default function CompareView({ compares = [], onOpen }: { compares?: VersionCompare[]; onOpen?: (params: CompareParams) => void }) {
  if (compares.length === 0) return null
  return (
    <div className="drafts">
      {compares.map(({ changes, untracked, ...compare }, i) => (
        <div className="draft-card compare-card" key={i}>
          <span className="draft-icon">
            <CompareIcon />
          </span>
          <span className="draft-doc">
            {`Compared ${compare.left} → ${compare.right}: ${changes} change${changes === 1 ? "" : "s"}, ${untracked} made without tracked changes`}
          </span>
          {onOpen && (
            <button className="linklike" onClick={() => onOpen(compareParams(compare))}>
              Open blackline
            </button>
          )}
        </div>
      ))}
    </div>
  )
}

// The comparison to open, without what the turn listed of it.
export function compareParams({ left, right, leftLabel, rightLabel, ignoreFormatting, detectMoves, headersFooters }: Omit<VersionCompare, "changes" | "untracked">): CompareParams {
  return { left, right, leftLabel, rightLabel, ignoreFormatting, detectMoves, headersFooters }
}
