import { useEffect, useRef, useState } from "react"
import type {
  Chronology,
  ChronologyDecision,
  ChronologyEvent,
  ChronologyEvidence,
  ChronologyProposal,
} from "../../api/ingest/chronology"
import { dateLabel, eventEvidence, eventQualifications, simpleAddition } from "./model"
export default function ProposalReview({
  chronology,
  busy,
  initialProposal,
  onDecision,
  onBatch,
  onDiscard,
  onClose,
  onEvidence,
  onAsk,
}: {
  chronology: Chronology
  busy: boolean
  initialProposal?: string
  onDecision: (proposal: ChronologyProposal, decision: ChronologyDecision) => Promise<boolean>
  onBatch: (ids: string[]) => Promise<boolean>
  onDiscard: () => Promise<boolean>
  onClose: () => void
  onEvidence: (evidence: ChronologyEvidence) => void
  onAsk: (event: ChronologyEvent, proposal: ChronologyProposal) => void
}) {
  const proposals = chronology.proposals.filter(
    (proposal) => proposal.status === "pending" || proposal.status === "deferred",
  )
  const [selected, setSelected] = useState(initialProposal ?? proposals[0]?.id)
  const [batch, setBatch] = useState(false)
  const [discard, setDiscard] = useState(false)
  const [batchIds, setBatchIds] = useState<string[]>([])
  const heading = useRef<HTMLHeadingElement>(null)
  useEffect(() => {
    heading.current?.focus({ preventScroll: true })
  }, [])
  useEffect(() => {
    if (!proposals.some((proposal) => proposal.id === selected))
      setSelected(proposals.find((proposal) => proposal.status === "pending")?.id ?? proposals[0]?.id)
  }, [chronology.revision])
  const proposal = proposals.find((item) => item.id === selected)
  const additions = proposals.filter(simpleAddition)
  return (
    <section className="chronology-proposal-review" aria-label="Review chronology updates">
      <div className="chronology-section-head">
        <h2 tabIndex={-1} ref={heading}>
          Review updates <small>{proposals.length}</small>
        </h2>
        <button type="button" onClick={onClose}>
          Back to saved events
        </button>
      </div>
      <p>
        Saved events stay unchanged until you accept a proposal. Review one change at a time; competing accounts are
        retained. Individual acceptance marks the resulting event reviewed.
      </p>
      <div className="chronology-inline chronology-proposal-controls">
        <label>
          Change
          <select
            aria-label="Update to review"
            value={selected ?? ""}
            onChange={(event) => setSelected(event.target.value)}
          >
            {proposals.map((item, index) => (
              <option key={item.id} value={item.id}>
                {index + 1}. {item.kind} · {item.events[0]?.title ?? item.reason}
                {item.status === "deferred" ? " (for later)" : ""}
              </option>
            ))}
          </select>
        </label>
        {additions.length > 0 && (
          <button
            type="button"
            onClick={() => {
              setBatch((value) => !value)
              setBatchIds(additions.map((item) => item.id))
            }}
          >
            Batch straightforward additions ({additions.length})
          </button>
        )}
        <button type="button" onClick={() => setDiscard((value) => !value)}>
          Discard pending refresh…
        </button>
      </div>
      {discard && (
        <div className="chronology-notice">
          <p>Discard the pending refresh proposals? Saved events, reviews, hidden choices and corrections are kept.</p>
          <div className="chronology-inline">
            <button type="button" onClick={() => setDiscard(false)}>
              Keep proposals
            </button>
            <button
              type="button"
              disabled={busy}
              onClick={async () => {
                if (await onDiscard()) setDiscard(false)
              }}
            >
              Confirm discard refresh
            </button>
          </div>
        </div>
      )}
      {batch && (
        <div className="chronology-preview" aria-label="Batch addition preview">
          <h3>Review additions before accepting</h3>
          <p>
            Only straightforward additions are eligible. Conflicts, removals and changes to existing events require
            individual review. Batch additions remain unreviewed.
          </p>
          {additions.map((item) => (
            <label className="chronology-check" key={item.id}>
              <input
                type="checkbox"
                checked={batchIds.includes(item.id)}
                onChange={(change) =>
                  setBatchIds((current) =>
                    change.target.checked ? [...current, item.id] : current.filter((id) => id !== item.id),
                  )
                }
              />
              <span>{item.events.map((event) => `${dateLabel(event.date)} · ${event.title}`).join("; ")}</span>
            </label>
          ))}
          <div className="chronology-inline">
            <button type="button" onClick={() => setBatch(false)}>
              Cancel batch
            </button>
            <button
              type="button"
              className="primary"
              disabled={busy || batchIds.length === 0}
              onClick={async () => {
                if (await onBatch(batchIds)) setBatch(false)
              }}
            >
              Accept {batchIds.length} additions as unreviewed
            </button>
          </div>
        </div>
      )}
      {proposal ? (
        <article key={proposal.id} className="chronology-proposal">
          <div className="chronology-tags">
            <span className="chronology-tag">{proposal.kind}</span>
            <span className="chronology-tag">{proposal.origin}</span>
            {proposal.status === "deferred" && <span className="chronology-tag">For later</span>}
          </div>
          <h3>{proposal.reason}</h3>
          <p>
            <strong>Recommendation:</strong> {proposal.recommendation}
          </p>
          <div className="chronology-proposal-comparison">
            <div>
              <h4>Current saved entry</h4>
              {proposal.targetIds.length ? (
                proposal.targetIds.map((id) => {
                  const event = chronology.events.find((item) => item.id === id)
                  return event ? (
                    <Preview key={id} event={event} onEvidence={onEvidence} />
                  ) : (
                    <p key={id} className="chronology-warning">
                      The previous event is unavailable. Reload before deciding.
                    </p>
                  )
                })
              ) : (
                <p>New entry; no saved event is replaced.</p>
              )}
            </div>
            <div>
              <h4>Proposed result</h4>
              {proposal.events.length ? (
                proposal.events.map((event) => <Preview key={event.id} event={event} onEvidence={onEvidence} />)
              ) : (
                <p>The event would be removed from the active chronology; its history is retained.</p>
              )}
            </div>
          </div>
          <div className="chronology-inline chronology-decision-actions">
            <button
              type="button"
              className="primary"
              disabled={busy}
              onClick={() => void onDecision(proposal, "accept")}
            >
              Accept this change
            </button>
            <button type="button" disabled={busy} onClick={() => void onDecision(proposal, "keep")}>
              Keep current
            </button>
            <button
              type="button"
              disabled={busy}
              onClick={async () => {
                if (await onDecision(proposal, "defer")) {
                  const next = proposals.find((item) => item.id !== proposal.id && item.status === "pending")
                  if (next) setSelected(next.id)
                }
              }}
            >
              Skip for later
            </button>
            {(proposal.events[0] ?? chronology.events.find((event) => proposal.targetIds.includes(event.id))) && (
              <button
                type="button"
                onClick={() =>
                  onAsk(
                    proposal.events[0] ?? chronology.events.find((event) => proposal.targetIds.includes(event.id))!,
                    proposal,
                  )
                }
              >
                Ask doc.haus
              </button>
            )}
          </div>
        </article>
      ) : (
        <p className="chronology-empty">No pending updates. Your saved chronology is unchanged.</p>
      )}
    </section>
  )
}
function Preview({
  event,
  onEvidence,
}: {
  event: ChronologyEvent
  onEvidence: (evidence: ChronologyEvidence) => void
}) {
  return (
    <article className="chronology-preview-event">
      <strong>{event.title}</strong>
      <p>{dateLabel(event.date)}</p>
      <p>{event.detail}</p>
      <div className="chronology-tags">
        {eventQualifications(event).map((text) => (
          <span key={text} className="chronology-tag">
            {text}
          </span>
        ))}
      </div>
      {event.hidden && (
        <p className="chronology-warning">This event is hidden. Refresh does not restore it automatically.</p>
      )}
      {event.lawyerEdited && <p className="chronology-muted">Contains lawyer corrections.</p>}
      {eventEvidence(event).map((evidence) => (
        <div key={evidence.id} className="chronology-evidence">
          <strong>
            {evidence.documentName} · {evidence.section}
          </strong>
          <blockquote>{evidence.excerpt}</blockquote>
          <p className="chronology-muted">
            {evidence.verified ? "Quotation verified" : "Unverified"} · {evidence.freshness}
          </p>
          <button type="button" onClick={() => onEvidence(evidence)}>
            Inspect source
          </button>
        </div>
      ))}
    </article>
  )
}
