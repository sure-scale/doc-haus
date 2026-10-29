import { useEffect, useRef, useState } from "react"
import type {
  Chronology,
  ChronologyDecision,
  ChronologyEvent,
  ChronologyEvidence,
  ChronologyProposal,
} from "../../api/ingest/chronology"
import JurisdictionSelect from "../JurisdictionSelect"
import { dateLabel, eventEvidence, eventQualifications, simpleAddition } from "./model"
import "./proposal-review.css"

const purposes: Record<ChronologyProposal["kind"], { title: string; explanation: string; short: string }> = {
  addition: {
    title: "Add an event",
    short: "New event",
    explanation:
      "This event is not yet in your saved chronology. Check the proposed entry and its sources before adding it.",
  },
  amendment: {
    title: "Update this event",
    short: "Update",
    explanation:
      "Compare the saved entry with the proposed result. Your saved wording, dates and review stay as they are until you accept.",
  },
  conflict: {
    title: "Review a competing account",
    short: "Competing account",
    explanation:
      "The sources give different accounts. Accepting retains the existing accounts and the disagreement; it does not decide which account is true.",
  },
  removal: {
    title: "Hide this event",
    short: "Hide event",
    explanation:
      "This proposes hiding an entry from the active chronology. The entry, evidence and history are retained, and you can restore it later.",
  },
  merge: {
    title: "Combine these events",
    short: "Combine events",
    explanation:
      "Check that these entries describe the same occurrence. The combined result retains the source accounts and links to the previous records.",
  },
  split: {
    title: "Separate this event",
    short: "Separate events",
    explanation:
      "Check the proposed separate occurrences. Their source accounts and links to the previous entry are retained.",
  },
}

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
    if (!proposals.some((proposal) => proposal.id === selected))
      setSelected(proposals.find((proposal) => proposal.status === "pending")?.id ?? proposals[0]?.id)
    setBatchIds((current) =>
      current.filter((id) => proposals.some((proposal) => proposal.id === id && simpleAddition(proposal))),
    )
    // A stored revision changes the available decisions; local navigation does not.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [chronology.revision])
  const proposal = proposals.find((item) => item.id === selected)
  const index = proposals.findIndex((item) => item.id === selected)
  const additions = proposals.filter(simpleAddition)
  const later = proposals.filter((item) => item.status === "deferred").length
  const saved = proposal
    ? proposal.targetIds.flatMap((id) => {
        const event = chronology.events.find((event) => event.id === id)
        return event ? [event] : []
      })
    : []
  const eventTitle = (item: ChronologyProposal) =>
    item.events[0]?.title ??
    chronology.events.find((event) => item.targetIds.includes(event.id))?.title ??
    "Previous event unavailable"
  useEffect(() => {
    if (!batch && !discard) heading.current?.focus({ preventScroll: true })
  }, [proposal?.id, batch, discard])

  return (
    <section className="chronology-proposal-review update-review" aria-label="Review chronology updates">
      <div className="update-review-intro">
        <p className="update-review-progress" role="status">
          {proposals.length
            ? `${proposals.length} ${proposals.length === 1 ? "change awaiting" : "changes awaiting"} your decision${later ? ` · ${later} saved for later` : ""}`
            : "No changes awaiting a decision"}
        </p>
        <p>Review one change at a time. Your saved chronology stays unchanged until you accept.</p>
      </div>
      {proposals.length > 0 && (
        <div className="update-review-navigation">
          <div className="update-review-picker">
            <span className="update-review-field-label">Choose a change</span>
            <JurisdictionSelect
              jurisdictions={proposals.map((item, at) => ({
                code: item.id,
                name: `${at + 1}. ${purposes[item.kind].short} — ${eventTitle(item)}${item.status === "deferred" ? " · For later" : ""}`,
              }))}
              selected={selected ? [selected] : []}
              multiple={false}
              showCode={false}
              disabled={busy}
              label="Update to review"
              searchLabel="Search updates"
              emptyLabel="Choose an update"
              noResultsLabel="No updates match your search."
              trigger={
                <>
                  <span>
                    <small>
                      {index + 1} of {proposals.length}
                    </small>
                    <strong>{proposal ? eventTitle(proposal) : "Choose an update"}</strong>
                  </span>
                  <span aria-hidden="true">⌄</span>
                </>
              }
              onChange={(ids) => {
                if (ids[0]) {
                  setSelected(ids[0])
                  setBatch(false)
                  setDiscard(false)
                }
              }}
            />
          </div>
          <div className="update-review-step-buttons">
            <button
              type="button"
              aria-label="Previous update"
              title="Previous update"
              disabled={busy || index <= 0}
              onClick={() => {
                setSelected(proposals[index - 1]?.id)
                setBatch(false)
                setDiscard(false)
              }}
            >
              ‹
            </button>
            <button
              type="button"
              aria-label="Next update"
              title="Next update"
              disabled={busy || index < 0 || index >= proposals.length - 1}
              onClick={() => {
                setSelected(proposals[index + 1]?.id)
                setBatch(false)
                setDiscard(false)
              }}
            >
              ›
            </button>
          </div>
        </div>
      )}
      <div className="update-review-secondary-actions">
        {additions.length > 0 && (
          <button
            type="button"
            disabled={busy}
            aria-label={`Batch straightforward additions (${additions.length})`}
            onClick={() => {
              setBatch((value) => !value)
              setBatchIds(additions.map((item) => item.id))
              setDiscard(false)
            }}
          >
            Add straightforward events together ({additions.length})
          </button>
        )}
        {proposals.length > 0 && (
          <button
            type="button"
            className="update-review-discard"
            disabled={busy}
            onClick={() => {
              setDiscard((value) => !value)
              setBatch(false)
            }}
          >
            Discard pending refresh…
          </button>
        )}
      </div>
      {discard && (
        <div className="update-review-notice" role="note">
          <strong>Discard the pending refresh proposals?</strong>
          <p>Your saved events, reviews, notes, hidden choices and corrections are kept.</p>
          <div className="update-review-inline">
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
        <div className="chronology-preview update-review-batch" aria-label="Batch addition preview">
          <h3>Add new events together</h3>
          <p>
            Only straightforward new events are included. They will remain <strong>unreviewed</strong>, ready to check
            individually in the chronology. Changes to existing events, conflicts and hiding decisions need individual
            review.
          </p>
          <div className="update-review-batch-list">
            {additions.map((item) => (
              <label className="chronology-check" key={item.id}>
                <input
                  type="checkbox"
                  disabled={busy}
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
          </div>
          <div className="update-review-inline">
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
      {proposal && !batch && !discard ? (
        <article key={proposal.id} className="chronology-proposal update-review-change">
          <header className="update-review-purpose">
            <p className="update-review-eyebrow">
              {proposal.status === "deferred" ? "Set aside for later" : "Suggested change"}
            </p>
            <h3 ref={heading} tabIndex={-1}>
              {purposes[proposal.kind].title}
            </h3>
            <p>{purposes[proposal.kind].explanation}</p>
          </header>
          <div
            className={`update-review-comparison${proposal.kind === "addition" || proposal.kind === "removal" ? " is-single" : ""}`}
          >
            {proposal.kind !== "addition" && (
              <section className="update-review-saved" aria-label="Current saved entry">
                <h4>{saved.length > 1 ? "Saved entries now" : "Saved entry now"}</h4>
                {saved.map((event) => (
                  <Preview key={event.id} event={event} onEvidence={onEvidence} />
                ))}
                {saved.length !== proposal.targetIds.length && (
                  <p className="chronology-warning">A previous entry is unavailable. Reload before deciding.</p>
                )}
              </section>
            )}
            {proposal.kind !== "removal" && (
              <section className="update-review-proposed" aria-label="Proposed result">
                <h4>{proposal.kind === "addition" ? "New event to add" : "If you accept"}</h4>
                {proposal.events.map((event) => (
                  <Preview key={event.id} event={event} proposed onEvidence={onEvidence} />
                ))}
              </section>
            )}
          </div>
          {proposal.kind === "removal" && (
            <div className="update-review-hide-note">
              <strong>After accepting: hidden, with its history kept.</strong>
              <p>
                This does not delete the record or establish that the event did not happen. You can find and restore it
                using the hidden-events filter.
              </p>
            </div>
          )}
          <div className="update-review-reason">
            <h4>Why this is suggested</h4>
            <p>{proposal.reason}</p>
            <h4>What to consider</h4>
            <p>{proposal.recommendation}</p>
          </div>
          <p className="update-review-evidence-note">
            A verified quotation confirms a text match, not that an account is true.
          </p>
          <div className="update-review-decision-area">
            <div className="update-review-decisions">
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
              {(proposal.events[0] ?? saved[0]) && (
                <button
                  type="button"
                  className="update-review-ask"
                  disabled={busy}
                  onClick={() => onAsk(proposal.events[0] ?? saved[0]!, proposal)}
                >
                  Ask doc.haus
                </button>
              )}
            </div>
            <p>
              <strong>Accept</strong>{" "}
              {proposal.kind === "removal"
                ? "hides the entry and keeps its previous review."
                : "saves the result as reviewed."}{" "}
              <strong>Keep current</strong> rejects this suggestion
              {proposal.kind === "addition" ? " without adding the event" : ""}. <strong>Skip</strong> keeps it for a
              later decision.
            </p>
          </div>
        </article>
      ) : !proposal ? (
        <div className="update-review-empty">
          <span aria-hidden="true">✓</span>
          <h3 ref={heading} tabIndex={-1}>
            No pending updates
          </h3>
          <p>Your saved chronology is ready to return to.</p>
        </div>
      ) : null}
      <footer className="update-review-footer">
        <button type="button" onClick={onClose}>
          Back to saved events
        </button>
      </footer>
    </section>
  )
}

function Preview({
  event,
  proposed = false,
  onEvidence,
}: {
  event: ChronologyEvent
  proposed?: boolean
  onEvidence: (evidence: ChronologyEvidence) => void
}) {
  const evidence = eventEvidence(event)
  const qualifications = eventQualifications(event)
  return (
    <article className="chronology-preview-event update-review-entry">
      <strong className="update-review-event-title">{event.title}</strong>
      <p className="update-review-event-date">
        {dateLabel(event.date)}
        {!proposed ? ` · ${event.reviewed ? "Reviewed" : "Unreviewed"}` : ""}
      </p>
      <p>{event.detail}</p>
      {qualifications.length > 0 && <p className="update-review-qualifications">{qualifications.join(" · ")}</p>}
      {event.hidden && (
        <p className="chronology-warning">
          This entry is hidden. Accepting an update does not restore it automatically.
        </p>
      )}
      {event.lawyerEdited && <p className="chronology-muted">Includes lawyer corrections.</p>}
      {event.accounts.length > 1 && (
        <details className="update-review-accounts">
          <summary>Source accounts ({event.accounts.length})</summary>
          {event.accounts.map((account) => (
            <div key={account.id}>
              <strong>{dateLabel(account.date)}</strong>
              {account.title !== event.title && <p>{account.title}</p>}
              {account.detail && account.detail !== event.detail && <p>{account.detail}</p>}
              {account.attribution && <p className="chronology-muted">Reported by {account.attribution}</p>}
            </div>
          ))}
        </details>
      )}
      <details className="update-review-evidence" open>
        <summary>Supporting evidence ({evidence.length})</summary>
        {evidence.length ? (
          evidence.map((source, index) => (
            <div key={`${source.id}:${index}`} className="update-review-source">
              <strong>
                {source.documentName}
                {source.section ? ` · ${source.section}` : ""}
              </strong>
              <blockquote>{source.excerpt || "Source passage unavailable."}</blockquote>
              <p>
                {source.verified ? "Quotation verified" : "Quotation unverified"} ·{" "}
                {source.freshness === "current"
                  ? "Source unchanged"
                  : source.freshness === "stale"
                    ? "Source changed — recheck"
                    : source.freshness === "unavailable"
                      ? "Source unavailable"
                      : "Source not checked"}
              </p>
              {source.provenance === "supplied-chronology" && (
                <p className="chronology-warning">
                  Supplied chronology — the underlying original source is not independently established by this passage.
                </p>
              )}
              <button type="button" onClick={() => onEvidence(source)}>
                Inspect source
              </button>
            </div>
          ))
        ) : (
          <p className="chronology-warning">No supporting source was supplied.</p>
        )}
      </details>
    </article>
  )
}
