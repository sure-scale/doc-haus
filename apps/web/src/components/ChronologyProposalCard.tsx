import { useEffect, useState } from "react"
import { Link, useParams } from "react-router-dom"
import {
  ChronologyConflict,
  decideChronologyProposal,
  getChronology,
  type Chronology,
  type ChronologyDecision,
  type ChronologyEvent,
} from "../api/ingest/chronology"
import "./ChronologyProposalCard.css"

/** Confirmation is application-owned. Reloading a stale proposal displays its
 * current basis and requires another click; it never retries a decision. */
export default function ChronologyProposalCard({ matterId, proposalIds }: { matterId: string; proposalIds: string[] }) {
  const { id } = useParams()
  const [state, setState] = useState<Chronology>()
  const [error, setError] = useState("")
  const [busy, setBusy] = useState(false)
  const [open, setOpen] = useState<string>()
  useEffect(() => {
    let disposed = false
    if (id !== matterId) return
    getChronology(matterId).then(
      (value) => {
        if (!disposed) setState(value)
      },
      (cause) => {
        if (!disposed) setError(cause instanceof Error ? cause.message : String(cause))
      },
    )
    return () => {
      disposed = true
    }
  }, [matterId, id, proposalIds.join("|")])
  if (id !== matterId) return null
  const decide = async (proposalId: string, decision: ChronologyDecision) => {
    if (!state || busy) return
    setBusy(true)
    setError("")
    try {
      setState(await decideChronologyProposal(matterId, state.revision, proposalId, decision))
    } catch (cause) {
      if (cause instanceof ChronologyConflict) {
        try {
          setState(await getChronology(matterId))
        } catch {
          /* Keep visible basis until a successful reload. */
        }
        setError("The chronology or its evidence changed. Review the reloaded entry before deciding again.")
      } else setError(cause instanceof Error ? cause.message : String(cause))
    } finally {
      setBusy(false)
    }
  }
  return (
    <section className="chat-chronology-proposals" aria-label="Chronology proposals awaiting your decision">
      <strong>Chronology changes</strong>
      <p>Proposals leave your saved events unchanged until you accept each change.</p>
      {error && <p role="alert">{error}</p>}
      {!state && !error && <p role="status">Loading current proposals…</p>}
      {state &&
        proposalIds.map((proposalId) => {
          const proposal = state.proposals.find((item) => item.id === proposalId)
          if (!proposal)
            return (
              <p key={proposalId}>
                This proposal is no longer available.{" "}
                <Link to={`/matter/${encodeURIComponent(matterId)}?view=chronology`}>Open chronology</Link>
              </p>
            )
          return (
            <article key={proposal.id}>
              <h3>
                {proposal.kind}: {proposal.events[0]?.title ?? proposal.reason}
              </h3>
              <p>{proposal.reason}</p>
              <Link
                to={`/matter/${encodeURIComponent(matterId)}?view=chronology&proposal=${encodeURIComponent(proposal.id)}`}
              >
                Review in workbench
              </Link>
              {proposal.status === "accepted" || proposal.status === "kept" ? (
                <p role="status">
                  {proposal.status === "accepted" ? "Accepted by the local operator." : "Current entry kept."}
                </p>
              ) : (
                <>
                  <button
                    type="button"
                    aria-expanded={open === proposal.id}
                    onClick={() => setOpen((value) => (value === proposal.id ? undefined : proposal.id))}
                  >
                    Inspect exact change and sources
                  </button>
                  {open === proposal.id && (
                    <div>
                      <h4>Current saved entry</h4>
                      {proposal.targetIds.length ? (
                        proposal.targetIds.map((target) => {
                          const event = state.events.find((item) => item.id === target)
                          return event && <Entry key={target} event={event} />
                        })
                      ) : (
                        <p>New event.</p>
                      )}
                      <h4>Proposed result</h4>
                      {proposal.events.length ? (
                        proposal.events.map((event) => <Entry key={event.id} event={event} />)
                      ) : (
                        <p>Hide the event reversibly; preserve its history.</p>
                      )}
                      <p>
                        <strong>Recommendation:</strong> {proposal.recommendation}
                      </p>
                      <div className="chat-chronology-decisions">
                        <button type="button" disabled={busy} onClick={() => void decide(proposal.id, "accept")}>
                          Accept this change
                        </button>
                        <button type="button" disabled={busy} onClick={() => void decide(proposal.id, "keep")}>
                          Keep current
                        </button>
                        <button type="button" disabled={busy} onClick={() => void decide(proposal.id, "defer")}>
                          Skip for later
                        </button>
                      </div>
                    </div>
                  )}
                </>
              )}
            </article>
          )
        })}
    </section>
  )
}

function Entry({ event }: { event: ChronologyEvent }) {
  return (
    <div className="chat-chronology-entry">
      <strong>{event.title}</strong>
      <p>
        {event.date.original} · {event.date.precision}
      </p>
      <p>{event.detail}</p>
      <p>
        {event.reviewed ? "Reviewed" : "Unreviewed"}
        {event.hidden ? " · Hidden" : ""}
        {event.disputed ? " · Disputed" : ""}
        {event.lawyerEdited ? " · Lawyer corrections" : ""}
      </p>
      {event.accounts.map((account) => (
        <div key={account.id}>
          <strong>{account.title}</strong>
          <p>{account.detail}</p>
          <p>
            Source account date: {account.date.original} · {account.date.precision}
            {account.date.qualifier ? ` · ${account.date.qualifier}` : ""}
          </p>
          <p>
            {account.attribution} · {account.status} · privilege: {account.privilege}
          </p>
          {account.gaps.map((gap, index) => (
            <p key={index}>Evidence gap: {gap}</p>
          ))}
          {account.evidence.map((evidence) => (
            <div key={evidence.id}>
              <strong>
                {evidence.documentName} · {evidence.section}
              </strong>
              {evidence.provenance === "supplied-chronology" && <p>Supplied chronology; origin may be unknown.</p>}
              <blockquote>{evidence.excerpt || "Source unavailable in this matter."}</blockquote>
              <p>
                {evidence.verified ? "Quotation verified" : "Unverified quotation"} · {evidence.freshness}
                {evidence.issue ? ` · ${evidence.issue}` : ""}
              </p>
            </div>
          ))}
        </div>
      ))}
    </div>
  )
}
