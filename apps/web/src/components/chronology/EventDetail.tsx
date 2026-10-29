import { useEffect, useRef, useState } from "react"
import {
  chronologyNeedsAttention,
  type Chronology,
  type ChronologyAction,
  type ChronologyEvent,
  type ChronologyEvidence,
} from "../../api/ingest/chronology"
import { dateLabel, eventQualifications } from "./model"
export default function EventDetail({
  event,
  chronology,
  busy,
  onAction,
  onEdit,
  onClose,
  onEvidence,
  onAsk,
  onNavigate,
}: {
  event: ChronologyEvent
  chronology: Chronology
  busy: boolean
  onAction: (action: ChronologyAction) => Promise<boolean>
  onEdit: (mode: "edit" | "merge" | "split") => void
  onClose: () => void
  onEvidence: (evidence: ChronologyEvidence) => void
  onAsk: () => void
  onNavigate: (eventId: string) => void
}) {
  const heading = useRef<HTMLHeadingElement>(null)
  const [note, setNote] = useState(event.note)
  const [dispute, setDispute] = useState(event.dispute ?? "")
  const [editingDispute, setEditingDispute] = useState(false)
  const priorNote = useRef(event.note)
  const priorDispute = useRef(event.dispute ?? "")
  useEffect(() => {
    heading.current?.focus({ preventScroll: true })
    if (window.matchMedia("(max-width: 800px)").matches)
      heading.current?.closest("aside")?.scrollIntoView({ block: "start", behavior: "instant" })
  }, [event.id])
  useEffect(() => {
    // A refresh may update the saved record while a lawyer is typing. Keep an
    // unsaved note/dispute; adopting a new server revision never erases it.
    const savedNote = priorNote.current,
      savedDispute = priorDispute.current
    setNote((current) => (current === savedNote ? event.note : current))
    setDispute((current) => (current === savedDispute ? (event.dispute ?? "") : current))
    priorNote.current = event.note
    priorDispute.current = event.dispute ?? ""
  }, [event.note, event.dispute])
  return (
    <aside className="chronology-detail" aria-label="Event and evidence">
      <div className="chronology-detail-top">
        <button type="button" onClick={onClose}>
          ← Back to events
        </button>
        <button
          type="button"
          className="chronology-star"
          aria-label={event.key ? "Remove key-event star" : "Mark as key event"}
          aria-pressed={event.key}
          disabled={busy}
          onClick={() => void onAction({ type: "key", eventId: event.id, key: !event.key })}
        >
          {event.key ? "★" : "☆"}
        </button>
      </div>
      <h2 ref={heading} tabIndex={-1}>
        {event.title}
      </h2>
      <p className="chronology-detail-date">{dateLabel(event.date)}</p>
      <div className="chronology-tags">
        <span className={`chronology-tag ${event.reviewed ? "is-reviewed" : "is-pending"}`}>
          {event.reviewed ? "Reviewed" : "Unreviewed"}
        </span>
        {chronologyNeedsAttention(event) && <span className="chronology-tag is-alert">Needs attention</span>}
        {eventQualifications(event).map((text) => (
          <span key={text} className="chronology-tag">
            {text}
          </span>
        ))}
      </div>
      <p>{event.detail}</p>
      {!!event.supersededBy?.length && (
        <div className="chronology-notice">
          <strong>Previous event record</strong>
          <p>This record is retained after a merge or split. Current entries:</p>
          {event.supersededBy.map((id) => (
            <button key={id} type="button" onClick={() => onNavigate(id)}>
              {chronology.events.find((item) => item.id === id)?.title ?? id}
            </button>
          ))}
        </div>
      )}
      {event.disputed && (
        <div className="chronology-notice">
          <strong>Unresolved dispute</strong>
          <p>{event.dispute || "The source accounts conflict. Keep each account until the lawyer resolves it."}</p>
        </div>
      )}
      {event.date.relative && (
        <section>
          <h3>Relative date</h3>
          <p>{event.date.relative.expression}</p>
          <p>
            Anchor: {event.date.relative.anchorDocument} · {event.date.relative.anchorDate ?? "not established"}
          </p>
          {event.date.anchorIssue && <p className="chronology-warning">{event.date.anchorIssue}</p>}
          {event.date.relative.calculation && <p>Calculation: {event.date.relative.calculation}</p>}
          <p>
            {event.date.relative.interpretation === "legal-interpretation"
              ? "Legal interpretation — verify separately"
              : "Document-stated expression"}
          </p>
        </section>
      )}
      {event.date.documentDate && <p className="chronology-muted">Document date: {event.date.documentDate}</p>}
      {event.date.correction && (
        <p className="chronology-muted">
          Date corrected locally by {event.date.correction.actor}. Original: {dateLabel(event.date.correction.original)}
          . The source file is unchanged.
        </p>
      )}
      <section>
        <h3>Source accounts & evidence</h3>
        <p className="chronology-muted">
          A verified quotation is a text match, not proof that the reported event occurred. Source types carry no
          automatic credibility ranking.
        </p>
        {event.accounts.length === 0 && (
          <p className="chronology-notice">Source not supplied. This entry was added by hand.</p>
        )}
        {event.accounts.map((account) => (
          <article key={account.id} className="chronology-account">
            <div className="chronology-account-label">
              <strong>{account.title}</strong>
              <span>{account.status}</span>
            </div>
            <p>
              {dateLabel(account.date)}
              {account.attribution ? ` · ${account.attribution}` : ""}
            </p>
            <p>{account.detail}</p>
            {account.date.anchorIssue && <p className="chronology-warning">{account.date.anchorIssue}</p>}
            <p className="chronology-muted">
              Significance: {account.significance} · Privilege: {account.privilege}
              {account.people.length ? ` · ${account.people.join(", ")}` : ""}
            </p>
            {account.nameAssociations?.map((association, index) => (
              <p key={index} className="chronology-muted">
                Name association: {association.name} → {association.canonicalName} ·{" "}
                {association.status === "uncertain"
                  ? "Uncertain; kept separate in person filters. Review the supporting passage."
                  : association.evidenceIds.every((id) =>
                        account.evidence.some(
                          (evidence) => evidence.id === id && evidence.verified && evidence.freshness === "current",
                        ),
                      )
                    ? "Source-supported; person filters use this association unless another source assigns the same name differently."
                    : "Support is stale or unavailable; kept separate in person filters."}
              </p>
            ))}
            {account.obligationId && (
              <p className="chronology-muted">
                Linked obligation: {account.obligationId}. The original deadline is preserved.
              </p>
            )}
            {account.evidence.map((evidence) => (
              <div key={evidence.id} className="chronology-evidence">
                <div className="chronology-evidence-heading">
                  <strong>{evidence.documentName}</strong>
                  <span>{evidence.section || "Passage"}</span>
                </div>
                <blockquote>{evidence.excerpt}</blockquote>
                <p className="chronology-muted">
                  {evidence.verified ? "Quotation verified" : "Quotation unverified"} · {evidence.freshness} · Version{" "}
                  {evidence.version}
                  {evidence.sourceKind ? ` · ${evidence.sourceKind}` : ""}
                </p>
                {evidence.provenance === "supplied-chronology" && (
                  <p className="chronology-notice">
                    Supplied chronology account. Original provenance may be unknown; this does not establish independent
                    corroboration.
                  </p>
                )}
                {evidence.issue && (
                  <p role="status" className="chronology-warning">
                    {evidence.issue}
                  </p>
                )}
                <button type="button" onClick={() => onEvidence(evidence)}>
                  {evidence.freshness === "current" ? "View in document" : "Inspect current source"}
                </button>
              </div>
            ))}
            {account.evidence.length === 0 && (
              <p className="chronology-warning">No verified supporting passage is available.</p>
            )}
            {account.gaps.map((gap, index) => (
              <p key={index} className="chronology-warning">
                {gap}
              </p>
            ))}
          </article>
        ))}
      </section>
      <section>
        <h3>Lawyer review</h3>
        <div className="chronology-inline chronology-detail-actions">
          <button
            type="button"
            className="primary"
            disabled={busy}
            onClick={() => void onAction({ type: "review", eventId: event.id, reviewed: !event.reviewed })}
          >
            {event.reviewed ? "Mark unreviewed" : "Mark reviewed"}
          </button>
          <button type="button" disabled={busy} onClick={() => onEdit("edit")}>
            Edit / correct date
          </button>
          <button type="button" onClick={onAsk}>
            Ask doc.haus
          </button>
        </div>
        <p className="chronology-muted">Review does not resolve a dispute or renew stale evidence.</p>
        <form
          onSubmit={async (submit) => {
            submit.preventDefault()
            await onAction({ type: "note", eventId: event.id, note })
          }}
        >
          <label>
            Optional review note
            <textarea rows={3} value={note} onChange={(change) => setNote(change.target.value)} />
          </label>
          <button type="submit" disabled={busy || note === event.note}>
            Save note
          </button>
        </form>
        <div className="chronology-inline chronology-detail-actions">
          <button type="button" disabled={busy} onClick={() => setEditingDispute((value) => !value)}>
            Record / update dispute
          </button>
          <button type="button" disabled={busy} onClick={() => onEdit("merge")}>
            Merge…
          </button>
          <button type="button" disabled={busy} onClick={() => onEdit("split")}>
            Split…
          </button>
          <button
            type="button"
            disabled={busy}
            onClick={() => void onAction({ type: "hide", eventId: event.id, hidden: !event.hidden })}
          >
            {event.hidden ? "Restore event" : "Hide event"}
          </button>
        </div>
        {editingDispute && (
          <form
            onSubmit={async (submit) => {
              submit.preventDefault()
              if (await onAction({ type: "dispute", eventId: event.id, dispute, disputed: true }))
                setEditingDispute(false)
            }}
          >
            <label>
              Dispute or competing account
              <textarea value={dispute} onChange={(change) => setDispute(change.target.value)} />
            </label>
            <div className="chronology-inline">
              <button type="submit" disabled={busy}>
                Save unresolved dispute
              </button>
              {event.disputed && (
                <button
                  type="button"
                  disabled={busy}
                  onClick={async () => {
                    if (await onAction({ type: "dispute", eventId: event.id, dispute, disputed: false }))
                      setEditingDispute(false)
                  }}
                >
                  Record dispute resolved
                </button>
              )}
            </div>
          </form>
        )}
      </section>
      {event.original && (
        <details>
          <summary>Original generated entry</summary>
          <p>
            {event.original.title} · {dateLabel(event.original.date)}
          </p>
          <p>{event.original.detail}</p>
        </details>
      )}
      <details>
        <summary>Local change history</summary>
        <p className="chronology-muted">Local operator/session attribution; not authenticated multi-user sign-off.</p>
        {chronology.history
          .filter(
            (item) => item.eventIds.includes(event.id) || item.eventIds.some((id) => event.previousIds.includes(id)),
          )
          .slice()
          .reverse()
          .map((item) => (
            <p key={item.id}>
              <time>{new Date(item.at).toLocaleString()}</time> · {item.actor} · {item.action}
              {item.detail ? ` · ${item.detail}` : ""}
            </p>
          ))}
        {event.previousIds.length > 0 && (
          <div>
            <p>Previous records:</p>
            {event.previousIds.map((id) => (
              <button key={id} type="button" onClick={() => onNavigate(id)}>
                {chronology.events.find((item) => item.id === id)?.title ?? id}
              </button>
            ))}
          </div>
        )}
      </details>
    </aside>
  )
}
