import { Link } from "react-router-dom"
import type { ChronologyChatReference } from "./chronology/model"
import "./ChronologyReferenceCard.css"

export default function ChronologyReferenceCard({
  reference,
  sessionID,
  busy,
  onRemove,
}: {
  reference: ChronologyChatReference
  sessionID?: string
  busy?: boolean
  onRemove: () => void
}) {
  const sourceCount = new Set(reference.sources.map((source) => source.documentName)).size
  const query = new URLSearchParams(reference.link.split("?")[1])
  if (sessionID) query.set("session", sessionID)
  const link = `${reference.link.split("?")[0]}?${query}`
  return (
    <article className="chronology-reference" aria-label={`Chronology reference: ${reference.title}`}>
      <div className="chronology-reference-heading">
        <div>
          <span className="chronology-reference-kind">Chronology reference</span>
          <strong>{reference.title}</strong>
          <p>
            {reference.date} · {reference.reviewed ? "Reviewed" : "Unreviewed"}
            {reference.hidden ? " · Hidden in chronology" : ""}
            {reference.proposalId ? " · Pending proposal included" : ""}
          </p>
        </div>
        <button
          type="button"
          className="chronology-reference-remove"
          disabled={busy}
          aria-label={`Remove chronology reference: ${reference.title}`}
          onClick={onRemove}
        >
          ×
        </button>
      </div>
      {reference.qualifications.length > 0 && (
        <p className="chronology-reference-qualifications">{reference.qualifications.join(" · ")}</p>
      )}
      <div className="chronology-reference-actions">
        <Link
          to={link}
          aria-disabled={busy || undefined}
          onClick={(event) => {
            if (busy) event.preventDefault()
          }}
        >
          Return to chronology
        </Link>
        <span>
          {sourceCount} {sourceCount === 1 ? "source" : "sources"}
        </span>
      </div>
      <details>
        <summary>Included evidence</summary>
        <p className="chronology-reference-snapshot">
          Included as a snapshot. A verified quotation is a text match; review the source and its current status before
          relying on it.
        </p>
        {reference.sources.length ? (
          reference.sources.map((source, index) => (
            <section className="chronology-reference-evidence" key={index}>
              <strong>{source.documentName}</strong>
              <p>
                {source.scope === "proposal" ? "Proposed evidence · " : ""}
                {source.section ? `${source.section} · ` : ""}
                {source.verified ? "Quotation verified" : "Quotation unverified"} · {source.freshness}
              </p>
              {source.provenance === "supplied-chronology" && (
                <p>
                  Supplied chronology — its underlying original source is not independently established by this passage.
                </p>
              )}
              {source.excerpt ? <blockquote>{source.excerpt}</blockquote> : <p>Source passage unavailable.</p>}
            </section>
          ))
        ) : (
          <p>No supporting source was supplied.</p>
        )}
      </details>
    </article>
  )
}
