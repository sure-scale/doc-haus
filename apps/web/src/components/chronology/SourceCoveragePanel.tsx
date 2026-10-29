import { useId } from "react"
import type { Chronology, ChronologyCoverage, ChronologyEvent } from "../../api/ingest/chronology"
import SourceIcon from "./SourceIcon"
import "./source-coverage.css"

const statusLabels: Record<ChronologyCoverage["status"], string> = {
  complete: "Recorded complete",
  partial: "Partly covered",
  pending: "Not yet covered",
  processing: "Unfinished",
  failed: "Could not finish",
  unreadable: "No readable text",
  unsupported: "Format not supported",
  removed: "No longer available",
}

function citedSources(event: ChronologyEvent) {
  return new Set(event.accounts.flatMap((account) => account.evidence.map((evidence) => evidence.sourceId)))
}

export default function SourceCoveragePanel({
  chronology,
  onSourceDate,
}: {
  chronology: Chronology
  onSourceDate: () => void
}) {
  const heading = useId()
  const active = chronology.coverage.filter((source) => source.status !== "removed")
  const complete = active.filter((source) => source.status === "complete").length
  const removed = chronology.coverage.length - active.length
  const saved = new Map<string, Set<string>>()
  const historical = new Map<string, Set<string>>()
  const suggestions = new Map<string, Set<string>>()
  for (const event of chronology.events) {
    const index = event.supersededBy?.length ? historical : saved
    for (const sourceId of citedSources(event)) {
      const entries = index.get(sourceId) ?? new Set<string>()
      entries.add(event.id)
      index.set(sourceId, entries)
    }
  }
  for (const proposal of chronology.proposals.filter(
    (item) => item.status === "pending" || item.status === "deferred",
  )) {
    const referenced = [
      ...proposal.events,
      ...chronology.events.filter((event) => proposal.targetIds.includes(event.id)),
    ]
    for (const sourceId of new Set(referenced.flatMap((event) => [...citedSources(event)]))) {
      const updates = suggestions.get(sourceId) ?? new Set<string>()
      updates.add(proposal.id)
      suggestions.set(sourceId, updates)
    }
  }
  const running = chronology.run?.status === "running"
  return (
    <section className="source-coverage" aria-labelledby={heading}>
      <div className="source-coverage-heading">
        <div>
          <h2 id={heading}>Saved source coverage</h2>
          <p>See which files were available and which sections were recorded as read.</p>
        </div>
        <button type="button" onClick={onSourceDate}>
          Supply / correct source date
        </button>
      </div>
      {chronology.coverage.length > 0 && (
        <dl className="source-coverage-summary">
          <div>
            <dt>Sources in the record</dt>
            <dd>{chronology.coverage.length}</dd>
          </div>
          <div>
            <dt>Recorded complete</dt>
            <dd>{complete}</dd>
          </div>
          <div>
            <dt>Not complete</dt>
            <dd>{active.length - complete}</dd>
          </div>
          {removed > 0 && (
            <div>
              <dt>No longer available</dt>
              <dd>{removed}</dd>
            </div>
          )}
        </dl>
      )}
      <p className="source-coverage-context">
        {running && <strong>Building chronology. </strong>}
        Saved coverage includes earlier attempts and does not show live reading progress.
      </p>
      {chronology.run?.errors.length ? (
        <section className="source-coverage-issues" aria-label="Processing issues">
          <h3>Processing issues to check</h3>
          <ul>
            {chronology.run.errors.map((issue, index) => (
              <li key={index}>{issue}</li>
            ))}
          </ul>
        </section>
      ) : null}
      {chronology.coverage.length === 0 ? (
        <div className="source-coverage-empty">
          <h3>No source coverage recorded yet</h3>
          <p>
            Building the chronology records the available source set. The source check lets you inspect the included
            files before making a fresh privilege choice.
          </p>
        </div>
      ) : (
        <ul className="source-coverage-list">
          {chronology.coverage.map((source, index) => {
            const sections = [...new Set(source.processedSections)].sort((a, b) => a - b)
            const currentEntries = saved.get(source.sourceId)?.size ?? 0
            const pendingUpdates = suggestions.get(source.sourceId)?.size ?? 0
            const previousEntries = historical.get(source.sourceId)?.size ?? 0
            const sourceDate = chronology.sourceDates[source.documentName]
            const caution = source.status !== "complete" || source.issues.length > 0
            return (
              <li key={source.sourceId}>
                <article className="source-coverage-card" aria-labelledby={`${heading}-${index}`}>
                  <div className="source-coverage-card-heading">
                    <SourceIcon name={source.documentName} />
                    <h3 id={`${heading}-${index}`}>{source.documentName}</h3>
                    <span className={`source-coverage-status ${caution ? "needs-checking" : "is-complete"}`}>
                      {statusLabels[source.status]}
                    </span>
                  </div>
                  {source.suppliedChronology && (
                    <p className="source-coverage-provenance">
                      Supplied chronology. Original provenance may be unknown; this does not establish independent
                      corroboration.
                    </p>
                  )}
                  <dl className="source-coverage-facts">
                    <div>
                      <dt>File</dt>
                      <dd>{source.uploaded ? "Uploaded" : "Unavailable"}</dd>
                    </div>
                    <div>
                      <dt>Text</dt>
                      <dd>{source.textExtracted ? "Extracted" : "Not available"}</dd>
                    </div>
                    <div>
                      <dt>Recorded reading</dt>
                      <dd>
                        {source.sections ? `${sections.length} of ${source.sections} sections` : "No sections recorded"}
                      </dd>
                    </div>
                  </dl>
                  <p className="source-coverage-contribution">
                    Cited in{" "}
                    <strong>
                      {currentEntries} saved {currentEntries === 1 ? "entry" : "entries"}
                    </strong>
                    {pendingUpdates > 0 && (
                      <>
                        {" "}
                        · Used in{" "}
                        <strong>
                          {pendingUpdates} suggested {pendingUpdates === 1 ? "update" : "updates"}
                        </strong>
                      </>
                    )}
                  </p>
                  {sourceDate && (
                    <p className="source-coverage-date">
                      Supplied source date: {sourceDate.date}
                      {sourceDate.freshness && sourceDate.freshness !== "current"
                        ? ` · ${sourceDate.freshness}; check before relying on it`
                        : ""}
                    </p>
                  )}
                  {sourceDate?.issue && <p className="source-coverage-warning">{sourceDate.issue}</p>}
                  {source.issues.length > 0 && (
                    <ul className="source-coverage-card-issues">
                      {source.issues.map((issue, issueIndex) => (
                        <li key={issueIndex}>{issue}</li>
                      ))}
                    </ul>
                  )}
                  <details className="source-coverage-source-detail">
                    <summary>Sections and contributions</summary>
                    <p>
                      <strong>Sections recorded as read:</strong>{" "}
                      {sections.length ? sections.map((section) => section + 1).join(", ") : "None recorded"}.
                    </p>
                    <p>
                      {source.events} source-account {source.events === 1 ? "record" : "records"} retained. This count
                      can include earlier records and proposed changes; it is not a count of distinct events.
                    </p>
                    <p>
                      The saved-entry count above includes hidden entries and excludes superseded ones. Suggested
                      updates have not been accepted.
                      {previousEntries > 0
                        ? ` ${previousEntries} earlier ${previousEntries === 1 ? "entry also cites" : "entries also cite"} this source.`
                        : ""}
                    </p>
                    {sourceDate && (
                      <p>Source date recorded locally by {sourceDate.actor}. The source file is unchanged.</p>
                    )}
                  </details>
                </article>
              </li>
            )
          })}
        </ul>
      )}
      {chronology.gaps.length > 0 && (
        <section className="source-coverage-questions" aria-label="Evidence questions to check">
          <h3>Evidence questions to check</h3>
          <p>
            Check these references against the whole record. Material not found in one section may be supplied elsewhere
            in the matter.
          </p>
          <ul>
            {chronology.gaps.slice(0, 3).map((gap, index) => (
              <li key={index}>{gap}</li>
            ))}
          </ul>
          {chronology.gaps.length > 3 && (
            <details className="source-coverage-question-details">
              <summary>Show all {chronology.gaps.length} references</summary>
              <ul>
                {chronology.gaps.slice(3).map((gap, index) => (
                  <li key={index + 3}>{gap}</li>
                ))}
              </ul>
            </details>
          )}
        </section>
      )}
      <details className="source-coverage-method">
        <summary>How source processing works</summary>
        <p>
          These statuses include work saved from earlier attempts. They are not a progress estimate for the current
          build; an unfinished status does not mean that source is being read right now.
        </p>
        <p>
          The first build covers the whole authorised matter. Later checks normally process only new, changed, or
          unfinished source sections; unchanged completed sections are reused.
        </p>
        <p>
          Display filters change only the events shown. They never restrict the source set. Every build or update starts
          with a fresh privilege choice.
        </p>
        <p>
          Initial results are saved as an unreviewed draft. Later proposed changes wait for your review. Completed work,
          review notes and previous source accounts are kept.
        </p>
        <p>
          Uploaded files, extracted text and chronology coverage are different checks. Coverage describes recorded
          reading; it does not establish that a reported account is true or independently corroborated. No identified
          events does not prove that nothing happened.
        </p>
      </details>
      <p className="source-coverage-limitations">
        Archives and unreadable material may contain sources that could not be covered. A quiet period describes the
        supplied record; it does not prove that nothing occurred or that evidence is missing.
      </p>
    </section>
  )
}
