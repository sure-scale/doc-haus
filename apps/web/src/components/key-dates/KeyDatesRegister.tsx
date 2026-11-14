import { useEffect, useState } from "react"
import { getKeyDates, type KeyDateRegisterRow } from "../../api/ingest/key-dates"
import "./key-dates.css"

// The matter's key dates, read straight from the register that reviews write.
// Nothing here confirms a deadline: a resolved calendar date is still a record
// of what a source says, and a relative formula stays visibly unresolved until
// someone fixes its inputs.

const KIND_LABELS: Record<string, string> = {
  renewal: "Renewal",
  termination_notice: "Termination notice",
  expiry: "Expiry",
  payment: "Payment",
  deliverable: "Deliverable",
  other: "Other",
}

const DAY = new Intl.DateTimeFormat(undefined, { day: "numeric", month: "short", year: "numeric", timeZone: "UTC" })
const RECORDED = new Intl.DateTimeFormat(undefined, { day: "numeric", month: "short", year: "numeric" })

/** A register date is stored as strict YYYY-MM-DD; format it without letting a
 * browser timezone move the day. Anything else is shown as stored rather than
 * silently reinterpreted. */
export function formatKeyDate(value: string) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) return value
  const date = new Date(`${value}T00:00:00Z`)
  return Number.isNaN(date.getTime()) ? value : DAY.format(date)
}

export function keyDateKindLabel(kind: string) {
  return KIND_LABELS[kind] ?? "Date"
}

export function keyDateSummary(rows: KeyDateRegisterRow[]) {
  return { total: rows.length, unresolved: rows.filter((row) => !row.date).length }
}

export default function KeyDatesRegister({
  matterId,
  documents,
  onClose,
  onOpenDocument,
  onAsk,
}: {
  matterId: string
  documents: string[]
  onClose: () => void
  onOpenDocument: (name: string) => void
  onAsk?: () => void
}) {
  const [rows, setRows] = useState<KeyDateRegisterRow[]>()
  const [error, setError] = useState("")
  const [attempt, setAttempt] = useState(0)

  useEffect(() => {
    const controller = new AbortController()
    setRows(undefined)
    setError("")
    getKeyDates(matterId, controller.signal)
      .then(setRows)
      .catch((e: unknown) => {
        if (controller.signal.aborted) return
        setError(e instanceof Error ? e.message : "Could not read the key-date register.")
      })
    return () => controller.abort()
  }, [matterId, attempt])

  const summary = rows && keyDateSummary(rows)

  return (
    <section className="key-dates-register" aria-labelledby="key-dates-title">
      <header className="key-dates-head">
        <div className="key-dates-heading">
          <p className="key-dates-eyebrow">Matter register</p>
          <h1 id="key-dates-title">Key dates</h1>
          <p className="key-dates-lead">
            Dates and relative formulas recorded from this matter&rsquo;s documents during reviews, soonest first. Each
            entry records what a source says; it is not a confirmed deadline.
          </p>
        </div>
        <button type="button" className="ghost" onClick={onClose}>
          Back to conversation
        </button>
      </header>

      {error && (
        <div className="key-dates-state key-dates-error" role="alert">
          <p>{error}</p>
          <button type="button" onClick={() => setAttempt((value) => value + 1)}>
            Try again
          </button>
        </div>
      )}

      {!rows && !error && (
        <div className="key-dates-state" aria-busy="true">
          <span className="key-dates-spinner" aria-hidden />
          Reading the key-date register&hellip;
        </div>
      )}

      {rows && rows.length === 0 && (
        <div className="key-dates-state key-dates-empty">
          <h2>No key dates recorded yet</h2>
          <p>
            Reviews record renewals, termination-notice windows, expiries, payment milestones and deliverable deadlines
            here as they read the documents. An empty register means none have been recorded, not that the matter has no
            deadlines.
          </p>
          {onAsk && (
            <button type="button" className="primary" onClick={onAsk}>
              Ask about deadlines
            </button>
          )}
        </div>
      )}

      {rows && rows.length > 0 && summary && (
        <>
          <p className="key-dates-summary" aria-live="polite">
            {summary.total} {summary.total === 1 ? "entry" : "entries"}
            {summary.unresolved > 0 &&
              ` · ${summary.unresolved} ${summary.unresolved === 1 ? "entry has" : "entries have"} no calendar date`}
          </p>
          <ol className="key-dates-list">
            {rows.map((row) => (
              <KeyDateEntry key={row.id} row={row} documents={documents} onOpenDocument={onOpenDocument} />
            ))}
          </ol>
          <p className="key-dates-footnote">
            Check each entry against its quoted source before relying on it. Formula-only entries have no resolved
            calendar date; a resolved date is still recorded evidence, not an accepted or independently verified
            deadline.
          </p>
        </>
      )}
    </section>
  )
}

function KeyDateEntry({
  row,
  documents,
  onOpenDocument,
}: {
  row: KeyDateRegisterRow
  documents: string[]
  onOpenDocument: (name: string) => void
}) {
  const present = documents.includes(row.document)
  return (
    <li className={`key-dates-entry${row.date ? "" : " is-unresolved"}`}>
      <div className="key-dates-when">
        {row.date ? (
          <time dateTime={row.date}>{formatKeyDate(row.date)}</time>
        ) : (
          <span className="key-dates-no-date">No calendar date</span>
        )}
        <span className="key-dates-kind">{keyDateKindLabel(row.kind)}</span>
      </div>
      <div className="key-dates-body">
        <p className="key-dates-reason">{row.reason}</p>
        {row.formula && (
          <p className="key-dates-formula">
            <span>{row.date ? "Document formula" : "Unresolved formula"}</span> &ldquo;{row.formula}&rdquo;
          </p>
        )}
        <blockquote className="key-dates-quote">&ldquo;{row.quote}&rdquo;</blockquote>
        <dl className="key-dates-facts">
          <div>
            <dt>Document</dt>
            <dd>
              {present ? (
                <button type="button" className="linklike" onClick={() => onOpenDocument(row.document)}>
                  {row.document}
                </button>
              ) : (
                <>
                  {row.document} <span className="key-dates-missing">(not in this matter now)</span>
                </>
              )}
            </dd>
          </div>
          {row.party && (
            <div>
              <dt>Binds</dt>
              <dd>{row.party}</dd>
            </div>
          )}
          {row.noticeDays != null && (
            <div>
              <dt>Notice</dt>
              <dd>{row.noticeDays === 1 ? "1 day" : `${row.noticeDays} days`}</dd>
            </div>
          )}
          {row.actBy && (
            <div>
              <dt>Act by</dt>
              <dd>
                <time dateTime={row.actBy}>{formatKeyDate(row.actBy)}</time>
              </dd>
            </div>
          )}
          <div>
            <dt>Recorded</dt>
            <dd>
              {row.author} &middot; {RECORDED.format(new Date(row.recordedAt))}
            </dd>
          </div>
        </dl>
      </div>
    </li>
  )
}
