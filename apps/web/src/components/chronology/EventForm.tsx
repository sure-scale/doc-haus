import { useState } from "react"
import type { ChronologyAction, ChronologyDate, ChronologyEvent } from "../../api/ingest/chronology"
import ChronologyDialog from "./Dialog"
import { dateLabel, eventEvidence } from "./model"
export default function EventForm({
  event,
  events,
  mode,
  busy,
  error,
  onClose,
  onSave,
  onPreview,
}: {
  event?: ChronologyEvent
  events: ChronologyEvent[]
  mode: "add" | "edit" | "merge" | "split"
  busy: boolean
  error?: string
  onClose: () => void
  onSave: (action: ChronologyAction) => Promise<boolean>
  onPreview: (action: ChronologyAction) => Promise<ChronologyEvent[] | undefined>
}) {
  const [title, setTitle] = useState(event?.title ?? "")
  const [detail, setDetail] = useState(event?.detail ?? "")
  const [date, setDate] = useState<ChronologyDate>(event?.date ?? { original: "", precision: "undated" })
  const [other, setOther] = useState("")
  const [secondTitle, setSecondTitle] = useState("")
  const [firstAccounts, setFirstAccounts] = useState(
    event?.accounts.slice(0, Math.ceil(event.accounts.length / 2)).map((account) => account.id) ?? [],
  )
  const [preview, setPreview] = useState<ChronologyEvent[]>()
  const label = { add: "Add event", edit: "Edit event", merge: "Merge events", split: "Split event" }[mode]
  const action: ChronologyAction =
    mode === "add"
      ? { type: "add", title, detail, date }
      : mode === "edit"
        ? { type: "edit", eventId: event!.id, title, detail, date }
        : mode === "merge"
          ? { type: "merge", eventIds: [event!.id, other], title, detail, date }
          : {
              type: "split",
              eventId: event!.id,
              parts: [
                { title, detail, date, accountIds: firstAccounts },
                {
                  title: secondTitle,
                  detail: "",
                  date: event!.date,
                  accountIds: event!.accounts
                    .filter((account) => !firstAccounts.includes(account.id))
                    .map((account) => account.id),
                },
              ],
            }
  const valid = title.trim() && (mode !== "merge" || other) && (mode !== "split" || secondTitle.trim())
  return (
    <ChronologyDialog title={label} onClose={onClose}>
      <form
        onSubmit={async (submit) => {
          submit.preventDefault()
          if (!valid) return
          if ((mode === "merge" || mode === "split") && !preview) {
            setPreview(await onPreview(action))
            return
          }
          if (await onSave(action)) onClose()
        }}
        onChange={() => setPreview(undefined)}
      >
        {error && (
          <p role="alert" className="chronology-warning">
            {error}
          </p>
        )}
        {mode === "add" && (
          <p className="chronology-notice">
            Added by hand. No supporting source will be manufactured. This event remains visibly unsupported until
            evidence is supplied.
          </p>
        )}
        {mode === "edit" && (
          <p className="chronology-muted">
            The generated account and its evidence are kept. A material correction returns this event to unreviewed.
          </p>
        )}
        {mode === "merge" && (
          <label>
            Merge with
            <select required value={other} onChange={(change) => setOther(change.target.value)}>
              <option value="">Choose another event</option>
              {events
                .filter((item) => item.id !== event?.id && !item.hidden)
                .map((item) => (
                  <option key={item.id} value={item.id}>
                    {dateLabel(item.date)} · {item.title}
                  </option>
                ))}
            </select>
          </label>
        )}
        <label>
          {mode === "split" ? "First event" : "Event"}
          <input
            required
            maxLength={1000}
            autoFocus
            value={title}
            onChange={(change) => setTitle(change.target.value)}
          />
        </label>
        <label>
          Description
          <textarea rows={3} value={detail} onChange={(change) => setDetail(change.target.value)} />
        </label>
        <div className="chronology-form-grid">
          <label>
            Date as recorded
            <input
              value={date.original}
              onChange={(change) => setDate({ ...date, original: change.target.value })}
              placeholder="e.g. March 2026, by 3 April, or unknown"
            />
          </label>
          <label>
            Precision
            <select
              value={date.precision}
              onChange={(change) => {
                const precision = change.target.value as ChronologyDate["precision"]
                setDate({
                  ...date,
                  precision,
                  relative: precision === "relative" ? date.relative : undefined,
                  start: ["undated", "ambiguous"].includes(precision)
                    ? undefined
                    : precision === "month"
                      ? date.start?.slice(0, 7)
                      : precision === "year"
                        ? date.start?.slice(0, 4)
                        : date.start,
                  end: ["undated", "ambiguous"].includes(precision) ? undefined : date.end,
                })
              }}
            >
              <option value="undated">Undated</option>
              <option value="exact">Exact</option>
              <option value="approximate">Approximate</option>
              <option value="month">Month only</option>
              <option value="year">Year only</option>
              <option value="ambiguous">Ambiguous</option>
              <option value="relative">Relative</option>
            </select>
          </label>
          <label>
            Start / anchor (where supported)
            <input
              value={date.start ?? ""}
              onChange={(change) => setDate({ ...date, start: change.target.value || undefined })}
              placeholder="YYYY-MM-DD; preserve supplied precision"
            />
          </label>
          <label>
            End (where supported)
            <input
              value={date.end ?? ""}
              onChange={(change) =>
                setDate({
                  ...date,
                  end: change.target.value || undefined,
                  rangeKind: change.target.value ? (date.rangeKind ?? "uncertainty") : undefined,
                })
              }
            />
          </label>
          {date.end && (
            <label>
              Interval meaning
              <select
                value={date.rangeKind ?? "uncertainty"}
                onChange={(change) =>
                  setDate({ ...date, rangeKind: change.target.value as "duration" | "uncertainty" })
                }
              >
                <option value="uncertainty">Uncertainty about when</option>
                <option value="duration">Actual start-to-end period</option>
              </select>
            </label>
          )}
          <label>
            Time
            <input
              value={date.time ?? ""}
              onChange={(change) => setDate({ ...date, time: change.target.value || undefined })}
            />
          </label>
          <label>
            Timezone
            <input
              value={date.timezone ?? ""}
              onChange={(change) => setDate({ ...date, timezone: change.target.value || undefined })}
            />
          </label>
        </div>
        {mode === "split" && (
          <>
            <label>
              Second event
              <input required value={secondTitle} onChange={(change) => setSecondTitle(change.target.value)} />
            </label>
            <fieldset>
              <legend>Accounts to keep with the first event</legend>
              {event!.accounts.map((account) => (
                <label className="chronology-check" key={account.id}>
                  <input
                    type="checkbox"
                    checked={firstAccounts.includes(account.id)}
                    onChange={(change) =>
                      setFirstAccounts((current) =>
                        change.target.checked ? [...current, account.id] : current.filter((id) => id !== account.id),
                      )
                    }
                  />
                  {account.title} · {dateLabel(account.date)}
                </label>
              ))}
              <p className="chronology-muted">
                The other accounts move to the second event. Each resulting event can then be corrected separately.
              </p>
            </fieldset>
          </>
        )}
        {preview && (
          <section className="chronology-preview" aria-label="Resulting events">
            <h3>Preview resulting events</h3>
            {preview.map((item) => (
              <article key={item.id}>
                <strong>{item.title}</strong>
                <p>
                  {dateLabel(item.date)} · {item.accounts.length} accounts · {eventEvidence(item).length} sources
                </p>
                {item.note && <p>Notes: {item.note}</p>}
              </article>
            ))}
            <p>Previous records and source accounts are kept. Changed events need a fresh review.</p>
          </section>
        )}
        <div className="chronology-form-actions">
          <button type="button" onClick={onClose}>
            Cancel
          </button>
          <button type="submit" className="primary" disabled={!valid || busy}>
            {busy
              ? "Saving…"
              : (mode === "merge" || mode === "split") && !preview
                ? "Preview changes"
                : `Confirm ${mode === "merge" ? "merge" : mode === "split" ? "split" : "event"}`}
          </button>
        </div>
      </form>
    </ChronologyDialog>
  )
}
