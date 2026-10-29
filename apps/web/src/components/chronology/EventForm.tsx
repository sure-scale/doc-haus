import { useEffect, useMemo, useState } from "react"
import type { ChronologyAction, ChronologyDate, ChronologyEvent } from "../../api/ingest/chronology"
import JurisdictionSelect from "../JurisdictionSelect"
import ChronologyDialog from "./Dialog"
import DateInput, { useChronologyDateFormat } from "./DateInput"
import { formatDateInput, type DateInputFormat } from "./date-input"
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
  const format = useChronologyDateFormat()
  const [title, setTitle] = useState(event?.title ?? "")
  const [detail, setDetail] = useState(event?.detail ?? "")
  const [date, setDate] = useState<ChronologyDate>(event?.date ?? { original: "", precision: "undated" })
  const [dateValid, setDateValid] = useState(true)
  const [adjustmentsOpened, setAdjustmentsOpened] = useState(false)
  const [other, setOther] = useState("")
  const [secondTitle, setSecondTitle] = useState("")
  const [firstAccounts, setFirstAccounts] = useState(
    event?.accounts.slice(0, Math.ceil(event.accounts.length / 2)).map((account) => account.id) ?? [],
  )
  const [preview, setPreview] = useState<{ signature: string; events: ChronologyEvent[] }>()
  const transformation = mode === "merge" || mode === "split"
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
  // A preview describes both these edits and the source records being changed.
  // A refreshed source must be previewed again before it can be combined.
  const sourceIds = mode === "merge" ? [event!.id, other] : mode === "split" ? [event!.id] : []
  const signature = JSON.stringify({ action, sources: events.filter((item) => sourceIds.includes(item.id)) })
  const currentPreview = preview?.signature === signature ? preview.events : undefined
  const valid = Boolean(
    title.trim() && dateValid && (mode !== "merge" || other) && (mode !== "split" || secondTitle.trim()),
  )
  const dateFields = (
    <EventDateFields date={date} onChange={setDate} format={format} disabled={busy} onValidityChange={setDateValid} />
  )
  const titleField = (
    <label>
      {mode === "split" ? "First event" : "Event"}
      <input
        required
        maxLength={1000}
        disabled={busy}
        autoFocus={mode !== "merge"}
        value={title}
        onChange={(change) => setTitle(change.target.value)}
      />
    </label>
  )
  const descriptionField = (
    <label>
      Description
      <textarea rows={3} disabled={busy} value={detail} onChange={(change) => setDetail(change.target.value)} />
    </label>
  )
  return (
    <ChronologyDialog title={label} onClose={onClose}>
      <form
        onSubmit={async (submit) => {
          submit.preventDefault()
          if (!valid) return
          if (transformation && !currentPreview) {
            const result = await onPreview(action)
            if (result) setPreview({ signature, events: result })
            return
          }
          if (await onSave(action)) onClose()
        }}
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
          <>
            <p>
              Combine <strong>{event!.title}</strong> with another event. Sources and previous records are kept.
            </p>
            <div>
              <span>Merge with</span>
              <JurisdictionSelect
                jurisdictions={events
                  .filter((item) => item.id !== event?.id && !item.hidden && !item.supersededBy?.length)
                  .map((item) => ({ code: item.id, name: `${dateLabel(item.date)} · ${item.title}` }))}
                selected={other ? [other] : []}
                onChange={(codes) => setOther(codes[0] ?? "")}
                multiple={false}
                label="Merge with"
                searchLabel="Search events to merge"
                emptyLabel="Choose another event"
                noResultsLabel="No matching events."
                showCode={false}
                disabled={busy}
              />
            </div>
            <details
              onToggle={(event) => {
                if (event.currentTarget.open) setAdjustmentsOpened(true)
              }}
            >
              <summary>Adjust resulting event</summary>
              {titleField}
              {descriptionField}
              {adjustmentsOpened && dateFields}
            </details>
          </>
        )}
        {mode === "split" && (
          <>
            <p>Name the two resulting events, then choose which source accounts stay with the first.</p>
            <div className="chronology-form-grid">
              {titleField}
              <label>
                Second event
                <input
                  required
                  disabled={busy}
                  value={secondTitle}
                  onChange={(change) => setSecondTitle(change.target.value)}
                />
              </label>
            </div>
            <fieldset>
              <legend>Sources for the first event</legend>
              {event!.accounts.map((account) => (
                <label className="chronology-check" key={account.id}>
                  <input
                    type="checkbox"
                    disabled={busy}
                    checked={firstAccounts.includes(account.id)}
                    onChange={(change) =>
                      setFirstAccounts((current) =>
                        change.target.checked ? [...current, account.id] : current.filter((id) => id !== account.id),
                      )
                    }
                  />
                  <span>
                    {account.title} · {dateLabel(account.date)}
                    {account.evidence.length > 0 && (
                      <span className="chronology-muted">
                        {" "}
                        · {account.evidence.map((evidence) => evidence.documentName).join(", ")}
                      </span>
                    )}
                  </span>
                </label>
              ))}
              <p className="chronology-muted">
                Unchecked sources move to the second event. Every source account is retained. An event without sources
                remains visibly unsupported.
              </p>
            </fieldset>
            <details
              onToggle={(event) => {
                if (event.currentTarget.open) setAdjustmentsOpened(true)
              }}
            >
              <summary>Adjust first event description or date</summary>
              {descriptionField}
              {adjustmentsOpened && dateFields}
              <p className="chronology-muted">
                The second event keeps the original date. Each event can be corrected separately after the split.
              </p>
            </details>
          </>
        )}
        {!transformation && (
          <>
            {titleField}
            {descriptionField}
            {dateFields}
          </>
        )}
        {transformation && !dateValid && (
          <p className="chronology-warning">
            Complete or correct the date under the adjustment options before continuing.
          </p>
        )}
        {currentPreview && (
          <section className="chronology-preview" aria-label="Resulting events">
            <h3>Preview resulting events</h3>
            {currentPreview.map((item) => (
              <article key={item.id}>
                <strong>{item.title}</strong>
                <p>
                  {dateLabel(item.date)} · {item.accounts.length} accounts · {eventEvidence(item).length} sources
                </p>
                {item.detail && <p>{item.detail}</p>}
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
              : transformation && !currentPreview
                ? "Preview changes"
                : `Confirm ${mode === "merge" ? "merge" : mode === "split" ? "split" : "event"}`}
          </button>
        </div>
      </form>
    </ChronologyDialog>
  )
}

function EventDateFields({
  date,
  onChange,
  format,
  disabled,
  onValidityChange,
}: {
  date: ChronologyDate
  onChange: (date: ChronologyDate) => void
  format: DateInputFormat
  disabled: boolean
  onValidityChange: (valid: boolean) => void
}) {
  const [startValid, setStartValid] = useState(true)
  const [endValid, setEndValid] = useState(true)
  const [customWording, setCustomWording] = useState(false)
  const supported = !["undated", "ambiguous"].includes(date.precision)
  const partialDate = (value?: string) => (value?.length === 4 ? "year" : value?.length === 7 ? "month" : "day")
  const precision =
    date.precision === "month"
      ? "month"
      : date.precision === "year"
        ? "year"
        : ["approximate", "relative"].includes(date.precision)
          ? partialDate(date.start)
          : "day"
  const endPrecision =
    ["approximate", "relative"].includes(date.precision) && date.end ? partialDate(date.end) : precision
  const uncertainty = ["undated", "ambiguous", "approximate", "relative"].includes(date.precision)
  const ordered = !date.end || !date.start || date.end >= date.start
  useEffect(() => {
    onValidityChange(!supported || (startValid && endValid && ordered))
  }, [supported, startValid, endValid, ordered, onValidityChange])
  const zones = useMemo(() => {
    const known = ["UTC", ...(typeof Intl.supportedValuesOf === "function" ? Intl.supportedValuesOf("timeZone") : [])]
    const options = date.timezone && !known.includes(date.timezone) ? [date.timezone, ...known] : known
    return options.map((zone) => ({
      code: zone,
      name: `${zone.replace(/_/g, " ")}${known.includes(zone) ? "" : " (as recorded)"}`,
    }))
  }, [date.timezone])
  const words = (
    <label>
      Date as recorded
      <input
        disabled={disabled}
        value={date.original}
        onChange={(change) => {
          setCustomWording(true)
          onChange({ ...date, original: change.target.value })
        }}
        placeholder="Keep the source's wording, such as by 3 April or date unknown"
      />
    </label>
  )
  return (
    <>
      <div className="chronology-form-grid">
        <label>
          Precision
          <select
            disabled={disabled}
            value={date.precision}
            onChange={(change) => {
              const next = change.target.value as ChronologyDate["precision"]
              onChange({
                ...date,
                precision: next,
                relative: next === "relative" ? date.relative : undefined,
                start: ["undated", "ambiguous"].includes(next)
                  ? undefined
                  : next === "month"
                    ? date.start?.slice(0, 7)
                    : next === "year"
                      ? date.start?.slice(0, 4)
                      : date.start,
                end: ["undated", "ambiguous"].includes(next)
                  ? undefined
                  : next === "month"
                    ? date.end?.slice(0, 7)
                    : next === "year"
                      ? date.end?.slice(0, 4)
                      : date.end,
                rangeKind: ["undated", "ambiguous"].includes(next) ? undefined : date.rangeKind,
              })
            }}
          >
            <option value="undated">Date unknown</option>
            <option value="exact">Exact date</option>
            <option value="approximate">Approximate date</option>
            <option value="month">Month only</option>
            <option value="year">Year only</option>
            <option value="ambiguous">Ambiguous date</option>
            <option value="relative">Relative date</option>
          </select>
        </label>
        {supported && (
          <DateInput
            label={
              date.precision === "relative"
                ? "Calculated date (if supported)"
                : precision === "month"
                  ? "Month"
                  : precision === "year"
                    ? "Year"
                    : date.precision === "approximate"
                      ? "Approximate date"
                      : "Date"
            }
            value={date.start}
            format={format}
            precision={precision}
            required={date.precision !== "relative" && date.precision !== "approximate"}
            disabled={disabled}
            onValidityChange={setStartValid}
            onChange={(start) =>
              onChange({
                ...date,
                start,
                original:
                  !customWording && ["exact", "month", "year"].includes(date.precision)
                    ? formatDateInput(start, format, precision)
                    : date.original || formatDateInput(start, format, precision),
              })
            }
          />
        )}
      </div>
      {uncertainty && words}
      <details>
        <summary>Date details</summary>
        {!uncertainty && words}
        {supported && (
          <div className="chronology-form-grid">
            <DateInput
              label="End date (if applicable)"
              value={date.end}
              format={format}
              precision={endPrecision}
              disabled={disabled}
              onValidityChange={setEndValid}
              onChange={(end) =>
                onChange({ ...date, end, rangeKind: end ? (date.rangeKind ?? "uncertainty") : undefined })
              }
            />
            {date.end && (
              <label>
                Interval meaning
                <select
                  disabled={disabled}
                  value={date.rangeKind ?? "uncertainty"}
                  onChange={(change) =>
                    onChange({ ...date, rangeKind: change.target.value as "duration" | "uncertainty" })
                  }
                >
                  <option value="uncertainty">Uncertainty about when</option>
                  <option value="duration">Actual start-to-end period</option>
                </select>
              </label>
            )}
          </div>
        )}
        {!ordered && <p className="chronology-warning">The end date must be on or after the start date.</p>}
        {date.time && !/^\d{2}:\d{2}(:\d{2})?$/.test(date.time) && (
          <p className="chronology-muted">Time as recorded: {date.time}</p>
        )}
        <div className="chronology-form-grid">
          <label>
            Time
            <input
              type="time"
              step={date.time?.split(":").length === 3 ? 1 : 60}
              disabled={disabled}
              value={date.time && /^\d{2}:\d{2}(:\d{2})?$/.test(date.time) ? date.time : ""}
              onChange={(change) => onChange({ ...date, time: change.target.value || undefined })}
            />
          </label>
          <div>
            <span>Timezone</span>
            <JurisdictionSelect
              jurisdictions={zones}
              selected={date.timezone ? [date.timezone] : []}
              onChange={(values) => onChange({ ...date, timezone: values[0] })}
              multiple={false}
              label="Timezone"
              searchLabel="Search timezones"
              emptyLabel="No timezone recorded"
              noResultsLabel="No matching timezones."
              showCode={false}
              allowClear
              disabled={disabled}
            />
          </div>
        </div>
        <p className="chronology-muted">
          Keep the source timezone when known. No timezone is inferred from the matter or your device.
        </p>
      </details>
    </>
  )
}
