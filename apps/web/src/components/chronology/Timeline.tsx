import { useMemo, useState } from "react"
import type { ChronologyEvent } from "../../api/ingest/chronology"
import { dateLabel } from "./model"
const timestamp = (value?: string) => (value ? Date.parse(value) : NaN)
export default function ChronologyTimeline({
  events,
  selected,
  onSelect,
}: {
  events: ChronologyEvent[]
  selected?: string
  onSelect: (event: ChronologyEvent) => void
}) {
  const [zoom, setZoom] = useState(1)
  const [cluster, setCluster] = useState<string[]>([])
  const dated = events.filter(
    (event) => Number.isFinite(timestamp(event.date.start)) && event.date.precision !== "ambiguous",
  )
  const undated = events.filter((event) => !dated.includes(event))
  const bounds = dated
    .flatMap((event) => [timestamp(event.date.start), timestamp(event.date.end)])
    .filter(Number.isFinite)
  const min = Math.min(...bounds),
    max = Math.max(...bounds)
  const span = Math.max(max - min, 86400000)
  const groups = useMemo(() => {
    const groups = new Map<number, ChronologyEvent[]>()
    for (const event of dated) {
      const bucket = Math.floor(((timestamp(event.date.start) - min) / span) * 32 * zoom)
      groups.set(bucket, [...(groups.get(bucket) ?? []), event])
    }
    return [...groups.values()]
  }, [events, min, span, zoom])
  return (
    <section className="chronology-timeline" aria-label="Event timeline">
      <div className="chronology-timeline-heading">
        <span>
          Timeline <small>of the displayed events</small>
        </span>
        <div className="chronology-inline">
          <button
            type="button"
            aria-label="Zoom out timeline"
            disabled={zoom === 1}
            onClick={() => setZoom((value) => Math.max(1, value / 2))}
          >
            −
          </button>
          <button
            type="button"
            onClick={() => {
              setZoom(1)
              setCluster([])
            }}
          >
            Fit
          </button>
          <button
            type="button"
            aria-label="Zoom in timeline"
            disabled={zoom === 8}
            onClick={() => setZoom((value) => Math.min(8, value * 2))}
          >
            +
          </button>
        </div>
      </div>
      {dated.length > 0 ? (
        <div className="chronology-timeline-scroll" tabIndex={0} aria-label="Timeline; scroll horizontally when zoomed">
          <div className="chronology-timeline-canvas" style={{ width: `${100 * zoom}%` }}>
            <div className="chronology-axis-labels">
              <span>
                {dateLabel(dated.reduce((a, b) => (timestamp(a.date.start) < timestamp(b.date.start) ? a : b)).date)}
              </span>
              <span>
                {dateLabel(dated.reduce((a, b) => (timestamp(a.date.start) > timestamp(b.date.start) ? a : b)).date)}
              </span>
            </div>
            <div className="chronology-timeline-axis" />
            {dated
              .filter((event) => event.date.end)
              .map((event) => (
                <span
                  key={event.id}
                  className={`chronology-period ${event.date.rangeKind === "duration" ? "is-duration" : "is-uncertainty"}`}
                  style={{
                    left: `${3 + ((timestamp(event.date.start) - min) / span) * 94}%`,
                    width: `${Math.max(0.25, ((timestamp(event.date.end) - timestamp(event.date.start)) / span) * 94)}%`,
                  }}
                  title={`${event.date.rangeKind === "duration" ? "Known duration" : "Date uncertainty interval"}: ${dateLabel(event.date)}`}
                />
              ))}
            {groups.map((group) => {
              const event = group[0]
              return (
                <button
                  key={group.map((item) => item.id).join(":")}
                  type="button"
                  className={`chronology-timeline-point${group.some((item) => item.id === selected) ? " is-selected" : ""}${event.date.precision !== "exact" ? " is-uncertain" : ""}${event.accounts.some((account) => account.status !== "reported") ? " is-scheduled" : ""}`}
                  style={{ left: `${3 + ((timestamp(event.date.start) - min) / span) * 94}%` }}
                  aria-label={
                    group.length > 1
                      ? `${group.length} clustered events near ${dateLabel(event.date)}`
                      : `${dateLabel(event.date)}: ${event.title}`
                  }
                  title={
                    group.length > 1
                      ? `${group.length} events — select to inspect`
                      : `${dateLabel(event.date)} — ${event.title}`
                  }
                  onClick={() => {
                    if (group.length === 1) onSelect(event)
                    setCluster(group.length > 1 ? group.map((item) => item.id) : [])
                  }}
                >
                  {group.length > 1 ? group.length : <span className="chronology-sr-only">{event.title}</span>}
                </button>
              )
            })}
          </div>
        </div>
      ) : (
        <p className="chronology-muted">No displayed events have a plottable date.</p>
      )}
      <div className="chronology-timeline-key">
        <span>
          <i className="chronology-dot" /> Exact
        </span>
        <span>
          <i className="chronology-dot is-uncertain" /> Partial / uncertain
        </span>
        <span>
          <i className="chronology-duration-key" /> Known period
        </span>
        <span>
          <i className="chronology-uncertainty-key" /> Uncertainty interval
        </span>
        {undated.length > 0 && (
          <details>
            <summary>{undated.length} undated / ambiguous</summary>
            <div className="chronology-timeline-list">
              {undated.map((event) => (
                <button key={event.id} type="button" onClick={() => onSelect(event)}>
                  {event.title} · {dateLabel(event.date)}
                </button>
              ))}
            </div>
          </details>
        )}
      </div>
      {cluster.length > 0 && (
        <div className="chronology-timeline-cluster" aria-label="Clustered events">
          {events
            .filter((event) => cluster.includes(event.id))
            .map((event) => (
              <button
                key={event.id}
                type="button"
                onClick={() => {
                  onSelect(event)
                  setCluster([])
                }}
              >
                {dateLabel(event.date)} · {event.title}
              </button>
            ))}
          <button type="button" onClick={() => setCluster([])}>
            Close cluster
          </button>
        </div>
      )}
    </section>
  )
}
