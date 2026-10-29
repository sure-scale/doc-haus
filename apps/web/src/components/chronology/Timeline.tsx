import { useEffect, useId, useLayoutEffect, useMemo, useRef, useState } from "react"
import { createPortal } from "react-dom"
import type { ChronologyEvent } from "../../api/ingest/chronology"
import { dateLabel } from "./model"
import "./timeline.css"

const timestamp = (value?: string) => (value ? Date.parse(value) : NaN)
const DAY = 86400000
const lanes = [
  { id: "correspondence", label: "Correspondence" },
  { id: "documents", label: "Documents" },
  { id: "other", label: "Other events" },
] as const

type Lane = (typeof lanes)[number]["id"]
// Source kinds describe evidence, not necessarily the event itself. Keep mixed
// and unclassified sources together rather than inferring a type from a title.
function eventLane(event: ChronologyEvent): Lane {
  const kinds = event.accounts.flatMap((account) =>
    account.evidence.map((evidence): Lane => {
      const kind = evidence.sourceKind?.toLowerCase().trim() ?? ""
      const correspondence = /\b(e-?mail|correspondence|letter)\b/.test(kind)
      const document = /\b(document|contract|agreement|deed)\b/.test(kind)
      if (correspondence && document) return "other"
      if (correspondence) return "correspondence"
      if (document) return "documents"
      if (/\b(meeting|minutes|file note|attendance note|telephone|call)\b/.test(kind)) return "other"
      if (/\.(eml|msg)$/i.test(evidence.documentName)) return "correspondence"
      if (/\.(docx|pdf|xlsx)$/i.test(evidence.documentName)) return "documents"
      return "other"
    }),
  )
  if (kinds.length && kinds.every((kind) => kind === "correspondence")) return "correspondence"
  if (kinds.length && kinds.every((kind) => kind === "documents")) return "documents"
  return "other"
}

function eventMeaning(event: ChronologyEvent) {
  const precision: Record<string, string> = {
    exact: "Exact date",
    approximate: "Approximate date",
    month: "Month only",
    year: "Year only",
    relative: "Relative date",
  }
  const meanings = [precision[event.date.precision] ?? "Uncertain date"]
  for (const status of new Set(event.accounts.map((account) => account.status))) {
    if (status === "scheduled") meanings.push("Scheduled event")
    if (status === "required") meanings.push("Required event")
  }
  return meanings.join(" · ")
}

function periodMeaning(event: ChronologyEvent) {
  if (!event.date.end) return undefined
  const label = (value: string) => {
    const date = timestamp(value)
    if (!Number.isFinite(date)) return value
    const format: Intl.DateTimeFormatOptions = { year: "numeric", timeZone: "UTC" }
    if (!/^\d{4}$/.test(value)) format.month = "short"
    if (!/^\d{4}(?:-\d{2})?$/.test(value)) format.day = "numeric"
    return new Intl.DateTimeFormat(undefined, format).format(new Date(date))
  }
  const interval = `${label(event.date.start!)} – ${label(event.date.end)}`
  return event.date.rangeKind === "duration" ? `Known duration: ${interval}` : `Date uncertainty interval: ${interval}`
}

function LaneIcon({ lane }: { lane: Lane }) {
  return (
    <svg
      width="17"
      height="17"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.65"
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
    >
      {lane === "correspondence" ? (
        <>
          <rect x="3" y="5" width="18" height="14" rx="2" />
          <path d="m3 6 9 7 9-7" />
        </>
      ) : lane === "documents" ? (
        <>
          <path d="M14 3H5v18h14V8z" />
          <path d="M14 3v5h5M8 12h8M8 16h6" />
        </>
      ) : (
        <>
          <circle cx="9" cy="7" r="3" />
          <path d="M3 21v-3a6 6 0 0 1 12 0v3zM17 4a3 3 0 0 1 0 6M18 14a5 5 0 0 1 3 4v3h-3" />
        </>
      )}
    </svg>
  )
}

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
  const [canvasWidth, setCanvasWidth] = useState(0)
  const sectionRef = useRef<HTMLElement>(null)
  const canvasRef = useRef<HTMLDivElement>(null)
  const scrollRef = useRef<HTMLDivElement>(null)
  const clusterListRef = useRef<HTMLDivElement>(null)
  const clusterTriggerRef = useRef<HTMLButtonElement | null>(null)
  const pickerFocusedElement = useRef<HTMLElement | null>(null)
  const suppressPreviewFocus = useRef(false)
  const tooltipId = useId()
  const tooltipRef = useRef<HTMLDivElement>(null)
  const tooltipLeaveTimer = useRef<ReturnType<typeof setTimeout> | undefined>(undefined)
  const [tooltip, setTooltip] = useState<{ anchor: HTMLElement; ids: string[]; kind: "point" | "period" }>()
  const [tooltipPosition, setTooltipPosition] = useState<{ left: number; top: number }>()
  const cancelTooltipLeave = () => clearTimeout(tooltipLeaveTimer.current)
  const hideTooltip = () => {
    cancelTooltipLeave()
    pickerFocusedElement.current = null
    setTooltip(undefined)
    setTooltipPosition(undefined)
    setCluster((current) => (current.length ? [] : current))
  }
  const showTooltip = (anchor: HTMLElement, ids: string[], kind: "point" | "period" = "point", force = false) => {
    if (cluster.length && !force) return
    cancelTooltipLeave()
    clusterTriggerRef.current = anchor instanceof HTMLButtonElement ? anchor : null
    setTooltipPosition(undefined)
    setTooltip({ anchor, ids, kind })
  }
  const leaveTooltip = () => {
    cancelTooltipLeave()
    tooltipLeaveTimer.current = setTimeout(() => {
      if (
        !cluster.length &&
        document.activeElement !== tooltip?.anchor &&
        !tooltipRef.current?.contains(document.activeElement)
      )
        hideTooltip()
    }, 160)
  }
  useEffect(() => () => clearTimeout(tooltipLeaveTimer.current), [])
  useLayoutEffect(() => {
    hideTooltip()
  }, [selected])
  const dated = useMemo(
    () =>
      events.filter((event) => Number.isFinite(timestamp(event.date.start)) && event.date.precision !== "ambiguous"),
    [events],
  )
  const undated = events.filter((event) => !dated.includes(event))
  const occupiedLanes = useMemo(
    () => lanes.filter((lane) => dated.some((event) => eventLane(event) === lane.id)),
    [dated],
  )
  const clusterEvents = useMemo(() => dated.filter((event) => cluster.includes(event.id)), [dated, cluster])
  const tooltipEvents = useMemo(() => dated.filter((event) => tooltip?.ids.includes(event.id)), [dated, tooltip])
  useLayoutEffect(() => {
    const bubble = tooltipRef.current
    if (!tooltip) return
    if (!tooltip.anchor.isConnected || !tooltipEvents.length) {
      const focused = pickerFocusedElement.current
      const restore =
        bubble?.contains(document.activeElement) ||
        (focused && !focused.isConnected && document.activeElement === document.body)
      hideTooltip()
      if (restore) restoreClusterFocus()
      return
    }
    if (!bubble) return
    const scale = Number.parseFloat(getComputedStyle(document.body).zoom) || 1
    const viewport = window.visualViewport
    const left = viewport?.offsetLeft ?? 0
    const top = viewport?.offsetTop ?? 0
    const width = viewport?.width ?? window.innerWidth
    const height = viewport?.height ?? window.innerHeight
    const gap = 8
    bubble.style.maxWidth = `${Math.max(0, Math.min(320, (width - gap * 2) / scale))}px`
    bubble.style.maxHeight = `${Math.max(0, Math.min(420, (height - gap * 2) / scale))}px`
    const anchor = tooltip.anchor.getBoundingClientRect()
    const rect = bubble.getBoundingClientRect()
    const above = anchor.top - rect.height - gap
    const preferredTop = above >= top + gap ? above : anchor.bottom + gap
    setTooltipPosition({
      left:
        Math.max(
          left + gap,
          Math.min(anchor.left + anchor.width / 2 - rect.width / 2, left + width - rect.width - gap),
        ) / scale,
      top: Math.max(top + gap, Math.min(preferredTop, top + height - rect.height - gap)) / scale,
    })
    const escape = (event: KeyboardEvent) => {
      if (event.key !== "Escape" || event.defaultPrevented) return
      event.preventDefault()
      event.stopPropagation()
      const restore = bubble.contains(document.activeElement)
      hideTooltip()
      if (restore) restoreClusterFocus()
    }
    const outside = (event: PointerEvent) => {
      if (!(event.target instanceof Node) || bubble.contains(event.target) || tooltip.anchor.contains(event.target))
        return
      hideTooltip()
    }
    const dismissForGeometry = () => {
      const restore = bubble.contains(document.activeElement)
      hideTooltip()
      if (restore) restoreClusterFocus()
    }
    const scroll = (event: Event) => {
      if (event.target instanceof Node && bubble.contains(event.target)) return
      dismissForGeometry()
    }
    document.addEventListener("keydown", escape, true)
    document.addEventListener("pointerdown", outside, true)
    window.addEventListener("scroll", scroll, true)
    window.addEventListener("resize", dismissForGeometry)
    viewport?.addEventListener("scroll", dismissForGeometry)
    viewport?.addEventListener("resize", dismissForGeometry)
    return () => {
      document.removeEventListener("keydown", escape, true)
      document.removeEventListener("pointerdown", outside, true)
      window.removeEventListener("scroll", scroll, true)
      window.removeEventListener("resize", dismissForGeometry)
      viewport?.removeEventListener("scroll", dismissForGeometry)
      viewport?.removeEventListener("resize", dismissForGeometry)
    }
  }, [tooltip, tooltipEvents])
  const bounds = dated
    .flatMap((event) => [timestamp(event.date.start), timestamp(event.date.end)])
    .filter(Number.isFinite)
  const min = bounds.length ? Math.min(...bounds) : 0
  const max = bounds.length ? Math.max(...bounds) : DAY
  const span = Math.max(max - min, DAY)
  const hasDatedEvents = dated.length > 0
  useLayoutEffect(() => {
    const canvas = canvasRef.current
    if (!canvas) return
    const measure = () => setCanvasWidth(canvas.clientWidth)
    measure()
    if (typeof ResizeObserver === "undefined") return
    const observer = new ResizeObserver(measure)
    observer.observe(canvas)
    return () => observer.disconnect()
  }, [hasDatedEvents, zoom])
  const inset = canvasWidth > 0 ? Math.min(50, Math.max(5, 2200 / canvasWidth)) : 5
  const plotWidth = 100 - 2 * inset
  const position = (event: ChronologyEvent) => inset + ((timestamp(event.date.start) - min) / span) * plotWidth
  const row = (event: ChronologyEvent) => occupiedLanes.findIndex((lane) => lane.id === eventLane(event))
  const groups = useMemo(() => {
    const result: ChronologyEvent[][] = []
    for (const lane of lanes) {
      const ordered = dated
        .filter((event) => eventLane(event) === lane.id)
        .sort((a, b) => timestamp(a.date.start) - timestamp(b.date.start) || a.id.localeCompare(b.id))
      let group: ChronologyEvent[] = []
      for (const event of ordered) {
        const previous = group[group.length - 1]
        const distance = previous
          ? (((timestamp(event.date.start) - timestamp(previous.date.start)) / span) * canvasWidth * plotWidth) / 100
          : Infinity
        // Connected overlapping hit areas form one cluster. Every event retains
        // its date in the picker; no displayed representative overlaps a neighbour.
        if (!previous || distance >= 44) {
          group = []
          result.push(group)
        }
        group.push(event)
      }
    }
    return result
  }, [dated, span, canvasWidth, plotWidth])
  const selectedEvent = dated.find((event) => event.id === selected)
  const selectedStart = selectedEvent?.date.start
  useLayoutEffect(() => {
    const viewport = scrollRef.current
    const canvas = canvasRef.current
    if (!viewport || !canvas || !selectedStart || !viewport.clientWidth) return
    const x = ((inset + ((timestamp(selectedStart) - min) / span) * plotWidth) * canvas.clientWidth) / 100
    const left = viewport.scrollLeft
    const right = left + viewport.clientWidth
    if (x - 22 < left) viewport.scrollLeft = Math.max(0, x - 30)
    else if (x + 22 > right)
      viewport.scrollLeft = Math.min(viewport.scrollWidth - viewport.clientWidth, x + 30 - viewport.clientWidth)
  }, [selected, selectedStart, min, span, inset, plotWidth, canvasWidth, zoom])
  const restoreClusterFocus = () => {
    const trigger = clusterTriggerRef.current
    const target = trigger?.isConnected
      ? trigger
      : (canvasRef.current?.querySelector<HTMLButtonElement>('button[aria-pressed="true"]') ??
        scrollRef.current ??
        sectionRef.current?.querySelector<HTMLElement>("summary") ??
        sectionRef.current)
    suppressPreviewFocus.current = Boolean(target?.classList.contains("chronology-timeline-point"))
    target?.focus({ preventScroll: true })
  }
  useLayoutEffect(() => {
    const focused = pickerFocusedElement.current
    const lostPickerFocus = focused && !focused.isConnected && document.activeElement === document.body
    if (cluster.length && !clusterEvents.length) {
      pickerFocusedElement.current = null
      setCluster([])
      if (lostPickerFocus) restoreClusterFocus()
    } else if (lostPickerFocus) {
      pickerFocusedElement.current = null
      clusterListRef.current?.querySelector<HTMLButtonElement>("button")?.focus({ preventScroll: true })
    }
  }, [cluster, clusterEvents])
  useEffect(() => {
    if (cluster.length && tooltipPosition && !tooltipRef.current?.contains(document.activeElement))
      clusterListRef.current?.querySelector<HTMLButtonElement>("button")?.focus({ preventScroll: true })
  }, [cluster, tooltipPosition])
  const closeCluster = () => {
    hideTooltip()
    restoreClusterFocus()
  }
  const openPicker = (anchor: HTMLButtonElement, ids: string[]) => {
    clusterTriggerRef.current = anchor
    showTooltip(anchor, ids, "point", true)
    setCluster(ids)
  }
  const tickFormat = new Intl.DateTimeFormat(
    undefined,
    span > 150 * DAY
      ? { month: "short", year: "numeric", timeZone: "UTC" }
      : { day: "numeric", month: "short", timeZone: "UTC" },
  )
  const tickDates: number[] = []
  if (span > 150 * DAY) {
    const first = new Date(min)
    const last = new Date(max)
    const months = (last.getUTCFullYear() - first.getUTCFullYear()) * 12 + last.getUTCMonth() - first.getUTCMonth()
    const step = Math.max(1, Math.ceil((months + 1) / 6))
    let month = first.getUTCMonth()
    if (Date.UTC(first.getUTCFullYear(), month, 1) < min) {
      tickDates.push(min)
      month += 1
    }
    for (
      let date = Date.UTC(first.getUTCFullYear(), month, 1);
      date <= max;
      month += step, date = Date.UTC(first.getUTCFullYear(), month, 1)
    )
      tickDates.push(date)
  } else {
    const count = max === min ? 1 : Math.min(6, Math.floor(span / DAY) + 1)
    for (let index = 0; index < count; index += 1) tickDates.push(min + (span * index) / Math.max(1, count - 1))
  }
  const ticks = tickDates.map((date) => ({
    position: inset + ((date - min) / span) * plotWidth,
    label: tickFormat.format(new Date(date)),
  }))
  return (
    <section
      ref={sectionRef}
      tabIndex={-1}
      className="chronology-timeline chronology-swimlane-timeline"
      aria-label="Event timeline"
    >
      <h2 className="chronology-sr-only">Timeline of the displayed events</h2>
      {dated.length > 0 ? (
        <div className="chronology-swimlane-chart">
          <div className="chronology-swimlane-labels" aria-hidden="true">
            {occupiedLanes.map((lane) => (
              <div key={lane.id} className={`chronology-swimlane-label lane-${lane.id}`}>
                <LaneIcon lane={lane.id} />
                <span>{lane.label}</span>
              </div>
            ))}
          </div>
          <div
            ref={scrollRef}
            className="chronology-timeline-scroll"
            tabIndex={0}
            aria-label="Timeline; scroll horizontally when zoomed"
          >
            <div
              ref={canvasRef}
              className="chronology-timeline-canvas"
              style={{ width: `${100 * zoom}%`, height: `${42 + occupiedLanes.length * 50}px` }}
            >
              {ticks.map((tick, index) => (
                <div key={index} className="chronology-date-grid" style={{ left: `${tick.position}%` }}>
                  <span>{tick.label}</span>
                </div>
              ))}
              {occupiedLanes.map((lane, index) => (
                <div key={lane.id} className="chronology-swimlane-rule" style={{ top: `${91 + index * 50}px` }} />
              ))}
              {dated
                .filter(
                  (event) =>
                    Number.isFinite(timestamp(event.date.end)) &&
                    timestamp(event.date.end) >= timestamp(event.date.start),
                )
                .map((event) => (
                  <span
                    key={event.id}
                    className={`chronology-period lane-${eventLane(event)} ${event.date.rangeKind === "duration" ? "is-duration" : "is-uncertainty"}`}
                    style={{
                      left: `${position(event)}%`,
                      top: `${62 + row(event) * 50}px`,
                      width: `${Math.max(0.25, ((timestamp(event.date.end) - timestamp(event.date.start)) / span) * plotWidth)}%`,
                    }}
                    onPointerEnter={(pointer) => {
                      if (pointer.pointerType !== "touch") showTooltip(pointer.currentTarget, [event.id], "period")
                    }}
                    onPointerLeave={leaveTooltip}
                  />
                ))}
              {selectedEvent && (
                <div
                  className="chronology-selected-guide"
                  style={{ left: `${position(selectedEvent)}%` }}
                  aria-hidden="true"
                />
              )}
              {groups.map((group) => {
                const event = group.find((item) => item.id === selected) ?? group[0]
                const active = event.id === selected
                return (
                  <button
                    key={group.map((item) => item.id).join(":")}
                    type="button"
                    className={`chronology-timeline-point lane-${eventLane(event)}${active ? " is-selected" : ""}${event.date.precision !== "exact" ? " is-uncertain" : ""}${event.accounts.some((account) => account.status !== "reported") ? " is-scheduled" : ""}${group.length > 1 ? " is-cluster" : ""}`}
                    style={{ left: `${position(event)}%`, top: `${44 + row(event) * 50}px` }}
                    aria-pressed={active}
                    aria-label={
                      group.length > 1
                        ? `${group.length} clustered events near ${dateLabel(event.date)}`
                        : `${dateLabel(event.date)}: ${event.title}`
                    }
                    aria-haspopup="dialog"
                    aria-expanded={
                      tooltip?.kind === "point" &&
                      tooltip.anchor === clusterTriggerRef.current &&
                      tooltip.ids.length === group.length &&
                      tooltip.ids.every((id) => group.some((item) => item.id === id))
                    }
                    aria-controls={
                      tooltip?.kind === "point" &&
                      tooltip.ids.length === group.length &&
                      tooltip.ids.every((id) => group.some((item) => item.id === id))
                        ? tooltipId
                        : undefined
                    }
                    onPointerEnter={(pointer) => {
                      if (pointer.pointerType !== "touch")
                        showTooltip(
                          pointer.currentTarget,
                          group.map((item) => item.id),
                        )
                    }}
                    onPointerLeave={leaveTooltip}
                    onFocus={(focus) => {
                      if (suppressPreviewFocus.current) {
                        suppressPreviewFocus.current = false
                        return
                      }
                      showTooltip(
                        focus.currentTarget,
                        group.map((item) => item.id),
                      )
                    }}
                    onBlur={(blur) => {
                      if (!tooltipRef.current?.contains(blur.relatedTarget as Node | null)) hideTooltip()
                    }}
                    onKeyDown={(key) => {
                      if (
                        key.key === "ArrowDown" ||
                        (key.key === "Tab" && !key.shiftKey && tooltip?.anchor === key.currentTarget)
                      ) {
                        key.preventDefault()
                        openPicker(
                          key.currentTarget,
                          group.map((item) => item.id),
                        )
                      }
                    }}
                    onClick={(click) => {
                      clusterTriggerRef.current = click.currentTarget
                      if (group.length === 1) {
                        hideTooltip()
                        onSelect(event)
                      } else
                        openPicker(
                          click.currentTarget,
                          group.map((item) => item.id),
                        )
                    }}
                  >
                    <span className="chronology-point-mark" aria-hidden="true">
                      {group.length > 1 ? group.length : ""}
                    </span>
                    <span className="chronology-sr-only">{event.title}</span>
                  </button>
                )
              })}
              {selectedEvent && (
                <div
                  className={`chronology-selected-callout${position(selectedEvent) > 63 ? " opens-left" : ""}`}
                  style={{ left: `${position(selectedEvent)}%`, top: `${32 + row(selectedEvent) * 50}px` }}
                  aria-hidden="true"
                >
                  <strong>{dateLabel(selectedEvent.date)}</strong>
                  <span>{selectedEvent.title}</span>
                </div>
              )}
            </div>
          </div>
        </div>
      ) : (
        <p className="chronology-muted">No displayed events have a plottable date.</p>
      )}
      <div className="chronology-timeline-footer">
        <div className="chronology-timeline-key" aria-label="Timeline date markers">
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
          <span>
            <i className="chronology-scheduled-key" /> Scheduled / required
          </span>
        </div>
        {undated.length > 0 && (
          <details
            className="chronology-undated-events"
            onToggle={(toggle) => {
              if (toggle.currentTarget.open) setCluster([])
            }}
            onKeyDown={(key) => {
              if (key.key === "Escape") {
                key.preventDefault()
                key.currentTarget.open = false
                key.currentTarget.querySelector("summary")?.focus({ preventScroll: true })
              }
            }}
          >
            <summary>{undated.length} undated / ambiguous</summary>
            <div className="chronology-timeline-list">
              {undated.map((event) => (
                <button
                  key={event.id}
                  type="button"
                  onClick={(click) => {
                    click.currentTarget.closest("details")?.removeAttribute("open")
                    onSelect(event)
                  }}
                >
                  {event.title} · {dateLabel(event.date)}
                </button>
              ))}
            </div>
          </details>
        )}
        <div className="chronology-timeline-controls" aria-label="Timeline zoom">
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
      {tooltip &&
        tooltipEvents.length > 0 &&
        createPortal(
          <div
            ref={tooltipRef}
            id={tooltipId}
            role="dialog"
            aria-label="Timeline events"
            className="chronology-marker-popup"
            style={{
              left: tooltipPosition?.left ?? 0,
              top: tooltipPosition?.top ?? 0,
              visibility: tooltipPosition ? "visible" : "hidden",
            }}
            onPointerEnter={cancelTooltipLeave}
            onPointerLeave={leaveTooltip}
            onFocusCapture={(event) => {
              pickerFocusedElement.current = event.target as HTMLElement
            }}
            onBlurCapture={(event) => {
              if (!event.currentTarget.contains(event.relatedTarget as Node | null)) {
                pickerFocusedElement.current = null
                if (event.relatedTarget !== tooltip.anchor) hideTooltip()
              }
            }}
            onKeyDown={(key) => {
              if (key.key !== "Tab") return
              const buttons = [...key.currentTarget.querySelectorAll<HTMLButtonElement>("button")]
              const first = buttons[0],
                last = buttons[buttons.length - 1]
              if (key.shiftKey && document.activeElement === first) {
                key.preventDefault()
                closeCluster()
              } else if (!key.shiftKey && document.activeElement === last) {
                key.preventDefault()
                const controls = [
                  ...(sectionRef.current?.querySelectorAll<HTMLElement>(
                    'button:not(:disabled), summary, [tabindex="0"]',
                  ) ?? []),
                ].filter((element) => element.getClientRects().length && !element.closest("[inert]"))
                const next = controls[controls.indexOf(tooltip.anchor) + 1]
                hideTooltip()
                if (next) next.focus({ preventScroll: true })
                else restoreClusterFocus()
              }
            }}
          >
            <strong>{tooltipEvents.length > 1 ? `${tooltipEvents.length} nearby events` : "Timeline event"}</strong>
            <span className="chronology-marker-popup-hint">Choose an event to view its evidence.</span>
            <div
              ref={clusterListRef}
              role="group"
              aria-label={tooltipEvents.length > 1 ? "Clustered events" : "Event preview"}
              className="chronology-marker-popup-events"
            >
              {tooltipEvents.map((event) => (
                <button
                  key={event.id}
                  type="button"
                  className="chronology-marker-popup-event"
                  onClick={() => {
                    hideTooltip()
                    onSelect(event)
                  }}
                >
                  <span className="chronology-marker-popup-date">{dateLabel(event.date)}</span>
                  <span className="chronology-marker-popup-title">{event.title}</span>
                  <small>{eventMeaning(event)}</small>
                  {periodMeaning(event) && (
                    <small className="chronology-marker-popup-period">{periodMeaning(event)}</small>
                  )}
                </button>
              ))}
            </div>
            <button type="button" className="chronology-marker-popup-close" onClick={closeCluster}>
              Close event list
            </button>
          </div>,
          document.body,
        )}
    </section>
  )
}
