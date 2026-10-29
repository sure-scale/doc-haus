import { useEffect, useLayoutEffect, useMemo, useRef, useState, type CSSProperties, type FocusEvent as ReactFocusEvent, type KeyboardEvent as ReactKeyboardEvent, type ReactNode, type PointerEvent as ReactPointerEvent, type RefObject } from "react"
import { addDays } from "./data"

// The dashboard's chart marks, drawn as inline SVG and HTML. Marks paint in
// currentColor, so one colour set on each chart's root tints them all (see
// styles/dashboard.css); text keeps the text tokens. Sizes are proportional to
// the data: a zero draws nothing and a non-zero value gets at most a 2px floor,
// so small values never look substantial. Every chart has an accessible name,
// and keyboard focus shows the same exact values as hover.

type Tip = { x: number; y: number; content: ReactNode }

// Touch selection survives pointer leave after the finger lifts. An outside
// tap or Escape dismisses it without changing the chart's keyboard focus.
function useDismiss(ref: RefObject<HTMLElement | null>, visible: boolean, dismiss: () => void) {
  useEffect(() => {
    if (!visible) return
    const outside = (event: PointerEvent) => { if (!ref.current?.contains(event.target as Node)) dismiss() }
    const escape = (event: KeyboardEvent) => { if (event.key === "Escape") dismiss() }
    document.addEventListener("pointerdown", outside)
    document.addEventListener("keydown", escape)
    return () => {
      document.removeEventListener("pointerdown", outside)
      document.removeEventListener("keydown", escape)
    }
  }, [ref, visible, dismiss])
}

// One hover, focus or touch tooltip per chart, inside the chart's own box.
function useTip() {
  const box = useRef<HTMLDivElement>(null)
  const [tip, setTip] = useState<Tip>()
  const show = (event: { clientX: number; clientY: number } | DOMRect, content: ReactNode) => {
    const bounds = box.current?.getBoundingClientRect()
    if (!bounds) return
    const x = "clientX" in event ? event.clientX : event.left + event.width / 2
    const y = "clientY" in event ? event.clientY : event.top
    setTip({ x: x - bounds.left, y: y - bounds.top, content })
  }
  const hide = () => setTip(undefined)
  useDismiss(box, !!tip, hide)
  const view = tip && <TipBox x={tip.x} y={tip.y}>{tip.content}</TipBox>
  return { box, show, hide, leave: (event: ReactPointerEvent) => { if (event.pointerType !== "touch") hide() }, view }
}

// A tooltip centred over its point, shifted in so it stays inside the dashboard.
function TipBox({ x, y, children }: { x: number; y: number; children: ReactNode }) {
  const ref = useRef<HTMLDivElement>(null)
  useLayoutEffect(() => {
    const element = ref.current
    if (!element) return
    element.style.removeProperty("--tip-shift")
    const tip = element.getBoundingClientRect()
    const area = element.closest(".dash")?.getBoundingClientRect() ?? { left: 0, right: document.documentElement.clientWidth }
    const shift = tip.right > area.right ? area.right - tip.right : tip.left < area.left ? area.left - tip.left : 0
    if (shift) element.style.setProperty("--tip-shift", `${shift}px`)
  })
  return <div ref={ref} className="dash-tip" role="tooltip" style={{ left: x, top: y }}>{children}</div>
}

function useSize<T extends HTMLElement>() {
  const ref = useRef<T>(null)
  const [size, setSize] = useState({ width: 0, height: 0 })
  useLayoutEffect(() => {
    const element = ref.current
    if (!element) return
    setSize({ width: element.clientWidth, height: element.clientHeight })
    const observer = new ResizeObserver(([entry]) => setSize({ width: entry.contentRect.width, height: entry.contentRect.height }))
    observer.observe(element)
    return () => observer.disconnect()
  }, [])
  return [ref, size] as const
}

// One tab stop per chart: the arrow keys (plus Home and End) move focus
// between its marks, so a long series doesn't flood the tab order. `moves`
// maps a key to an index step.
const ALONG: Record<string, number> = { ArrowLeft: -1, ArrowRight: 1 }
function useRoving(count: number, initial: number, moves = ALONG) {
  const marks = useRef<(HTMLElement | SVGElement | null)[]>([])
  const [current, setCurrent] = useState<number>()
  const at = Math.max(0, Math.min(count - 1, current ?? initial))
  function keyDown(event: ReactKeyboardEvent) {
    const next = event.key === "Home" ? 0 : event.key === "End" ? count - 1 : event.key in moves ? at + moves[event.key] : undefined
    if (next === undefined) return
    event.preventDefault()
    marks.current[Math.max(0, Math.min(count - 1, next))]?.focus()
  }
  return (index: number, onFocus: (element: Element) => void) => ({
    ref: (element: HTMLElement | SVGElement | null) => { marks.current[index] = element },
    tabIndex: index === at ? 0 : -1,
    onKeyDown: keyDown,
    onFocus: (event: ReactFocusEvent) => { setCurrent(index); onFocus(event.currentTarget) },
  })
}

// A bar's height as its share of the chart's maximum, with a 2px floor so a
// small non-zero value stays visible. Callers draw no bar for a zero.
const barHeight = (value: number, max: number) => `max(2px, ${(value / max) * 100}%)`

// A stat tile's trend: a 2px line over a faint flat area down to zero, its
// latest point marked.
export function Sparkline({ values, label }: { values: number[]; label: string }) {
  const max = Math.max(...values, 0)
  if (values.length < 2 || max <= 0) return <div className="dash-spark is-flat" role="img" aria-label={`${label}: no activity yet`} />
  const width = 120, height = 40, base = height - 3, step = width / (values.length - 1)
  const points = values.map((value, index) => [index * step, base - (value / max) * (base - 5)] as const)
  const line = points.map(([x, y], index) => `${index ? "L" : "M"}${x.toFixed(1)} ${y.toFixed(1)}`).join(" ")
  const [lastX, lastY] = points[points.length - 1]
  return (
    <svg className="dash-spark" viewBox={`0 0 ${width} ${height}`} preserveAspectRatio="none" role="img" aria-label={label}>
      <path className="dash-spark-area" d={`${line} L${width} ${base} L0 ${base} Z`} fill="currentColor" fillOpacity=".1" />
      <path className="dash-spark-line" d={line} fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" vectorEffect="non-scaling-stroke" />
      {/* A near-zero-length round-capped stroke: it stays a circle where the
          stretched viewBox would squash a <circle> into an oval. */}
      <path className="dash-spark-end" d={`M${lastX.toFixed(1)} ${lastY.toFixed(1)}h0.01`} fill="none" stroke="currentColor" strokeWidth="8" strokeLinecap="round" vectorEffect="non-scaling-stroke" />
    </svg>
  )
}

// Spread behind an average: each item shares the available width, with its
// average shown as a labelled reference rule. Every item is drawn; bars thin
// out as the count grows. Each slot is a hit target even when its value is zero.
export function SpreadBars({ values, average, label, format }: { values: number[]; average: number; label: string; format: (value: number) => string }) {
  const { box, show, hide, leave, view } = useTip()
  const mark = useRoving(values.length, 0)
  const max = Math.max(...values, average, 0)
  if (!values.length || max <= 0) return <div className="dash-spark is-flat" role="img" aria-label={`${label}: nothing yet`} />
  return (
    <div className="dash-spread" ref={box} role="group" aria-label={label} style={{ "--bars": values.length } as CSSProperties} onPointerLeave={leave}>
      {values.map((value, index) => {
        const text = format(value)
        return (
          <span
            key={index}
            className={`dash-spread-slot${value > 0 ? "" : " is-zero"}`}
            role="img"
            aria-label={`${text}, ${index + 1} of ${values.length}`}
            {...mark(index, (element) => show(element.getBoundingClientRect(), text))}
            onBlur={hide}
            onPointerEnter={(event) => show(event.currentTarget.getBoundingClientRect(), text)}
            onPointerDown={(event) => show(event.currentTarget.getBoundingClientRect(), text)}
          >
            {value > 0 && <span className="dash-spread-bar" style={{ height: barHeight(value, max) }} />}
          </span>
        )
      })}
      <span className="dash-spread-average" style={{ bottom: `${(average / max) * 100}%` }} aria-hidden="true"><span className="dash-spread-average-label">avg {format(average)}</span></span>
      {view}
    </div>
  )
}

// The week view: a bar per day on a full-height track, as tall as its share
// of the busiest day. The track is the hit target, so a zero day still has a
// tooltip; today's label is emphasised. A day's optional `detail` (say, its
// top matters) joins the tooltip and the accessible name.
export function WeekBars({ days, format }: { days: { day: number; cost: number; label: string; today: boolean; detail?: string }[]; format: (value: number) => string }) {
  const { box, show, hide, leave, view } = useTip()
  const mark = useRoving(days.length, days.length - 1)
  const max = Math.max(...days.map((row) => row.cost), 0)
  return (
    <div className="dash-week" ref={box} onPointerLeave={leave}>
      <ol className="dash-week-bars" aria-label="Spend on each of the last seven days">
        {days.map((row, index) => {
          const text = `${row.today ? "Today" : row.label}: ${format(row.cost)}`
          const tip = row.detail ? <>{text}<span className="dash-tip-detail">{row.detail}</span></> : text
          return (
            <li key={row.day} className={`dash-week-day${row.today ? " is-today" : ""}`}>
              <span
                className="dash-week-track"
                role="img"
                aria-label={row.detail ? `${text}. ${row.detail}` : text}
                {...mark(index, (element) => show(element.getBoundingClientRect(), tip))}
                onBlur={hide}
                onPointerEnter={(event) => show(event.currentTarget.getBoundingClientRect(), tip)}
                onPointerDown={(event) => show(event.currentTarget.getBoundingClientRect(), tip)}
              >
                {row.cost > 0 && <span className="dash-week-fill" style={{ height: barHeight(row.cost, max) }} />}
              </span>
              <span className="dash-week-label" aria-hidden="true">{row.today ? "Today" : row.label}</span>
            </li>
          )
        })}
      </ol>
      {view}
    </div>
  )
}

// Shades for the calendar: 0 is no work; 1–4 rank a day among the days with
// any work (cut at the quartiles of their counts), so one very busy day
// doesn't wash every other day out to the palest step. The busiest day is
// always the darkest.
function shades(counts: number[]) {
  const busy = counts.filter((count) => count > 0).sort((a, b) => a - b)
  const max = busy[busy.length - 1] ?? 0
  const cuts = [0.25, 0.5, 0.75].map((share) => busy[Math.max(0, Math.ceil(share * busy.length) - 1)] ?? 0)
  return (count: number) => (count <= 0 ? 0 : count >= max ? 4 : 1 + cuts.filter((cut) => count > cut).length)
}

// Calendar keys: a column is a week, so left and right step a week and up
// and down a day.
const GRID: Record<string, number> = { ArrowLeft: -7, ArrowRight: 7, ArrowUp: -1, ArrowDown: 1 }

// A year of days as a calendar grid: one column per week (Monday first), one
// cell per day, darker for busier days. A 24px target surrounds each mark;
// older weeks scroll out of a narrow card. Keyboard focus starts on today.
export function ActivityCalendar({ counts, first, describe, label }: { counts: number[]; first: number; describe: (day: number, count: number) => string; label?: string }) {
  const { box, show, hide, leave, view } = useTip()
  const mark = useRoving(counts.length, counts.length - 1, GRID)
  const [scroller, { width: room }] = useSize<HTMLDivElement>()
  const top = 18, left = 30
  const weeks = Math.ceil(counts.length / 7)
  const pitch = 24
  const cell = pitch * 0.8
  // Start scrolled to today, as the newest weeks matter most on a narrow screen.
  useLayoutEffect(() => {
    const element = scroller.current
    if (element) element.scrollLeft = element.scrollWidth
  }, [counts.length, pitch])
  // Labels and shades are worked out once per data change, not on every
  // tooltip move.
  const days = useMemo(() => {
    const shade = shades(counts)
    return counts.map((count, index) => ({ level: shade(count), text: describe(addDays(first, index), count) }))
  }, [counts, first, describe])
  const max = Math.max(...counts, 0)
  const height = top + 7 * pitch
  const months: { x: number; label: string }[] = []
  for (let week = 0; week < weeks; week++) {
    const date = new Date(addDays(first, week * 7))
    const label = date.toLocaleDateString("en-GB", { month: "short" })
    if (!months.length || months[months.length - 1].label !== label) months.push({ x: week * pitch, label })
  }
  if (months.length > 1 && months[1].x - months[0].x < 3 * pitch) months.shift()
  return (
    <div className="dash-calendar" ref={box} onPointerLeave={leave}>
      {/* Day names stay in view while the weeks scroll. */}
      <div className="dash-calendar-frame">
        <svg className="dash-calendar-days" width={left} height={height} aria-hidden="true">
          {["Mon", "Wed", "Fri"].map((name, index) => (
            <text key={name} x={0} y={top + index * 2 * pitch + cell / 2} dominantBaseline="central" className="dash-calendar-axis">{name}</text>
          ))}
        </svg>
        <div className="dash-calendar-scroll" ref={scroller}>
          <svg width={weeks * pitch} height={height} role="group" aria-label={label ?? `Conversations started each day; busiest day had ${max}.`}>
            <g aria-hidden="true">
              {months.map((month) => (
                <text key={month.x} x={month.x} y={11} className="dash-calendar-axis">{month.label}</text>
              ))}
            </g>
            {days.map((day, index) => (
              <g key={index} transform={`translate(${Math.floor(index / 7) * pitch} ${top + (index % 7) * pitch})`}>
                <rect width={cell} height={cell} rx={cell * 0.3} className={`dash-cell-fill level-${day.level}`} aria-hidden="true" pointerEvents="none" />
                <rect
                  width={pitch}
                  height={pitch}
                  rx={cell * 0.3}
                  fill="transparent"
                  className="dash-cell"
                  role="img"
                  aria-label={day.text}
                  {...mark(index, (element) => show(element.getBoundingClientRect(), day.text))}
                  onBlur={hide}
                  onPointerEnter={(event) => show(event.currentTarget.getBoundingClientRect(), day.text)}
                  onPointerDown={(event) => show(event.currentTarget.getBoundingClientRect(), day.text)}
                />
              </g>
            ))}
          </svg>
        </div>
      </div>
      <div className="dash-calendar-footer">
        <p className="dash-calendar-hint">{weeks * pitch > room ? "Scroll horizontally to see earlier weeks · " : ""}Select a square to inspect its day.</p>
        <div className="dash-calendar-key" aria-hidden="true">
          <span>Less</span>
          {[0, 1, 2, 3, 4].map((step) => <span key={step} className={`dash-key-cell level-${step}`} />)}
          <span>More</span>
        </div>
      </div>
      {view}
    </div>
  )
}

// Progress as a ring: the done share as an arc from twelve o'clock (the
// -90° turn only moves the circle's start there), the rest as track. Each
// round cap reaches half a stroke past its dash, so the dash is one stroke
// shorter and starts half a stroke in: the painted arc is exactly the done
// share. Nothing is drawn when none is done.
export function Ring({ done, total, size = 112 }: { done: number; total: number; size?: number }) {
  const stroke = 12
  const radius = (size - stroke) / 2
  const length = 2 * Math.PI * radius
  const share = total > 0 ? Math.min(1, Math.max(0, done / total)) : 0
  return (
    <svg className="dash-ring" width={size} height={size} viewBox={`0 0 ${size} ${size}`} role="img" aria-label={`${done} of ${total} done`}>
      <circle cx={size / 2} cy={size / 2} r={radius} className="dash-ring-track" stroke="currentColor" strokeOpacity=".14" strokeWidth={stroke} fill="none" />
      {share > 0 && (
        <circle
          cx={size / 2}
          cy={size / 2}
          r={radius}
          className="dash-ring-value"
          stroke="currentColor"
          strokeWidth={stroke}
          fill="none"
          strokeLinecap="round"
          strokeDasharray={`${Math.max(0, share * length - stroke)} ${length}`}
          strokeDashoffset={-stroke / 2}
          transform={`rotate(-90 ${size / 2} ${size / 2})`}
        />
      )}
      <text x="50%" y="50%" dominantBaseline="central" textAnchor="middle" className="dash-ring-text">
        {done}<tspan className="dash-ring-total">/{total}</tspan>
      </text>
    </svg>
  )
}

export type Segment = { key: string; label: string; value: number; display: string }

// Parts of a whole on one bar, 2px apart, each named in the legend below.
// Segments are categorical, so each takes its own series class.
export function StackedBar({ segments, label }: { segments: Segment[]; label: string }) {
  const { box, show, hide, leave, view } = useTip()
  const positive = segments.filter((row) => row.value > 0)
  const mark = useRoving(positive.length, 0)
  const total = segments.reduce((sum, row) => sum + row.value, 0)
  const percent = (value: number) => (total ? `${Math.round((value / total) * 100)}%` : "0%")
  const color = (row: Segment) => {
    const semantic = ["input", "cached", "output", "reasoning"].indexOf(row.key)
    return `var(--dash-series-${(semantic < 0 ? segments.indexOf(row) % 4 : semantic) + 1})`
  }
  return (
    <div className="dash-stack" ref={box} onPointerLeave={leave}>
      <div className="dash-stack-bar" role="group" aria-label={`${label}: ${segments.map((row) => `${row.label} ${percent(row.value)}`).join(", ")}`}>
        {total > 0
          ? positive.map((row, index) => (
              <span
                key={row.key}
                className={`dash-stack-part series-${row.key}`}
                style={{ flexGrow: row.value, background: color(row) }}
                role="img"
                aria-label={`${row.label}: ${row.display} (${percent(row.value)})`}
                {...mark(index, (element) => show(element.getBoundingClientRect(), `${row.label}: ${row.display} (${percent(row.value)})`))}
                onBlur={hide}
                onPointerEnter={(event) => show(event, `${row.label}: ${row.display} (${percent(row.value)})`)}
                onPointerDown={(event) => show(event, `${row.label}: ${row.display} (${percent(row.value)})`)}
              />
            ))
          : <span className="dash-stack-part is-empty" />}
      </div>
      <ul className="dash-legend">
        {segments.map((row) => (
          <li key={row.key}>
            <span className={`dash-swatch series-${row.key}`} style={{ background: color(row) }} aria-hidden="true" />
            <span className="dash-legend-label">{row.label}</span>
            <span className="dash-legend-value">{row.display}</span>
            <span className="dash-legend-share">{percent(row.value)}</span>
          </li>
        ))}
      </ul>
      {view}
    </div>
  )
}

// Daily values over the window as a line over a faint flat area, with a
// crosshair that follows the pointer and names the day under it. Each day also
// has a focusable column, so the arrow keys move the same crosshair. `format`
// writes the axis; `detail`, the exact value in the tooltip.
export function AreaChart({ points, format, detail = format, label, integer = false }: { points: { time: number; value: number; label: string; detail?: string }[]; format: (value: number) => string; detail?: (value: number) => string; label: string; integer?: boolean }) {
  const [ref, { width, height }] = useSize<HTMLDivElement>()
  const [active, setActive] = useState<number>()
  useDismiss(ref, active !== undefined, () => setActive(undefined))
  const mark = useRoving(points.length, points.length - 1)
  const top = 12, bottom = 26, left = 52, right = 8
  const max = Math.max(...points.map((row) => row.value), 0)
  const scale = integer ? Math.ceil(niceMax(max)) : niceMax(max)
  const plot = Math.max(0, width - left - right)
  const step = points.length > 1 ? plot / (points.length - 1) : plot
  const x = (index: number) => left + (points.length > 1 ? index * step : plot / 2)
  const y = (value: number) => top + (1 - (scale ? value / scale : 0)) * (height - top - bottom)
  const line = points.map((row, index) => `${index ? "L" : "M"}${x(index).toFixed(1)} ${y(row.value).toFixed(1)}`).join(" ")
  const ticks = integer ? [...new Set([0, Math.round(scale / 2), scale])] : [0, scale / 2, scale]
  const labelEvery = Math.max(1, Math.ceil(points.length / Math.max(2, Math.floor(plot / 72))))
  function track(event: ReactPointerEvent<SVGSVGElement>) {
    const bounds = event.currentTarget.getBoundingClientRect()
    const ratio = (event.clientX - bounds.left - left) / Math.max(1, plot)
    setActive(Math.max(0, Math.min(points.length - 1, Math.round(ratio * (points.length - 1)))))
  }
  const current = active === undefined ? undefined : points[active]
  return (
    <div className="dash-area" ref={ref}>
      {width > 0 && height > 0 && (
        <svg width={width} height={height} role="group" aria-label={label} onPointerMove={track} onPointerDown={track} onPointerLeave={(event) => { if (event.pointerType !== "touch") setActive(undefined) }}>
          <g aria-hidden="true">
            {ticks.map((tick) => (
              <g key={tick}>
                <line x1={left} x2={width - right} y1={y(tick)} y2={y(tick)} className="dash-gridline" />
                <text x={left - 8} y={y(tick)} dominantBaseline="central" textAnchor="end" className="dash-axis">{format(tick)}</text>
              </g>
            ))}
            {points.map((row, index) => index % labelEvery === 0 && (
              <text key={`${row.time}-${index}`} x={x(index)} y={height - 8} textAnchor="middle" className="dash-axis">{row.label}</text>
            ))}
            {points.length > 1 && <path d={`${line} L${x(points.length - 1)} ${y(0)} L${x(0)} ${y(0)} Z`} className="dash-area-fill" fill="currentColor" fillOpacity=".1" />}
            <path d={line} fill="none" className="dash-area-line" stroke="currentColor" strokeWidth="2" strokeLinejoin="round" strokeLinecap="round" />
            {current && active !== undefined && (
              <g>
                <line x1={x(active)} x2={x(active)} y1={top} y2={y(0)} className="dash-crosshair" />
                <circle cx={x(active)} cy={y(current.value)} r="5" className="dash-area-dot" fill="currentColor" />
              </g>
            )}
          </g>
          {points.map((row, index) => {
            const from = Math.max(left, x(index) - step / 2), to = Math.min(left + plot, x(index) + step / 2)
            return (
              <rect
                key={`${row.time}-${index}`}
                className="dash-area-hit"
                x={from}
                y={top}
                width={Math.max(0, to - from)}
                height={Math.max(0, y(0) - top)}
                fill="transparent"
                role="img"
                aria-label={`${row.label}: ${detail(row.value)}${row.detail ? `. ${row.detail}` : ""}`}
                {...mark(index, () => setActive(index))}
                onBlur={() => setActive(undefined)}
              />
            )
          })}
        </svg>
      )}
      {current && active !== undefined && (
        <TipBox x={x(active)} y={y(current.value)}>
          <strong>{detail(current.value)}</strong> {current.label}
          {current.detail && <span className="dash-tip-detail">{current.detail}</span>}
        </TipBox>
      )}
    </div>
  )
}

// A round axis maximum (1, 2, 2.5 or 5 times a power of ten) above the data.
function niceMax(value: number) {
  if (value <= 0) return 1
  const power = 10 ** Math.floor(Math.log10(value))
  return ([1, 2, 2.5, 5, 10].find((step) => step * power >= value) ?? 10) * power
}
