import { useId, useLayoutEffect, useRef, useState, type CSSProperties, type ReactNode, type PointerEvent as ReactPointerEvent } from "react"
import { addDays } from "./data"

// The dashboard's chart marks, drawn as inline SVG against the widget's own
// colour tokens (see styles/dashboard.css). Every chart has an aria-label that
// states what it shows; hover tooltips add exact values.

type Tip = { x: number; y: number; content: ReactNode }

// One hover tooltip per chart, positioned inside the chart's own box.
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
  const view = tip && (
    <div className="dash-tip" role="tooltip" style={{ left: tip.x, top: tip.y }}>{tip.content}</div>
  )
  return { box, show, hide: () => setTip(undefined), view }
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

// A stat tile's trend: a 2px line over a soft area, its last point marked.
export function Sparkline({ values, label }: { values: number[]; label: string }) {
  const gradient = useId()
  const max = Math.max(...values, 0)
  if (values.length < 2 || max <= 0) return <div className="dash-spark is-flat" role="img" aria-label={`${label}: no activity yet`} />
  const width = 120, height = 40, step = width / (values.length - 1)
  const points = values.map((value, index) => [index * step, height - 3 - (value / max) * (height - 8)] as const)
  const line = points.map(([x, y], index) => `${index ? "L" : "M"}${x.toFixed(1)} ${y.toFixed(1)}`).join(" ")
  const [lastX, lastY] = points[points.length - 1]
  return (
    <svg className="dash-spark" viewBox={`0 0 ${width} ${height}`} preserveAspectRatio="none" role="img" aria-label={label}>
      <defs>
        <linearGradient id={gradient} x1="0" x2="0" y1="0" y2="1">
          <stop offset="0" stopColor="currentColor" stopOpacity=".32" />
          <stop offset="1" stopColor="currentColor" stopOpacity="0" />
        </linearGradient>
      </defs>
      <path d={`${line} L${width} ${height} L0 ${height} Z`} fill={`url(#${gradient})`} />
      <path d={line} fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" vectorEffect="non-scaling-stroke" />
      <circle cx={lastX} cy={lastY} r="3.5" fill="currentColor" className="dash-spark-dot" vectorEffect="non-scaling-stroke" />
    </svg>
  )
}

// Spread behind an average: one thin bar per item, the average as a rule.
export function SpreadBars({ values, average, label, format }: { values: number[]; average: number; label: string; format: (value: number) => string }) {
  const { box, show, hide, view } = useTip()
  const max = Math.max(...values, average, 0)
  if (!values.length || max <= 0) return <div className="dash-spark is-flat" role="img" aria-label={`${label}: nothing yet`} />
  const shown = values.slice(0, 16)
  return (
    <div className="dash-spread" ref={box} role="img" aria-label={label} onPointerLeave={hide}>
      {shown.map((value, index) => (
        <span
          key={index}
          className="dash-spread-bar"
          style={{ height: `${Math.max(6, (value / max) * 100)}%` }}
          onPointerEnter={(event) => show(event.currentTarget.getBoundingClientRect(), format(value))}
        />
      ))}
      <span className="dash-spread-average" style={{ bottom: `${(average / max) * 100}%` }} aria-hidden="true" />
      {view}
    </div>
  )
}

// The week view: a pill per day on a full-height track, a dot capping each
// fill. Each fill shows the slice of one track-high gradient up to its top, so
// a bar's colour follows its height; today's label is emphasised.
export function WeekBars({ days, format }: { days: { day: number; cost: number; label: string; today: boolean }[]; format: (value: number) => string }) {
  const { box, show, hide, view } = useTip()
  const max = Math.max(...days.map((row) => row.cost), 0)
  return (
    <div className="dash-week" ref={box} onPointerLeave={hide}>
      <ol className="dash-week-bars" aria-label="Spend on each of the last seven days">
        {days.map((row) => {
          const share = max > 0 ? row.cost / max : 0
          const text = `${row.today ? "Today" : row.label}: ${format(row.cost)}`
          return (
            <li key={row.day} className={`dash-week-day${row.today ? " is-today" : ""}`}>
              <span
                className="dash-week-track"
                tabIndex={0}
                aria-label={text}
                onPointerEnter={(event) => show(event.currentTarget.getBoundingClientRect(), text)}
                onFocus={(event) => show(event.currentTarget.getBoundingClientRect(), text)}
                onBlur={hide}
              >
                {share > 0 && (
                  <span className="dash-week-fill" style={{ height: `max(22px, ${share * 100}%)`, "--share": share } as CSSProperties}>
                    <span className="dash-week-dot" />
                  </span>
                )}
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

// A year of days as a calendar grid: one column per week (Monday first), one
// cell per day, darker for busier days. Cells grow to fill a wide card and
// keep a readable size on a narrow one, which scrolls instead.
export function ActivityCalendar({ counts, first, describe }: { counts: number[]; first: number; describe: (day: number, count: number) => string }) {
  const { box, show, hide, view } = useTip()
  const [scroller, { width: room }] = useSize<HTMLDivElement>()
  const top = 18, left = 30
  const weeks = Math.ceil(counts.length / 7)
  const pitch = Math.min(24, Math.max(13, room / weeks))
  const cell = pitch * 0.8
  // Start scrolled to today, as the newest weeks matter most on a narrow screen.
  useLayoutEffect(() => {
    const element = scroller.current
    if (element) element.scrollLeft = element.scrollWidth
  }, [counts.length, pitch])
  const max = Math.max(...counts, 0)
  const level = (count: number) => (count <= 0 || max <= 0 ? 0 : Math.min(4, Math.ceil((count / max) * 4)))
  const height = top + 7 * pitch
  const months: { x: number; label: string }[] = []
  for (let week = 0; week < weeks; week++) {
    const date = new Date(addDays(first, week * 7))
    const label = date.toLocaleDateString("en-GB", { month: "short" })
    if (!months.length || months[months.length - 1].label !== label) months.push({ x: week * pitch, label })
  }
  if (months.length > 1 && months[1].x - months[0].x < 3 * pitch) months.shift()
  return (
    <div className="dash-calendar" ref={box} onPointerLeave={hide}>
      {/* Day names stay in view while the weeks scroll. */}
      <div className="dash-calendar-frame">
        <svg className="dash-calendar-days" width={left} height={height} aria-hidden="true">
          {["Mon", "Wed", "Fri"].map((name, index) => (
            <text key={name} x={0} y={top + index * 2 * pitch + cell / 2} dominantBaseline="central" className="dash-calendar-axis">{name}</text>
          ))}
        </svg>
        <div className="dash-calendar-scroll" ref={scroller}>
          <svg width={weeks * pitch} height={height} role="img" aria-label={`Conversations started each day; busiest day had ${max}.`}>
            {months.map((month) => (
              <text key={month.x} x={month.x} y={11} className="dash-calendar-axis">{month.label}</text>
            ))}
            {counts.map((count, index) => {
              const day = addDays(first, index)
              return (
                <rect
                  key={index}
                  x={Math.floor(index / 7) * pitch}
                  y={top + (index % 7) * pitch}
                  width={cell}
                  height={cell}
                  rx={cell * 0.3}
                  className={`dash-cell level-${level(count)}`}
                  onPointerEnter={(event) => show(event.currentTarget.getBoundingClientRect(), describe(day, count))}
                />
              )
            })}
          </svg>
        </div>
      </div>
      <div className="dash-calendar-key" aria-hidden="true">
        <span>Less</span>
        {[0, 1, 2, 3, 4].map((step) => <span key={step} className={`dash-key-cell level-${step}`} />)}
        <span>More</span>
      </div>
      {view}
    </div>
  )
}

// Progress as a ring: the done share in the strong tone, the rest as track.
export function Ring({ done, total, size = 112 }: { done: number; total: number; size?: number }) {
  const stroke = 12
  const radius = (size - stroke) / 2
  const length = 2 * Math.PI * radius
  const share = total ? done / total : 0
  return (
    <svg className="dash-ring" width={size} height={size} viewBox={`0 0 ${size} ${size}`} role="img" aria-label={`${done} of ${total} done`}>
      <circle cx={size / 2} cy={size / 2} r={radius} className="dash-ring-track" strokeWidth={stroke} fill="none" />
      <circle
        cx={size / 2}
        cy={size / 2}
        r={radius}
        className="dash-ring-value"
        strokeWidth={stroke}
        fill="none"
        strokeLinecap="round"
        strokeDasharray={`${Math.max(share > 0 ? 0.001 : 0, share) * length} ${length}`}
        transform={`rotate(-90 ${size / 2} ${size / 2})`}
      />
      <text x="50%" y="50%" dominantBaseline="central" textAnchor="middle" className="dash-ring-text">
        {done}<tspan className="dash-ring-total">/{total}</tspan>
      </text>
    </svg>
  )
}

export type Segment = { key: string; label: string; value: number; display: string }

// Parts of a whole on one bar, 2px apart, each named in the legend below.
export function StackedBar({ segments, label }: { segments: Segment[]; label: string }) {
  const { box, show, hide, view } = useTip()
  const total = segments.reduce((sum, row) => sum + row.value, 0)
  const percent = (value: number) => (total ? `${Math.round((value / total) * 100)}%` : "0%")
  return (
    <div className="dash-stack" ref={box} onPointerLeave={hide}>
      <div className="dash-stack-bar" role="img" aria-label={`${label}: ${segments.map((row) => `${row.label} ${percent(row.value)}`).join(", ")}`}>
        {total > 0
          ? segments.filter((row) => row.value > 0).map((row) => (
              <span
                key={row.key}
                className={`dash-stack-part series-${row.key}`}
                style={{ flexGrow: row.value }}
                onPointerEnter={(event) => show(event, `${row.label}: ${row.display} (${percent(row.value)})`)}
              />
            ))
          : <span className="dash-stack-part is-empty" />}
      </div>
      <ul className="dash-legend">
        {segments.map((row) => (
          <li key={row.key}>
            <span className={`dash-swatch series-${row.key}`} aria-hidden="true" />
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

// Daily values over the window as an area with a crosshair that follows the
// pointer and names the day under it. `format` writes the axis; `detail`, the
// exact value in the tooltip.
export function AreaChart({ points, format, detail = format, label }: { points: { time: number; value: number; label: string }[]; format: (value: number) => string; detail?: (value: number) => string; label: string }) {
  const gradient = useId()
  const [ref, { width, height }] = useSize<HTMLDivElement>()
  const [active, setActive] = useState<number>()
  const top = 12, bottom = 26, left = 52, right = 8
  const max = Math.max(...points.map((row) => row.value), 0)
  const scale = niceMax(max)
  const plot = Math.max(0, width - left - right)
  const x = (index: number) => left + (points.length > 1 ? (index / (points.length - 1)) * plot : plot / 2)
  const y = (value: number) => top + (1 - (scale ? value / scale : 0)) * (height - top - bottom)
  const line = points.map((row, index) => `${index ? "L" : "M"}${x(index).toFixed(1)} ${y(row.value).toFixed(1)}`).join(" ")
  const ticks = [0, scale / 2, scale]
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
        <svg width={width} height={height} role="img" aria-label={label} onPointerMove={track} onPointerLeave={() => setActive(undefined)}>
          <defs>
            <linearGradient id={gradient} x1="0" x2="0" y1="0" y2="1">
              <stop offset="0" stopColor="var(--dash-series-1)" stopOpacity=".28" />
              <stop offset="1" stopColor="var(--dash-series-1)" stopOpacity="0" />
            </linearGradient>
          </defs>
          {ticks.map((tick) => (
            <g key={tick}>
              <line x1={left} x2={width - right} y1={y(tick)} y2={y(tick)} className="dash-gridline" />
              <text x={left - 8} y={y(tick)} dominantBaseline="central" textAnchor="end" className="dash-axis">{format(tick)}</text>
            </g>
          ))}
          {points.map((row, index) => index % labelEvery === 0 && (
            <text key={row.time} x={x(index)} y={height - 8} textAnchor="middle" className="dash-axis">{row.label}</text>
          ))}
          {points.length > 1 && <path d={`${line} L${x(points.length - 1)} ${y(0)} L${x(0)} ${y(0)} Z`} fill={`url(#${gradient})`} />}
          <path d={line} fill="none" className="dash-area-line" strokeWidth="2" strokeLinejoin="round" strokeLinecap="round" />
          {current && active !== undefined && (
            <g>
              <line x1={x(active)} x2={x(active)} y1={top} y2={y(0)} className="dash-crosshair" />
              <circle cx={x(active)} cy={y(current.value)} r="5" className="dash-area-dot" />
            </g>
          )}
        </svg>
      )}
      {current && active !== undefined && (
        <div className="dash-tip" role="tooltip" style={{ left: x(active), top: y(current.value) }}>
          <strong>{detail(current.value)}</strong> {current.label}
        </div>
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
