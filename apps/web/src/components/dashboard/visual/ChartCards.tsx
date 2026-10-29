import type { CSSProperties } from "react"
import { ActivityCalendar, AreaChart } from "../charts"

export type ChartValueFormat = "money" | "count" | "percent"

// JSON-safe inputs: changing these props redraws the view without remounting or
// fetching. Series and calendar values are nonnegative. Percent values are in
// percentage points (25 means 25%, rather than a fraction of one).
export type ChartCardPoint = { key?: string; label: string; value: number; time?: number; detail?: string }
export type ChartCardData = {
  form: "stat" | "line" | "bars" | "list" | "calendar"
  label: string
  points: ChartCardPoint[]
  format?: ChartValueFormat
  currency?: string
  // A ranked view may show only some rows. Its supplied total is the full
  // denominator, so shares remain honest when a datasource returns its top N.
  total?: number
  basis?: string
  period?: string
  badge?: string
  showShares?: boolean
  notes?: string[]
  empty?: string
  unavailable?: boolean
  integer?: boolean
}

const count = new Intl.NumberFormat("en-GB", { maximumFractionDigits: 0 })
const percent = new Intl.NumberFormat("en-GB", { maximumFractionDigits: 1 })
const share = (value: number, total: number) => value > 0 && value / total < 0.005 ? "<1%" : `${Math.round(total ? value / total * 100 : 0)}%`

export function formatChartValue(value: number, format: ChartValueFormat = "count", currency = "USD") {
  if (!Number.isFinite(value)) return "—"
  if (format === "percent") return `${percent.format(value)}%`
  if (format === "count") return count.format(value)
  const formatter = new Intl.NumberFormat("en-US", { style: "currency", currency: /^[A-Z]{3}$/.test(currency) ? currency : "USD" })
  return value > 0 && value < 0.01 ? `< ${formatter.format(0.01)}` : formatter.format(value)
}

function axisValue(value: number, format: ChartValueFormat, currency?: string) {
  if (format !== "money") return formatChartValue(value, format)
  return new Intl.NumberFormat("en-US", { style: "currency", currency: currency && /^[A-Z]{3}$/.test(currency) ? currency : "USD", maximumFractionDigits: value >= 10 || value === 0 ? 0 : value >= 1 ? 1 : 2, minimumFractionDigits: 0 }).format(value)
}

export function StatValue({ value, format = "count", currency }: { value: number | null; format?: ChartValueFormat; currency?: string }) {
  return <strong className="dash-chart-total">{value === null ? "—" : formatChartValue(value, format, currency)}</strong>
}

// The rank/order is supplied by the caller. Keeping it intact also lets this
// component display deliberate, non-ranked categorical orders.
export function RankedBars({ label, points, total, format = "count", currency, bars = true, showShares = false }: Pick<ChartCardData, "label" | "points" | "total" | "format" | "currency" | "showShares"> & { bars?: boolean }) {
  const maximum = Math.max(0, ...points.map((point) => point.value))
  const denominator = total ?? points.reduce((sum, point) => sum + point.value, 0)
  return <ol className={`dash-chart-rows${bars ? " with-bars" : ""}`} aria-label={label}>
    {points.map((point, index) => <li key={point.key ?? `${index}-${point.label}`}>
      <div className="dash-chart-row-values"><span>{point.label}{point.detail && <span className="dash-chart-row-detail">{point.detail}</span>}</span><strong>{formatChartValue(point.value, format, currency)}{showShares && denominator > 0 && <span className="dash-chart-row-share"> · {share(point.value, denominator)}</span>}</strong></div>
      {bars && <span className="dash-chart-track" aria-hidden="true"><span style={{ "--chart-share": `${maximum ? point.value / maximum * 100 : 0}%` } as CSSProperties} /></span>}
    </li>)}
  </ol>
}

export function RealtimeSeriesCard({ label, points, format = "count", currency, integer }: Pick<ChartCardData, "label" | "points" | "format" | "currency" | "integer">) {
  return <AreaChart points={points.map((point, index) => ({ time: point.time ?? index, value: point.value, label: point.label, detail: point.detail }))} format={(value) => axisValue(value, format, currency)} detail={(value) => formatChartValue(value, format, currency)} label={label} integer={integer ?? format === "count"} />
}

const dayStart = (time: number) => { const date = new Date(time); date.setHours(0, 0, 0, 0); return date.getTime() }
const plusDays = (time: number, count: number) => { const date = new Date(time); date.setDate(date.getDate() + count); return date.getTime() }

// Missing dates stay visibly empty and are described as unavailable, rather
// than inventing measured zero activity. Dated input is required for calendars.
export function ActivityChart({ label, points, format = "count", currency }: Pick<ChartCardData, "label" | "points" | "format" | "currency">) {
  const dated = points.filter((point) => point.time !== undefined && Number.isFinite(point.time) && Number.isFinite(new Date(point.time).getTime()))
  if (!dated.length) return <p className="dash-chart-empty">Add a date to each activity value to display a calendar.</p>
  const first = dayStart(Math.min(...dated.map((point) => point.time!)))
  const last = dayStart(Math.max(...dated.map((point) => point.time!)))
  const monday = plusDays(first, -(new Date(first).getDay() + 6) % 7)
  const byDate = new Map(dated.map((point) => [dayStart(point.time!), point]))
  if (byDate.size !== dated.length) return <p className="dash-chart-empty">Supply one activity value per calendar day.</p>
  const counts: number[] = []
  for (let day = monday; day <= last && counts.length < 3660; day = plusDays(day, 1)) counts.push(byDate.get(day)?.value ?? 0)
  if (plusDays(monday, counts.length) <= last) return <p className="dash-chart-empty">Choose a calendar period of ten years or less.</p>
  return <ActivityCalendar label={label} first={monday} counts={counts} describe={(day, value) => {
    const date = new Date(day).toLocaleDateString("en-GB", { day: "numeric", month: "short", year: "numeric" })
    const point = byDate.get(day)
    return `${date}: ${day < first ? "outside this period" : point ? `${formatChartValue(value, format, currency)}${point.detail ? ` · ${point.detail}` : ""}` : "no data supplied"}`
  }} />
}

// A chart's data view only. Dashboard/page composition owns its surrounding
// heading, card size and controls. Datasource subscriptions belong to callers.
export function ChartCard({ form, label, points: input, format = "count", currency, total, basis, period, badge, showShares, notes = [], empty = "Nothing recorded for this view yet.", unavailable = false, integer }: ChartCardData) {
  const points = input.filter((point) => Number.isFinite(point.value) && (form === "stat" || point.value >= 0))
  const sum = total !== undefined && Number.isFinite(total) ? total : points.reduce((value, point) => value + point.value, 0)
  const noData = unavailable || points.length === 0
  const props = { label, points, format, currency }
  return <div className={`dash-chart-widget form-${form}`}>
    <div className="dash-chart-summary">
      <StatValue value={unavailable ? null : sum} format={format} currency={currency} />
      {basis && <span className="dash-chart-basis">{basis}</span>}
      {badge && !noData && <span className="dash-chart-badge">{badge}</span>}
    </div>
    {period && <p className="dash-chart-period">{period}</p>}
    {noData ? <p className="dash-chart-empty">{empty}</p> : form === "stat" ? null : form === "line" ? <RealtimeSeriesCard {...props} integer={integer} /> : form === "calendar" ? <ActivityChart {...props} /> : <RankedBars {...props} total={sum} bars={form === "bars"} showShares={showShares} />}
    {notes.length > 0 && <ul className="dash-chart-notes">{notes.map((note, index) => <li key={`${index}-${note}`}>{note}</li>)}</ul>}
  </div>
}
