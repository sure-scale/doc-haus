import type { CSSProperties } from "react"
import type { Chart } from "../../../../../dochaus/lib/dashboard-plan"
import { ActivityCalendar, AreaChart } from "./charts"
import { chartData } from "./chart-data"
import { addDays, dayIndex, startOfDay, type DashboardData, type Period } from "./data"

const money = (value: number) => value > 0 && value < 0.01 ? "< $0.01" : new Intl.NumberFormat("en-US", { style: "currency", currency: "USD" }).format(value)
const count = (value: number) => new Intl.NumberFormat("en-GB", { maximumFractionDigits: 0 }).format(value)
const axisMoney = (value: number) => value >= 10 ? `$${Math.round(value)}` : value ? `$${value.toFixed(value >= 1 ? 1 : 2)}` : "$0"

// Cards, headings and pin/remove controls belong to the page. This renders
// only the data view, so ready-made and AI-selected widgets share the shell.
export function ChartWidget({ chart, data, period }: { chart: Chart; data: DashboardData; period: Period }) {
  const view = chartData(chart, data, period)
  const format = chart.metric === "spend" ? money : count
  const max = Math.max(0, ...view.points.map((row) => row.value))
  const empty = data.error || !view.points.length || view.total === 0
  const metric = chart.metric === "spend" ? "Estimated spend · USD" : chart.metric === "conversations" ? "Conversations started" : chart.metric === "matters" ? "Matters opened" : "Drafted documents"
  // Calendar columns begin on Monday. Leading cells fall outside the request
  // and are described as such rather than implying measured zero activity.
  const first = view.points[0]?.time ?? startOfDay(Date.now())
  const monday = addDays(first, -(new Date(first).getDay() + 6) % 7)
  const padding = dayIndex(monday, first)
  return (
    <div className={`dash-chart-widget form-${chart.form}`}>
      <div className="dash-chart-summary">
        <strong className="dash-chart-total">{view.unavailable ? "—" : format(view.total)}</strong>
        <span>{metric}{view.incomplete ? " · partial" : ""}</span>
      </div>
      <p className="dash-chart-period">{view.period}{chart.client ? ` · ${chart.client}` : chart.matter ? ` · ${chart.matter}` : ""}</p>
      {empty ? <p className="dash-chart-empty">{view.empty}</p>
        : chart.form === "stat" ? null
        : chart.form === "line" ? (
          <AreaChart points={view.points.map((row) => ({ time: row.time!, value: row.value, label: row.label }))} format={chart.metric === "spend" ? axisMoney : count} detail={format} label={chart.title} />
        ) : chart.form === "calendar" ? (
          <ActivityCalendar label={chart.title} first={monday} counts={[...new Array<number>(padding).fill(0), ...view.points.map((row) => row.value)]} describe={(day, value) => `${new Date(day).toLocaleDateString("en-GB", { day: "numeric", month: "short", year: "numeric" })}: ${day < first ? "outside this period" : format(value)}`} />
        ) : (
          <ol className={`dash-chart-rows${chart.form === "bars" ? " with-bars" : ""}`} aria-label={chart.title}>
            {view.points.map((row) => (
              <li key={row.key}>
                <div className="dash-chart-row-values"><span>{row.label}</span><strong>{format(row.value)}</strong></div>
                {chart.form === "bars" && <span className="dash-chart-track" aria-hidden="true"><span style={{ "--chart-share": `${max ? row.value / max * 100 : 0}%` } as CSSProperties} /></span>}
              </li>
            ))}
          </ol>
        )}
      {view.notes.length > 0 && <ul className="dash-chart-notes">{view.notes.map((note) => <li key={note}>{note}</li>)}</ul>}
    </div>
  )
}
