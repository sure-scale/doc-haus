import type { Chart } from "../../../../../dochaus/lib/dashboard-plan"
import { chartData } from "./chart-data"
import type { DashboardData, Period } from "./data"
import { ChartCard, formatChartValue } from "./visual/ChartCards"

// This adapter supplies practice data to the same presentational components a
// custom datasource can use. It computes no figures inside the visual layer.
export function ChartWidget({ chart, data, period }: { chart: Chart; data: DashboardData; period: Period }) {
  const view = chartData(chart, data, period)
  const format = chart.metric === "spend" ? "money" : "count"
  const maximum = Math.max(0, ...view.points.map((point) => point.value))
  const ranked = chart.groupBy === "client" || chart.groupBy === "matter"
  const interval = chart.groupBy === "day" ? "day" : chart.groupBy === "week" ? "wk" : "mo"
  const badge = ranked ? `${view.groups} ${chart.groupBy}${view.groups === 1 ? "" : "s"}` : chart.form === "line" ? `Peak: ${formatChartValue(maximum, format)} / ${interval}` : undefined
  const metric = chart.metric === "spend" ? "Estimated spend · USD" : chart.metric === "conversations" ? "Conversations started" : chart.metric === "matters" ? "Matters opened" : "Drafted documents"
  return <ChartCard
    form={chart.form}
    label={chart.title}
    points={view.total === 0 ? [] : view.points}
    total={view.total}
    format={format}
    basis={`${metric}${view.incomplete ? " · partial" : ""}`}
    period={`${view.period}${chart.client ? ` · ${chart.client}` : chart.matter ? ` · ${chart.matter}` : ""}`}
    badge={badge}
    showShares={ranked}
    unavailable={view.unavailable}
    notes={view.notes}
    empty={view.empty}
  />
}
