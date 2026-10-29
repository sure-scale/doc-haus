import { useLayoutEffect, useMemo, useState } from "react"
import type { Chart } from "../../../../../dochaus/lib/dashboard-plan"
import { chartData } from "./chart-data"
import type { DashboardData, Period } from "./data"
import { formatChartValue } from "./visual/ChartCards"
import { LiveComponent, type ComponentDefinition } from "./visual/LiveComponents"
import { createDashboardSource, type JsonValue } from "./visual/source"

// Only the computed view is published, never the underlying matter records or
// preferences. Both the source and its registration survive data refreshes.
export function chartProjection(chart: Chart, data: DashboardData, period: Period): { [key: string]: JsonValue } {
  const view = chartData(chart, data, period)
  const format = chart.metric === "spend" ? "money" : "count"
  const maximum = Math.max(0, ...view.points.map((point) => point.value))
  const ranked = chart.groupBy === "client" || chart.groupBy === "matter"
  const interval = chart.groupBy === "day" ? "day" : chart.groupBy === "week" ? "wk" : "mo"
  const badge = ranked ? `${view.groups} ${chart.groupBy}${view.groups === 1 ? "" : "s"}` : chart.form === "line" ? `Peak: ${formatChartValue(maximum, format)} / ${interval}` : ""
  const metric = chart.metric === "spend" ? "Estimated spend · USD" : chart.metric === "conversations" ? "Conversations started" : chart.metric === "matters" ? "Matters opened" : "Drafted documents"
  return {
    // Group keys can be empty or longer than component IDs; the renderer can
    // safely derive row identity from the ordered label and index instead.
    points: view.total === 0 ? [] : view.points.map(({ key: _key, ...point }) => point),
    total: view.total,
    basis: `${metric}${view.incomplete ? " · partial" : ""}`,
    period: `${view.period}${chart.client ? ` · ${chart.client}` : chart.matter ? ` · ${chart.matter}` : ""}`,
    badge,
    unavailable: view.unavailable,
    notes: view.notes,
    empty: view.empty,
  }
}

export function ChartWidget({ chart, data, period }: { chart: Chart; data: DashboardData; period: Period }) {
  const projection = useMemo(() => chartProjection(chart, data, period), [chart, data, period])
  const [source] = useState(() => createDashboardSource(projection))
  const [sources] = useState(() => ({ chart: source }))
  useLayoutEffect(() => { source.publish(projection) }, [source, projection])
  const definition: ComponentDefinition = {
    id: "chart", kind: "chart",
    props: { form: chart.form, label: chart.title, format: chart.metric === "spend" ? "money" : "count", showShares: chart.groupBy === "client" || chart.groupBy === "matter" },
    bindings: Object.fromEntries(["points", "total", "basis", "period", "badge", "unavailable", "notes", "empty"].map((field) => [field, { source: "chart", path: `/${field}` }])),
  }
  return <LiveComponent definition={definition} sources={sources} />
}
