import type { Chart, Grouping, Metric } from "../../../../../dochaus/lib/dashboard-plan"
import { addDays, draftedCount, isConversation, PERIOD_LABELS, periodRange, startOfDay, type DashboardData, type MatterActivity, type Period } from "./data"

export type ChartPoint = { key: string; label: string; value: number; time?: number }
export type ChartData = {
  points: ChartPoint[]
  total: number
  period: string
  notes: string[]
  empty: string
  groups: number
  unavailable: boolean
  incomplete: boolean
}

type Event = { time?: number; value: number; row: MatterActivity }
const TIME = new Set<Grouping>(["day", "week", "month"])
const dateLabel = (time: number, month = false) => new Date(time).toLocaleDateString("en-GB", month ? { month: "short", year: "numeric" } : { day: "numeric", month: "short" })
const plural = (n: number, one: string) => `${n} ${one}${n === 1 ? "" : "s"}`

function bucket(time: number, group: Grouping) {
  const date = new Date(startOfDay(time))
  if (group === "month") date.setDate(1)
  if (group === "week") date.setDate(date.getDate() - (date.getDay() + 6) % 7)
  return date.getTime()
}

function nextBucket(time: number, group: Grouping) {
  if (group !== "month") return addDays(time, group === "week" ? 7 : 1)
  const date = new Date(time)
  date.setMonth(date.getMonth() + 1)
  return date.getTime()
}

function events(row: MatterActivity, metric: Metric): Event[] {
  if (metric === "spend") return row.costs.map(({ time, cost }) => ({ time, value: cost, row }))
  if (metric === "conversations") return row.sessions.filter(isConversation).map(({ time }) => ({ time: time.created, value: 1, row }))
  if (metric === "matters") return [{ time: row.matter.created_at, value: 1, row }]
  // The latest run's delivered body files are distinct drafts. Its updated_at
  // is a mutable status timestamp, so it cannot date the delivered documents.
  return [{ value: row.drafting?.status === "completed" ? draftedCount(row.drafting) : 0, row }]
}

// The AI chooses settings only. All values come from the same matter/session
// records as the rest of the dashboard, with the fictional sample excluded.
export function chartData(chart: Chart, data: DashboardData, dashboardPeriod: Period, now = Date.now()): ChartData {
  const period = chart.period === "dashboard" ? dashboardPeriod : chart.period
  const range = periodRange(period, now)
  const own = data.matters.filter(({ matter }) => !matter.sample)
  const rows = own.filter(({ matter }) =>
    (chart.client === undefined || matter.client?.trim() === chart.client) &&
    (chart.matter === undefined || matter.title === chart.matter))
  const notes: string[] = []
  const result = (points: ChartPoint[], total: number, groups = points.length): ChartData => ({
    points, total, groups, notes, unavailable: Boolean(data.error), incomplete: chart.metric !== "matters" && rows.some((row) => row.unread),
    period: chart.metric === "drafts" ? "Latest finished run in each matter · all time" : PERIOD_LABELS[period],
    empty: data.error ? "These figures could not be loaded." : chart.client !== undefined && !rows.length ? `No matters found for ${chart.client}.`
      : chart.matter !== undefined && !rows.length ? `No matter named ${chart.matter} was found.`
      : rows.some((row) => row.unread) && chart.metric !== "matters" ? "No figures are available from the activity that could be read." : "Nothing recorded for this view yet.",
  })
  if (data.error) { notes.push(data.error); return result([], 0) }
  if (chart.matter !== undefined && rows.length > 1) {
    notes.push("Several matters share this name. Ask for a view by client instead.")
    return { ...result([], 0), unavailable: true, empty: "This matter name is not unique." }
  }
  if (chart.metric === "drafts") {
    notes.push("Draft dates are not tracked. This counts delivered body files from the latest finished run in each matter, across all time.")
    if (TIME.has(chart.groupBy)) {
      notes.push("Drafts cannot be drawn over time. Ask for a total, or a view by matter or client.")
      return { ...result([], 0), unavailable: true, empty: "Drafts cannot be dated for this view." }
    }
  }
  if (rows.some((row) => row.unread) && chart.metric !== "matters") notes.push("Some matter activity could not be read. These figures may be incomplete.")
  if (chart.metric === "spend" && rows.some((row) => row.unpriced > 0)) notes.push("Some AI use has no price and is missing from this estimate.")
  const values = rows.flatMap((row) => events(row, chart.metric)).filter((entry) =>
    Number.isFinite(entry.value) && entry.value >= 0 &&
    (chart.metric === "drafts" || (entry.time !== undefined && Number.isFinite(entry.time) && entry.time >= range.start && entry.time < range.end)))
  const total = values.reduce((sum, entry) => sum + entry.value, 0)
  if (chart.groupBy === "none") return result([{ key: "total", label: chart.title, value: total }], total)
  if (TIME.has(chart.groupBy)) {
    const first = Number.isFinite(range.start) ? range.start : Math.min(startOfDay(now), ...values.map((entry) => entry.time!))
    const groups = new Map<number, ChartPoint>()
    for (let time = bucket(first, chart.groupBy); time < range.end; time = nextBucket(time, chart.groupBy)) {
      groups.set(time, { key: String(time), time, label: chart.groupBy === "month" ? dateLabel(time, true) : `${chart.groupBy === "week" ? "Week of " : ""}${dateLabel(time)}`, value: 0 })
    }
    for (const entry of values) {
      const point = groups.get(bucket(entry.time!, chart.groupBy))
      if (point) point.value += entry.value
    }
    return result([...groups.values()], total)
  }
  const groups = new Map<string, ChartPoint>()
  for (const entry of values) {
    const matter = entry.row.matter
    const key = chart.groupBy === "matter" ? matter.id : matter.client?.trim() || ""
    const label = chart.groupBy === "matter" ? matter.title : key || "No client recorded"
    const group = groups.get(key) ?? { key, label, value: 0 }
    group.value += entry.value
    groups.set(key, group)
  }
  const sorted = [...groups.values()].filter((point) => point.value > 0).sort((a, b) => b.value - a.value || a.label.localeCompare(b.label))
  const shown = sorted.slice(0, chart.limit ?? 12)
  if (shown.length < sorted.length) notes.push(`Showing the top ${shown.length} of ${plural(sorted.length, chart.groupBy === "client" ? "client" : "matter")}. The total includes all of them.`)
  return result(shown, total, sorted.length)
}
