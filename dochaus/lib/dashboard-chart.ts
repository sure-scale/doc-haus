import type { Chart, Grouping, Metric } from "./dashboard-plan"
export type Period = "month" | "last-month" | "quarter" | "all"
export type MatterActivity = {
  matter: { id: string; title: string; client?: string; sample?: boolean; created_at: number }
  sessions: { id: string; parentID?: string; title?: string; metadata?: { dochaus_role?: unknown }; tokens?: { input?: number; output?: number; reasoning?: number; cache?: { read?: number; write?: number } }; time: { created: number; updated: number } }[]
  drafting?: { status?: string; documents?: { status: string; documents: { kind: string; deliverable?: unknown }[] } | null }
  documents?: number
  costs: { time: number; cost: number }[]
  unread: boolean
  unpriced: number
}
export type DashboardData = { matters: MatterActivity[]; error?: string }
const PERIOD_LABELS: Record<Period, string> = { month: "This month", "last-month": "Last month", quarter: "Last 3 months", all: "All time" }
function isConversation(session: MatterActivity["sessions"][number]) { const role = session.metadata?.dochaus_role; return !session.parentID && !/^doc\.haus/i.test(session.title ?? "") && (role === undefined || role === "lean-drafting-owner") }
function draftedCount(drafting: MatterActivity["drafting"]) { return drafting?.documents?.status === "completed" ? drafting.documents.documents.filter(row => row.kind === "body" && row.deliverable).length : 0 }
function calendar(zone: string) {
  const fmt = new Intl.DateTimeFormat("en-GB", { timeZone: zone, year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit", second: "2-digit", hourCycle: "h23" })
  const parts = (time: number) => { const p = Object.fromEntries(fmt.formatToParts(time).map(x => [x.type, Number(x.value)])); return Date.UTC(p.year, p.month - 1, p.day, p.hour, p.minute, p.second) }
  // Most midnights have a unique offset and converge immediately. A clock can
  // also skip midnight, repeat it, or skip an entire civil date. In those cases
  // find the first instant whose local date reaches the requested civil date.
  const midnight = (wall: number) => {
    let instant = wall
    for (let i = 0; i < 4; i++) { const next = instant + wall - parts(instant); if (next === instant) break; instant = next }
    const localDay = (time: number) => { const date = new Date(parts(time)); return Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate()) }
    if (parts(instant) === wall && localDay(instant - 1000) < wall) return instant
    const hour = 3600000
    let before = wall - 48 * hour
    for (let after = before + hour; after <= wall + 48 * hour; before = after, after += hour) {
      if (localDay(after) < wall) continue
      // Intl has second precision; return the first second of the local date.
      let lo = Math.floor(before / 1000), hi = Math.ceil(after / 1000)
      while (hi - lo > 1) { const mid = Math.floor((lo + hi) / 2); if (localDay(mid * 1000) >= wall) hi = mid; else lo = mid }
      return hi * 1000
    }
    throw new Error("Could not resolve dashboard calendar date.")
  }
  const wallDay = (time: number) => { const d = new Date(parts(time)); return Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()) }
  const day = (time: number) => midnight(wallDay(time))
  const add = (time: number, days: number) => midnight(wallDay(time) + days * 86400000)
  const month = (time: number, shift = 0) => { const d = new Date(parts(time)); return midnight(Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + shift, 1)) }
  return { day, add, month, wallDay, midnight }
}

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
const dateLabel = (time: number, month = false, zone?: string) => new Date(time).toLocaleDateString("en-GB", { timeZone: zone, ...(month ? { month: "short", year: "numeric" } : { day: "numeric", month: "short" }) })
const plural = (n: number, one: string) => `${n} ${one}${n === 1 ? "" : "s"}`

function bucket(time: number, group: Grouping, c: ReturnType<typeof calendar>) {
  const wall = new Date(c.wallDay(time))
  if (group === "month") wall.setUTCDate(1)
  if (group === "week") wall.setUTCDate(wall.getUTCDate() - (wall.getUTCDay() + 6) % 7)
  return c.midnight(wall.getTime())
}

function nextBucket(time: number, group: Grouping, c: ReturnType<typeof calendar>) {
  return group === "month" ? c.month(time, 1) : c.add(time, group === "week" ? 7 : 1)
}

function events(row: MatterActivity, metric: Metric): Event[] {
  if (metric === "spend") return row.costs.map(({ time, cost }) => ({ time, value: cost, row }))
  if (metric === "conversations") return row.sessions.filter(isConversation).map(({ time }) => ({ time: time.created, value: 1, row }))
  if (metric === "matters") return [{ time: row.matter.created_at, value: 1, row }]
  if (metric === "documents") return [{ value: row.documents ?? 0, row }]
  if (metric === "tokens") return row.sessions.map(({ time, tokens }) => ({ time: time.created, value: (tokens?.input || 0) + (tokens?.output || 0) + (tokens?.reasoning || 0) + (tokens?.cache?.read || 0) + (tokens?.cache?.write || 0), row }))
  // The latest run's delivered body files are distinct drafts. Its updated_at
  // is a mutable status timestamp, so it cannot date the delivered documents.
  return [{ value: row.drafting?.status === "completed" ? draftedCount(row.drafting) : 0, row }]
}

// The AI chooses settings only. All values come from the same matter/session
// records as the rest of the dashboard, with the fictional sample excluded.
export function chartData(chart: Chart, data: DashboardData, dashboardPeriod: Period, now = Date.now(), zone = Intl.DateTimeFormat().resolvedOptions().timeZone): ChartData {
  const period = chart.period === "dashboard" ? dashboardPeriod : chart.period
  const c = calendar(zone)
  const end = c.add(c.day(now), 1)
  const range = period === "month" ? { start: c.month(now), end } : period === "last-month" ? { start: c.month(now, -1), end: c.month(now) } : period === "quarter" ? { start: c.month(now, -2), end } : { start: -Infinity, end }
  const own = data.matters.filter(({ matter }) => !matter.sample)
  const rows = own.filter(({ matter }) =>
    (chart.client === undefined || matter.client?.trim() === chart.client) &&
    (chart.matter === undefined || matter.title === chart.matter))
  const notes: string[] = []
  const result = (points: ChartPoint[], total: number, groups = points.length): ChartData => ({
    points, total, groups, notes, unavailable: Boolean(data.error), incomplete: chart.metric !== "matters" && rows.some((row) => row.unread),
    period: chart.metric === "drafts" ? "Latest finished run in each matter · all time" : chart.metric === "documents" ? "Current source documents · all time" : PERIOD_LABELS[period],
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
  if (chart.metric === "documents") {
    notes.push("Counts current source documents across all time; document dates are not used.")
    if (TIME.has(chart.groupBy)) return { ...result([], 0), unavailable: true, empty: "Documents cannot be dated for this view." }
  }
  if (chart.metric === "tokens") notes.push("Tokens are attributed to the date each session started.")
  if (rows.some((row) => row.unread) && chart.metric !== "matters") notes.push("Some matter activity could not be read. These figures may be incomplete.")
  if (chart.metric === "spend" && rows.some((row) => row.unpriced > 0)) notes.push("Some AI use has no price and is missing from this estimate.")
  const values = rows.flatMap((row) => events(row, chart.metric)).filter((entry) =>
    Number.isFinite(entry.value) && entry.value >= 0 &&
    (["drafts", "documents"].includes(chart.metric) || (entry.time !== undefined && Number.isFinite(entry.time) && entry.time >= range.start && entry.time < range.end)))
  const total = values.reduce((sum, entry) => sum + entry.value, 0)
  if (chart.groupBy === "none") return result([{ key: "total", label: chart.title, value: total }], total)
  if (TIME.has(chart.groupBy)) {
    const first = Number.isFinite(range.start) ? range.start : Math.min(c.day(now), ...values.map((entry) => entry.time!))
    const groups = new Map<number, ChartPoint>()
    for (let time = bucket(first, chart.groupBy, c); time < range.end;) {
      groups.set(time, { key: String(time), time, label: chart.groupBy === "month" ? dateLabel(time, true, zone) : `${chart.groupBy === "week" ? "Week of " : ""}${dateLabel(time, false, zone)}`, value: 0 })
      const next = nextBucket(time, chart.groupBy, c)
      if (!(next > time)) {
        notes.push("This calendar period could not be calculated safely. Try a different grouping.")
        return { ...result([], total), unavailable: true, empty: "This calendar period is unavailable." }
      }
      time = next
    }
    for (const entry of values) {
      const point = groups.get(bucket(entry.time!, chart.groupBy, c))
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
