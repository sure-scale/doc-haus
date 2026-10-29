import { chartData } from "../../../dochaus/lib/dashboard-chart"
import { checkPlan, type Chart } from "../../../dochaus/lib/dashboard-plan"
import { engineV2Fetch } from "../../../dochaus/lib/engine-v2"
import { listMatters } from "./matter"
import { leanDraftingStatus } from "./lean-drafting"
import { engineURL } from "./matter-session"
import type { DashboardData, MatterActivity, Period } from "../../../dochaus/lib/dashboard-chart"
type Session = MatterActivity["sessions"][number] & { cost: number; tokens: { input: number; output: number; reasoning: number; cache?: { read: number; write: number } } }
type Row = Omit<MatterActivity, "sessions"> & { sessions: Session[] }

const PERIODS = new Set(["month", "last-month", "quarter", "all"])
async function pages(directory: string, pathname: string, signal: AbortSignal): Promise<unknown[]> {
  let cursor: string | undefined
  const all: unknown[] = [], seen = new Set<string>()
  for (let page = 0; page < 100; page++) {
    let next: string | undefined
    const transport = Object.assign(async (input: Parameters<typeof fetch>[0], init?: RequestInit) => {
      const response = await fetch(input, init)
      if (response.ok && new URL(String(input)).pathname === pathname) {
        const envelope: any = await response.clone().json()
        next = typeof envelope?.cursor?.next === "string" ? envelope.cursor.next : undefined
      }
      return response
    }, { preconnect: fetch.preconnect }) as typeof fetch
    const url = new URL(pathname, engineURL())
    url.searchParams.set("limit", "200")
    if (cursor) url.searchParams.set("cursor", cursor)
    else url.searchParams.set("order", pathname === "/api/session" ? "desc" : "asc")
    const response = await engineV2Fetch(transport)(url, { headers: { "x-opencode-directory": directory }, signal })
    if (!response.ok) throw new Error("Engine read failed")
    const rows: unknown = await response.json()
    if (!Array.isArray(rows)) throw new Error("Invalid engine response")
    all.push(...rows)
    if (!next) return all
    if (seen.has(next)) throw new Error("Repeated cursor")
    seen.add(next); cursor = next
  }
  throw new Error("Engine page limit reached")
}
async function loadPractice(): Promise<DashboardData> {
  const signal = AbortSignal.timeout(10_000)
  const matters = listMatters().filter(matter => !matter.sample)
  const rows: Row[] = new Array(matters.length)
  let index = 0
  await Promise.all(Array.from({ length: Math.min(4, matters.length) }, async () => {
    while (index < matters.length) {
      const i = index++; const matter = matters[i]!
      let sessions: Session[] = [], drafting: MatterActivity["drafting"], unread = false
      try { sessions = await pages(matter.dir, "/api/session", signal) as Session[] } catch { unread = true }
      try { drafting = leanDraftingStatus(matter.dir) as MatterActivity["drafting"] } catch { unread = true }
      const costs: MatterActivity["costs"] = []
      for (const session of sessions) {
        const whole = { time: session.time.created, cost: session.cost || 0, session: session.id }
        if (!session.cost) { costs.push(whole); continue }
        try {
          const messages = await pages(matter.dir, `/api/session/${encodeURIComponent(session.id)}/message`, signal) as { info: { cost?: number; time: { created: number } } }[]
          const entries = messages.flatMap(({ info }) => info.cost ? [{ time: info.time.created, cost: info.cost, session: session.id }] : [])
          const rest = session.cost - entries.reduce((sum, entry) => sum + entry.cost, 0)
          if (rest > .0001) entries.push({ time: session.time.updated, cost: rest, session: session.id })
          costs.push(...entries)
        } catch { costs.push(whole); unread = true }
      }
      const unpriced = sessions.filter(({ cost, tokens }) => !cost && ((tokens.input || 0) + (tokens.output || 0) + (tokens.reasoning || 0) + (tokens.cache?.read || 0) + (tokens.cache?.write || 0)) > 0).length
      rows[i] = { matter, sessions, drafting, costs, unread, unpriced }
    }
  }))
  return { matters: rows }
}
let pending: Promise<{ data: DashboardData; updatedAt: number }> | undefined
function practice() {
  if (!pending) {
    const current = loadPractice().then(data => ({ data, updatedAt: Date.now() }))
    pending = current
    void current.finally(() => { if (pending === current) pending = undefined }).catch(() => {})
  }
  return pending
}
export async function dashboardChartSource(raw: string | undefined, period: string | undefined, timeZone: string | undefined) {
  if (!raw || raw.length > 8192 || !period || !PERIODS.has(period) || !timeZone || timeZone.length > 100) throw new Error("Invalid dashboard source")
  try { new Intl.DateTimeFormat("en", { timeZone }).format() } catch { throw new Error("Invalid dashboard source") }
  let input: unknown
  try { input = JSON.parse(raw) } catch { throw new Error("Invalid dashboard source") }
  const chart = checkPlan({ widgets: [input] }).plan?.widgets[0]
  if (!chart || chart.type !== "chart" || chart.source?.kind !== "local-query" || chart.source.query !== "practice-chart" || chart.source.version !== 1) throw new Error("Invalid dashboard source")
  const snapshot = await practice()
  return { data: projection(chart, chartData(chart, snapshot.data, period as Period, Date.now(), timeZone)), updatedAt: snapshot.updatedAt }
}
function projection(chart: Chart, view: ReturnType<typeof chartData>) {
  const format = (value: number) => chart.metric === "spend" ? (value > 0 && value < .01 ? "< $0.01" : new Intl.NumberFormat("en-US", { style: "currency", currency: "USD" }).format(value)) : new Intl.NumberFormat("en-GB", { maximumFractionDigits: 0 }).format(value)
  const ranked = chart.groupBy === "client" || chart.groupBy === "matter"
  const interval = chart.groupBy === "day" ? "day" : chart.groupBy === "week" ? "wk" : "mo"
  const badge = ranked ? `${view.groups} ${chart.groupBy}${view.groups === 1 ? "" : "s"}` : chart.form === "line" ? `Peak: ${format(Math.max(0, ...view.points.map(point => point.value)))} / ${interval}` : ""
  const metric = chart.metric === "spend" ? "Estimated spend · USD" : chart.metric === "conversations" ? "Conversations started" : chart.metric === "matters" ? "Matters opened" : "Drafted documents"
  return { points: view.total === 0 ? [] : view.points.map(({ key: _key, ...point }) => point), total: view.total, basis: `${metric}${view.incomplete ? " · partial" : ""}`, period: `${view.period}${chart.client ? ` · ${chart.client}` : chart.matter ? ` · ${chart.matter}` : ""}`, badge, unavailable: view.unavailable, notes: view.notes, empty: view.empty }
}
