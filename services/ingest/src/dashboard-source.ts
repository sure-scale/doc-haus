import { chartData } from "../../../dochaus/lib/dashboard-chart"
import { checkPlan, type Chart } from "../../../dochaus/lib/dashboard-plan"
import { engineV2Fetch } from "../../../dochaus/lib/engine-v2"
import { listMatters, listPlaybooks, workspaceRoot } from "./matter"
import { listClients } from "./clients"
import { existsSync, readdirSync } from "node:fs"
import path from "node:path"
import { leanDraftingStatus } from "./lean-drafting"
import { engineURL } from "./matter-session"
import { openDb, listSourceDocuments } from "./db"
import type { DashboardData, MatterActivity, Period } from "../../../dochaus/lib/dashboard-chart"
type Session = MatterActivity["sessions"][number] & { cost: number; tokens: { input: number; output: number; reasoning: number; cache?: { read: number; write: number } } }
type Row = Omit<MatterActivity, "sessions" | "drafting"> & { sessions: Session[]; drafting?: ReturnType<typeof leanDraftingStatus>; sessionUnread: boolean; draftingUnread: boolean; documentsUnread: boolean }

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
async function loadPractice(): Promise<{ matters: Row[] }> {
  const signal = AbortSignal.timeout(10_000)
  const matters = listMatters().filter(matter => !matter.sample)
  const rows: Row[] = new Array(matters.length)
  let index = 0
  await Promise.all(Array.from({ length: Math.min(4, matters.length) }, async () => {
    while (index < matters.length) {
      const i = index++; const matter = matters[i]!
      let sessions: Session[] = [], drafting: Row["drafting"], sessionUnread = false, draftingUnread = false, documentsUnread = false
      try { sessions = await pages(matter.dir, "/api/session", signal) as Session[] } catch { sessionUnread = true }
      try { drafting = leanDraftingStatus(matter.dir) as Row["drafting"] } catch { draftingUnread = true }
      let documents = 0
      try {
        if (existsSync(path.join(matter.dir, ".dochaus", "legal.db"))) {
          const db = openDb(matter.dir)
          try { documents = listSourceDocuments(db).length } finally { db.close() }
        }
      } catch { documentsUnread = true }
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
        } catch { costs.push(whole); sessionUnread = true }
      }
      const unpriced = sessions.filter(({ cost, tokens }) => !cost && ((tokens.input || 0) + (tokens.output || 0) + (tokens.reasoning || 0) + (tokens.cache?.read || 0) + (tokens.cache?.write || 0)) > 0).length
      rows[i] = { matter, sessions, drafting, documents, costs, unread: sessionUnread || draftingUnread || documentsUnread, sessionUnread, draftingUnread, documentsUnread, unpriced }
    }
  }))
  return { matters: rows }
}
let pending: Promise<{ data: { matters: Row[] }; updatedAt: number }> | undefined
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
  return { data: projection(chart, chartData(chart, scopedData(chart, snapshot.data), period as Period, Date.now(), timeZone)), updatedAt: snapshot.updatedAt }
}
function projection(chart: Chart, view: ReturnType<typeof chartData>) {
  const format = (value: number) => chart.metric === "spend" ? (value > 0 && value < .01 ? "< $0.01" : new Intl.NumberFormat("en-US", { style: "currency", currency: "USD" }).format(value)) : new Intl.NumberFormat("en-GB", { maximumFractionDigits: 0 }).format(value)
  const ranked = chart.groupBy === "client" || chart.groupBy === "matter"
  const interval = chart.groupBy === "day" ? "day" : chart.groupBy === "week" ? "wk" : "mo"
  const badge = ranked ? `${view.groups} ${chart.groupBy}${view.groups === 1 ? "" : "s"}` : chart.form === "line" ? `Peak: ${format(Math.max(0, ...view.points.map(point => point.value)))} / ${interval}` : ""
  const metric = chart.metric === "spend" ? "Estimated spend · USD" : chart.metric === "conversations" ? "Conversations started" : chart.metric === "matters" ? "Matters opened" : chart.metric === "tokens" ? "Tokens · by session start" : chart.metric === "documents" ? "Source documents" : "Drafted documents"
  return { points: view.total === 0 ? [] : view.points.map(({ key: _key, ...point }) => point), total: view.total, basis: `${metric}${view.incomplete ? " · partial" : ""}`, period: `${view.period}${chart.client ? ` · ${chart.client}` : chart.matter ? ` · ${chart.matter}` : ""}`, badge, unavailable: view.unavailable, notes: view.notes, empty: view.empty }
}

type Availability = { status: "available" | "empty" | "partial" | "unavailable"; usable: boolean; dataBacked?: boolean; reason?: string; total?: number }
function validateContext(period: unknown, zone: unknown): asserts period is Period {
  if (typeof period !== "string" || !PERIODS.has(period) || typeof zone !== "string" || !zone || zone.length > 100) throw new Error("Invalid dashboard source")
  try { new Intl.DateTimeFormat("en", { timeZone: zone }).format() } catch { throw new Error("Invalid dashboard source") }
}
function scopedData(chart: Chart, data: { matters: Row[] }) {
  return { matters: data.matters.map(row => ({ ...row, unread: chart.metric === "matters" ? false : chart.metric === "documents" ? row.documentsUnread : chart.metric === "drafts" ? row.draftingUnread : row.sessionUnread })) }
}
function chartAvailability(chart: Chart, data: { matters: Row[] }, period: Period, zone: string): Availability {
  const view = chartData(chart, scopedData(chart, data), period, Date.now(), zone)
  if (view.unavailable) return { status: "unavailable", usable: false, reason: view.empty }
  if (view.total > 0) return { status: view.incomplete ? "partial" : "available", usable: true, total: view.total }
  if (view.incomplete) return { status: "unavailable", usable: false, reason: "Some matter activity could not be read; no figures are available." }
  if (chart.metric === "spend" && view.notes.some(note => note.includes("no price"))) return { status: "unavailable", usable: false, reason: "AI use has no recorded price." }
  return { status: "empty", usable: false, total: 0, reason: view.empty }
}
function inPeriod(time: number, period: Period, zone: string, now: number) {
  const fmt = new Intl.DateTimeFormat("en-CA", { timeZone: zone, year: "numeric", month: "2-digit", day: "2-digit" })
  const ymd = (value: number) => { const p = Object.fromEntries(fmt.formatToParts(value).map(x => [x.type, Number(x.value)])); return { month: p.year * 12 + p.month, day: p.year * 10000 + p.month * 100 + p.day } }
  const date = ymd(time), today = ymd(now)
  if (date.day > today.day) return false
  if (period === "all") return true
  if (period === "month") return date.month === today.month
  if (period === "last-month") return date.month === today.month - 1
  return date.month >= today.month - 2 && date.month <= today.month
}
function withinDays(time: number, days: number, zone: string, now: number) {
  const fmt = new Intl.DateTimeFormat("en-CA", { timeZone: zone, year: "numeric", month: "2-digit", day: "2-digit" })
  const day = (value: number) => { const p = Object.fromEntries(fmt.formatToParts(value).map(x => [x.type, Number(x.value)])); return Date.UTC(p.year, p.month - 1, p.day) }
  const difference = day(now) - day(time)
  return difference >= 0 && difference < days * 86400000
}
function activityDays(zone: string, now: number) {
  const fmt = new Intl.DateTimeFormat("en-CA", { timeZone: zone, year: "numeric", month: "2-digit", day: "2-digit" })
  const p = Object.fromEntries(fmt.formatToParts(now).map(x => [x.type, Number(x.value)]))
  const weekday = new Date(Date.UTC(p.year, p.month - 1, p.day)).getUTCDay()
  return 365 + (weekday + 6) % 7
}
export function templateAvailability(id: string, data: { matters: Row[] }, period: Period, zone: string): Availability {
  const rows = data.matters, unread = rows.some(row => id === "attention" || id === "drafts" ? row.draftingUnread : row.sessionUnread), now = Date.now()
  const query = (metric: Chart["metric"], p: Chart["period"] = "dashboard") => chartAvailability({ type: "chart", title: "Total", metric, groupBy: "none", form: "stat", period: p }, data, period, zone)
  if (["actions", "ideas", "library", "sample", "setup", "journey"].includes(id)) return { status: "available", usable: true }
  if (id === "recent") return query("matters", "all")
  if (id === "drafts") return query("drafts", "all")
  if (id === "activity") {
    const count = rows.flatMap(row => row.sessions).filter(session => !session.parentID && !/^doc\.haus/i.test(session.title ?? "") && (session.metadata?.dochaus_role === undefined || session.metadata.dochaus_role === "lean-drafting-owner") && withinDays(session.time.created, activityDays(zone, now), zone, now)).length
    return count ? { status: unread ? "partial" : "available", usable: true, total: count } : { status: unread ? "unavailable" : "empty", usable: false, reason: "No recent conversations." }
  }
  if (id === "week") {
    const total = rows.flatMap(row => row.costs).filter(cost => withinDays(cost.time, 7, zone, now)).reduce((sum, cost) => sum + cost.cost, 0)
    return total > 0 ? { status: unread ? "partial" : "available", usable: true, total } : { status: unread ? "unavailable" : "empty", usable: false, reason: "No priced spend in the last seven days." }
  }
  if (["spend", "matter-spend", "trend"].includes(id)) return query("spend")
  if (id === "matter-cost") return query("spend", "all")
  if (id === "draft-cost") {
    const runs = rows.filter(row => row.drafting?.status === "completed" && row.drafting.documents?.status === "completed" && row.drafting.documents.documents.some(document => document.kind === "body" && document.deliverable))
    const correlated = runs.some(row => {
      const draft = row.drafting!
      const ids = new Set([draft.owner_session_id, ...Object.values(draft.stages).flatMap(stage => [...(stage?.session_ids ?? []), ...(stage?.session_id ? [stage.session_id] : [])])])
      let changed = true
      while (changed) { changed = false; for (const session of row.sessions) if (session.parentID && ids.has(session.parentID) && !ids.has(session.id)) { ids.add(session.id); changed = true } }
      return row.sessions.some(session => ids.has(session.id) && session.cost > 0)
    })
    const uncertain = rows.some(row => row.draftingUnread || row.sessionUnread)
    return correlated ? { status: uncertain ? "partial" : "available", usable: true } : { status: uncertain ? "unavailable" : "empty", usable: false, reason: "Finished drafts with priced activity in the same run are needed." }
  }
  if (id === "tokens") {
    const total = rows.flatMap(row => row.sessions).filter(session => inPeriod(session.time.created, period, zone, now)).reduce((sum, session) => sum + (session.tokens.input || 0) + (session.tokens.output || 0) + (session.tokens.reasoning || 0) + (session.tokens.cache?.read || 0) + (session.tokens.cache?.write || 0), 0)
    return total ? { status: unread ? "partial" : "available", usable: true, total } : { status: unread ? "unavailable" : "empty", usable: false, reason: unread ? "Some activity could not be read." : "No tokens used in this period." }
  }
  if (id === "attention") {
    const found = rows.some(({ drafting: d }) => d && !d.stopped && (d.status === "failed" || d.status === "blocked" || d.status === "running" && !d.in_flight || d.questions.some(q => q.status === "open" && !d.queued_answer_ids?.includes(q.id))))
    return found ? { status: unread ? "partial" : "available", usable: true } : { status: unread ? "unavailable" : "empty", usable: false, reason: unread ? "Some drafting activity could not be read." : "No matters need attention." }
  }
  return { status: "unavailable", usable: false, reason: "This widget is not supported." }
}
function libraryCounts() {
  const count = (run: () => number): number | null => { try { return run() } catch { return null } }
  return { clients: count(() => listClients().length), templates: count(() => { const dir = path.join(workspaceRoot(), ".templates"); return existsSync(dir) ? readdirSync(dir).filter(name => name.endsWith(".docx")).length : 0 }), playbooks: count(() => listPlaybooks().length) }
}
export async function dashboardInventory(period: unknown, timeZone: unknown) {
  validateContext(period, timeZone)
  const zone = timeZone as string, snapshot = await practice(), data = snapshot.data
  const metrics = Object.fromEntries((["matters", "conversations", "spend", "drafts", "tokens", "documents"] as const).map(metric => {
    const periods = Object.fromEntries((["month", "last-month", "quarter", "all"] as const).map(p => [p, chartAvailability({ type: "chart", title: "Total", metric, groupBy: "none", form: "stat", period: p }, data, p, zone)]))
    const times = metric === "matters" ? data.matters.map(row => row.matter.created_at) : metric === "spend" ? data.matters.flatMap(row => row.costs.map(cost => cost.time)) : metric === "conversations" ? data.matters.flatMap(row => row.sessions.filter(session => !session.parentID && !/^doc\.haus/i.test(session.title ?? "") && (session.metadata?.dochaus_role === undefined || session.metadata.dochaus_role === "lean-drafting-owner")).map(session => session.time.created)) : metric === "tokens" ? data.matters.flatMap(row => row.sessions.map(session => session.time.created)) : []
    const valid = times.filter(Number.isFinite)
    return [metric, { ...periods[period], periods, ...(valid.length ? { firstAt: Math.min(...valid), lastAt: Math.max(...valid) } : {}), ...(metric === "drafts" ? { note: "Latest finished run per matter; all time only." } : metric === "documents" ? { note: "Current source documents; all time only." } : metric === "tokens" ? { note: "Session totals by session start date." } : {}) }]
  }))
  const { TEMPLATES } = await import("../../../dochaus/lib/dashboard-plan")
  return { updatedAt: snapshot.updatedAt, metrics, templates: Object.fromEntries(Object.keys(TEMPLATES).map(id => [id, templateAvailability(id, data, period, zone)])), library: libraryCounts(), filters: { clients: [...new Set(data.matters.map(row => row.matter.client?.trim()).filter((name): name is string => Boolean(name)))], matters: [...new Set(data.matters.map(row => row.matter.title))] } }
}
export async function validateDashboardSources(input: unknown) {
  if (!input || typeof input !== "object" || Array.isArray(input)) throw new Error("Invalid dashboard source")
  const value = input as Record<string, unknown>
  if (Object.keys(value).some(key => !["widgets", "period", "timeZone"].includes(key))) throw new Error("Invalid dashboard source")
  validateContext(value.period, value.timeZone)
  const checked = checkPlan({ widgets: value.widgets })
  if (!checked.plan) throw new Error("Invalid dashboard source")
  const snapshot = await practice(), library = libraryCounts()
  const results = checked.plan.widgets.map((widget, index) => {
    const result = widget.type === "chart" ? chartAvailability(widget, snapshot.data, value.period as Period, value.timeZone as string) : templateAvailability(widget.type, snapshot.data, value.period as Period, value.timeZone as string)
    return { index, ...result, dataBacked: result.usable && (widget.type === "chart" || widget.type === "library" && (snapshot.data.matters.length > 0 || Object.values(library).some(count => count !== null && count > 0)) || !["actions", "ideas", "library", "sample", "setup", "journey"].includes(widget.type)) }
  })
  return { updatedAt: snapshot.updatedAt, results }
}
