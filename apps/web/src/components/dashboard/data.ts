import { useEffect, useRef, useState } from "react"
import type { SessionMessageInfo } from "@opencode/client"
import { getMatter, listJurisdictions, listMatters, type Jurisdiction, type Matter } from "../../api/ingest/matters"
import { getDraftingStatus } from "../../api/ingest/drafting-status"
import { leanStageSessions, type LeanDocument, type LeanDraftingStatus } from "../../api/ingest/lean"
import { getPreferences, type DraftingPreferences } from "../../api/ingest/preferences"
import { fetchWordSetup, type WordSetupStatus } from "../../api/ingest/word"
import { listPlaybookLibrary, listTemplates } from "../../api/ingest/library"
import { listClients } from "../../api/ingest/clients"
import { listMessageInfos, listSessions } from "../../api/opencode/sessions"
import { matterClient, type Client } from "../../api/opencode/client"
import { checklistFindingsCount, cleanTitle, openQuestions } from "../drafting/model"

export type Session = Awaited<ReturnType<typeof listSessions>>[number]

// One piece of spend at the moment it happened. A conversation that runs over
// several days is spread across them, one entry per priced model call.
export type CostEntry = { time: number; cost: number; session: string }

// One matter as the dashboard sees it: every engine session in its folder
// (subagent runs included, each carrying its own cost) and its drafting run.
// `documents` is read only for the newest few of the lawyer's own matters.
export type MatterActivity = {
  matter: Matter
  sessions: Session[]
  drafting?: LeanDraftingStatus
  documents?: number
  costs: CostEntry[]
  // Sessions, cost timing or the drafting run could not be read completely.
  unread: boolean
  // Sessions that used tokens but report no cost: see isUnpriced.
  unpriced: number
}

export type DashboardData = {
  matters: MatterActivity[]
  preferences?: DraftingPreferences
  // The firm's preferences could not be read, as opposed to never saved.
  preferencesFailed?: boolean
  word?: WordSetupStatus
  templates?: number
  playbooks?: number
  clients?: number
  jurisdictions?: Jurisdiction[]
  // Set when the matter list itself could not be read, so an unreachable
  // service is never shown as an empty practice.
  error?: string
}

// Calendar windows, because lawyers bill by the month.
export type Period = "month" | "last-month" | "quarter" | "all"
export const PERIODS: Period[] = ["month", "last-month", "quarter", "all"]
export const PERIOD_LABELS: Record<Period, string> = { month: "This month", "last-month": "Last month", quarter: "Last 3 months", all: "All time" }

// [start, end) in local-time instants.
export type Range = { start: number; end: number }

const DAY = 86_400_000

// Each matter costs two reads (its sessions and its drafting status); a large
// practice must not open hundreds of requests at once.
async function mapLimited<T, R>(items: T[], limit: number, work: (item: T) => Promise<R>) {
  const results: R[] = new Array(items.length)
  let next = 0
  await Promise.all(Array.from({ length: Math.min(limit, items.length) }, async () => {
    while (next < items.length) {
      const index = next++
      results[index] = await work(items[index])
    }
  }))
  return results
}

const settled = <T,>(promise: Promise<T>) => promise.then((value) => value, () => undefined)

// Each priced message at its own time; whatever the messages do not account
// for lands when the session was last active.
export function messageCosts(session: Session, messages: SessionMessageInfo[]): CostEntry[] {
  const entries = messages.flatMap((message) =>
    "cost" in message && message.cost ? [{ time: message.time.created, cost: message.cost, session: session.id }] : [])
  const rest = (session.cost || 0) - entries.reduce((sum, entry) => sum + entry.cost, 0)
  if (rest > 0.0001) entries.push({ time: session.time.updated, cost: rest, session: session.id })
  return entries
}

// A session that stayed within one day needs no message read. If the read
// fails the session's spend falls on the day it started and is marked partial.
async function sessionCosts(client: Client, session: Session): Promise<{ entries: CostEntry[]; unread: boolean }> {
  const whole = [{ time: session.time.created, cost: session.cost || 0, session: session.id }]
  if (!session.cost || startOfDay(session.time.created) === startOfDay(session.time.updated)) return { entries: whole, unread: false }
  return listMessageInfos(client, session.id).then((messages) => ({ entries: messageCosts(session, messages), unread: false }), () => ({ entries: whole, unread: true }))
}

async function loadMatter(matter: Matter, withDocuments: boolean): Promise<MatterActivity> {
  const client = matterClient(matter.dir)
  const [sessions, drafting, documents] = await Promise.allSettled([
    listSessions(client),
    getDraftingStatus(matter.id),
    withDocuments ? getMatter(matter.id).then((detail) => detail.documents.length) : Promise.resolve(undefined),
  ])
  const list = sessions.status === "fulfilled" ? sessions.value : []
  const costs = await mapLimited(list, 2, (session) => sessionCosts(client, session))
  return {
    matter,
    sessions: list,
    drafting: drafting.status === "fulfilled" ? drafting.value : undefined,
    documents: documents.status === "fulfilled" ? documents.value : undefined,
    costs: costs.flatMap((result) => result.entries),
    unread: sessions.status === "rejected" || drafting.status === "rejected" || costs.some((result) => result.unread),
    unpriced: list.filter(isUnpriced).length,
  }
}

export async function loadDashboard(): Promise<DashboardData> {
  const [matters, preferences, word, templates, playbooks, clients, jurisdictions] = await Promise.all([
    listMatters().then((list) => ({ list }), () => ({ list: [] as Matter[], error: "Could not load matters" })),
    getPreferences().then((value) => ({ value, failed: false }), () => ({ value: undefined, failed: true })),
    settled(fetchWordSetup()),
    settled(listTemplates().then((result) => result.templates.length)),
    settled(listPlaybookLibrary().then((result) => result.playbooks.length)),
    settled(listClients().then((result) => result.length)),
    settled(listJurisdictions()),
  ])
  // First steps only ask whether documents were added, so the file lists of
  // the newest few own matters are enough.
  const byAge = matters.list.filter((matter) => !matter.sample).sort((a, b) => b.created_at - a.created_at)
  const newest = new Set(byAge.slice(0, 3).map((matter) => matter.id))
  const rows = await mapLimited(matters.list, 4, (matter) => loadMatter(matter, newest.has(matter.id)))
  // Before a first draft, an older matter may hold the only documents: look
  // further back, one at a time, until one does.
  if (!rows.some((row) => row.documents || draftedCount(row.drafting))) {
    for (const matter of byAge.slice(3)) {
      const row = rows.find((other) => other.matter.id === matter.id)
      if (!row) continue
      row.documents = await getMatter(matter.id).then((detail) => detail.documents.length, () => undefined)
      if (row.documents) break
    }
  }
  return {
    matters: rows,
    preferences: preferences.value,
    ...(preferences.failed ? { preferencesFailed: true } : {}),
    word, templates, playbooks, clients, jurisdictions,
    ...("error" in matters ? { error: matters.error } : {}),
  }
}

// One reader for the whole dashboard. Invalidation during a read coalesces into
// one subsequent read; an obsolete result never replaces the newer request.
export function dashboardRefreshController<T>(options: {
  load: () => Promise<T>
  publish: (value: T) => void
  pending: () => void
  visible: () => boolean
  intervalMs?: number
}) {
  let stopped = false
  let running = false
  let queued = false
  let timer: ReturnType<typeof setTimeout> | undefined
  const schedule = () => {
    clearTimeout(timer)
    if (!stopped && options.visible()) timer = setTimeout(() => { if (options.visible()) request() }, options.intervalMs ?? 30_000)
  }
  async function run() {
    running = true
    options.pending()
    try {
      const value = await options.load()
      if (!stopped && !queued) options.publish(value)
    } finally {
      running = false
      if (stopped) return
      if (queued) { queued = false; void run() }
      else schedule()
    }
  }
  function request() {
    if (stopped) return
    clearTimeout(timer)
    if (running) queued = true
    else void run()
  }
  return { request, onVisible() { if (options.visible()) request() }, stop() { stopped = true; queued = false; clearTimeout(timer) } }
}

// Retain previous figures during reads; setup changes and manual refresh still
// invalidate them, as do returning to a visible dashboard and periodic reads.
export function useDashboardData(setup: unknown, refresh: number) {
  const [data, setData] = useState<DashboardData>()
  const [loading, setLoading] = useState(true)
  const controller = useRef<ReturnType<typeof dashboardRefreshController<DashboardData>> | undefined>(undefined)
  const inputs = useRef({ setup, refresh })
  useEffect(() => {
    const reader = dashboardRefreshController({
      load: () => loadDashboard().catch((): DashboardData => ({ matters: [], error: "Could not load matters" })),
      pending: () => setLoading(true),
      publish: (next) => { setData(next); setLoading(false) },
      visible: () => document.visibilityState === "visible",
    })
    controller.current = reader
    const onVisible = reader.onVisible
    document.addEventListener("visibilitychange", onVisible)
    window.addEventListener("focus", onVisible)
    reader.request()
    return () => { reader.stop(); controller.current = undefined; document.removeEventListener("visibilitychange", onVisible); window.removeEventListener("focus", onVisible) }
  }, [])
  useEffect(() => {
    if (inputs.current.setup !== setup || inputs.current.refresh !== refresh) {
      inputs.current = { setup, refresh }
      controller.current?.request()
    }
  }, [setup, refresh])
  return { data, loading, reload: () => controller.current?.request() }
}

export function startOfDay(time: number) {
  const date = new Date(time)
  date.setHours(0, 0, 0, 0)
  return date.getTime()
}

// Local midnight `days` calendar days from `day`. Days are counted on the
// calendar, not in 24-hour steps, so a clock change never shifts a day.
export function addDays(day: number, days: number) {
  const date = new Date(day)
  date.setDate(date.getDate() + days)
  return startOfDay(date.getTime())
}

// Calendar days from `first` (a local midnight) to the day `time` falls on.
// Rounded, because a day with a clock change is 23 or 25 hours long.
export const dayIndex = (first: number, time: number) => Math.round((startOfDay(time) - first) / DAY)

// Local midnight on the 1st of the month `months` from the one `time` is in;
// Date rolls a negative month back into the previous year.
function monthStart(time: number, months = 0) {
  const date = new Date(time)
  return new Date(date.getFullYear(), date.getMonth() + months, 1).getTime()
}

// Windows run to the end of today, so spend made later today still counts.
export function periodRange(period: Period, now: number): Range {
  const end = addDays(startOfDay(now), 1)
  if (period === "month") return { start: monthStart(now), end }
  if (period === "last-month") return { start: monthStart(now, -1), end: monthStart(now) }
  if (period === "quarter") return { start: monthStart(now, -2), end }
  return { start: -Infinity, end }
}

// The calendar window before this one, for the change badge; all time has none.
export function previousRange(period: Period, now: number): Range | undefined {
  if (period === "month") return { start: monthStart(now, -1), end: monthStart(now) }
  if (period === "last-month") return { start: monthStart(now, -2), end: monthStart(now, -1) }
  if (period === "quarter") return { start: monthStart(now, -5), end: monthStart(now, -2) }
  return undefined
}

// Today and the `days - 1` calendar days before it.
export function lastDays(days: number, now: number): Range {
  const today = startOfDay(now)
  return { start: addDays(today, 1 - days), end: addDays(today, 1) }
}

const within = (time: number, range: Range) => time >= range.start && time < range.end

// Spend figures leave out the fictional sample matter.
const own = (matters: MatterActivity[]) => matters.filter((row) => !row.matter.sample)

const spent = (costs: CostEntry[], range?: Range) =>
  costs.reduce((sum, entry) => sum + (!range || within(entry.time, range) ? entry.cost : 0), 0)

// Where a window's charts begin: all time starts on the day of the first own
// spend or conversation, or today when there is none.
export function chartStart(matters: MatterActivity[], range: Range) {
  if (Number.isFinite(range.start)) return range.start
  let first = addDays(range.end, -1)
  for (const row of own(matters)) {
    for (const entry of row.costs) if (entry.cost > 0) first = Math.min(first, entry.time)
    for (const session of row.sessions) if (isConversation(session)) first = Math.min(first, session.time.created)
  }
  return startOfDay(first)
}

// A conversation the lawyer started, as the sidebar lists them: not a subagent
// run, not a session the runtime owns (drafting stages, the narrator) unless it
// is a drafting run's own conversation, and not a system-titled one such as a
// review grid's cell sessions. Their cost still counts towards spend.
export function isConversation(session: Session) {
  if (session.parentID || /^doc\.haus/i.test(session.title ?? "")) return false
  const role = typeof session.metadata === "object" && session.metadata !== null && "dochaus_role" in session.metadata ? session.metadata.dochaus_role : undefined
  return role === undefined || role === "lean-drafting-owner"
}

// OpenCode prices a model it has no price for at 0, so a session that used
// tokens yet cost nothing ran on an unpriced model; it was not free.
function isUnpriced({ cost, tokens }: Session) {
  const used = (tokens.input || 0) + (tokens.output || 0) + (tokens.reasoning || 0) + (tokens.cache?.read || 0) + (tokens.cache?.write || 0)
  return !cost && used > 0
}

export function totalCost(sessions: Session[]) {
  return sessions.reduce((sum, session) => sum + (session.cost || 0), 0)
}

// Own spend for sparklines.
export function costEntries(matters: MatterActivity[]) {
  return own(matters).flatMap((row) => row.costs.map((entry) => ({ time: entry.time, value: entry.cost })))
}

// A stat tile's small trend: [start, end) split into even buckets.
export function bucketed(values: { time: number; value: number }[], start: number, end: number, count: number) {
  const span = Math.max(DAY, end - start)
  const buckets = new Array<number>(count).fill(0)
  for (const { time, value } of values) {
    if (time < start || time >= end) continue
    buckets[Math.min(count - 1, Math.floor(((time - start) / span) * count))] += value
  }
  return buckets
}

export type DaySpend = { day: number; cost: number; conversations: number; matters: { matter: Matter; cost: number }[] }

// Own spend for each calendar day in the range (oldest first), each entry on
// the day it was spent, with the matters that spent it, largest first.
export function dailySpend(matters: MatterActivity[], range: Range): DaySpend[] {
  const days = dayIndex(range.start, range.end)
  const buckets = Array.from({ length: days }, (_, index) => ({ day: addDays(range.start, index), cost: 0, conversations: 0, byMatter: new Map<Matter, number>() }))
  for (const { matter, sessions, costs } of own(matters)) {
    for (const entry of costs) {
      const index = dayIndex(range.start, entry.time)
      if (!entry.cost || index < 0 || index >= days) continue
      buckets[index].cost += entry.cost
      buckets[index].byMatter.set(matter, (buckets[index].byMatter.get(matter) ?? 0) + entry.cost)
    }
    for (const session of sessions) {
      const index = dayIndex(range.start, session.time.created)
      if (index >= 0 && index < days && isConversation(session)) buckets[index].conversations += 1
    }
  }
  return buckets.map(({ byMatter, ...day }) => ({
    ...day,
    matters: [...byMatter].map(([matter, cost]) => ({ matter, cost })).sort((a, b) => b.cost - a.cost),
  }))
}

// Own conversations started each day — the activity calendar's cells.
export function dailyConversations(matters: MatterActivity[], first: number, days: number) {
  const counts = new Array<number>(days).fill(0)
  for (const session of own(matters).flatMap((row) => row.sessions)) {
    if (!isConversation(session)) continue
    const index = dayIndex(first, session.time.created)
    if (index >= 0 && index < days) counts[index] += 1
  }
  return counts
}

export type MatterCost = { matter: Matter; cost: number; conversations: number }

export function costByMatter(matters: MatterActivity[], range: Range): MatterCost[] {
  return own(matters)
    .map(({ matter, sessions, costs }) => ({
      matter,
      cost: spent(costs, range),
      conversations: sessions.filter((session) => isConversation(session) && within(session.time.created, range)).length,
    }))
    .filter((row) => row.cost > 0)
    .sort((a, b) => b.cost - a.cost)
}

export type ClientCost = { client?: string; cost: number; matters: number }

const clientOf = (matter: Matter) => matter.client?.trim() || undefined

// Matters with no client recorded are grouped together.
export function costByClient(rows: MatterCost[]): ClientCost[] {
  const groups = new Map<string | undefined, ClientCost>()
  for (const row of rows) {
    const client = clientOf(row.matter)
    const group = groups.get(client) ?? { ...(client ? { client } : {}), cost: 0, matters: 0 }
    group.cost += row.cost
    group.matters += 1
    groups.set(client, group)
  }
  return [...groups.values()].sort((a, b) => b.cost - a.cost)
}

// RFC 4180: quote a field holding a comma, quote or line break, doubling
// quotes. A leading = + - @ or tab is escaped, so a spreadsheet opening the
// export never reads a matter title as a formula.
function csvField(value: string) {
  const text = /^[=+\-@\t\r]/.test(value) ? `'${value}` : value
  return /[",\r\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text
}

// A spend export for billing, grouped by client (none last), largest first.
export function spendCsv(rows: MatterCost[]) {
  const sorted = [...rows].sort((a, b) => {
    const left = clientOf(a.matter)
    const right = clientOf(b.matter)
    if (left !== right) return left === undefined ? 1 : right === undefined ? -1 : left.localeCompare(right)
    return b.cost - a.cost
  })
  const lines = [
    ["Client", "Matter", "Reference", "Conversations", "Spend (USD)"],
    ...sorted.map((row) => [clientOf(row.matter) ?? "", row.matter.title, row.matter.reference ?? "", String(row.conversations), row.cost.toFixed(2)]),
  ]
  return lines.map((line) => line.map(csvField).join(",")).join("\r\n") + "\r\n"
}

// Every session a drafting run used: the conversation that owns it, each
// stage's sessions, and every subagent those started.
export function runSessions(status: LeanDraftingStatus, sessions: Session[]) {
  const ids = new Set([status.owner_session_id, ...Object.values(status.stages).flatMap((row) => leanStageSessions(row))])
  let grew = true
  while (grew) {
    grew = false
    for (const session of sessions) {
      if (session.parentID && ids.has(session.parentID) && !ids.has(session.id)) {
        ids.add(session.id)
        grew = true
      }
    }
  }
  return sessions.filter((session) => ids.has(session.id))
}

// A completed run's delivered bodies, as the run's own header counts them.
// Redlines, notes and combined files are copies of those.
function deliveredBodies(drafting: LeanDraftingStatus | undefined) {
  if (drafting?.documents?.status !== "completed") return []
  return drafting.documents.documents.filter((row) => row.kind === "body" && row.deliverable)
}

export function draftedCount(drafting: LeanDraftingStatus | undefined) {
  return deliveredBodies(drafting).length
}

export type Summary = {
  spend: number
  previous?: number
  conversations: number
  activeMatters: number
  unpriced: number
  tokens: { input: number; cached: number; output: number; reasoning: number }
}

// Spend counts when it happened; conversations, tokens and unpriced sessions
// count in the window they started in.
export function summarize(matters: MatterActivity[], period: Period, now: number): Summary {
  const range = periodRange(period, now)
  const before = previousRange(period, now)
  const rows = own(matters)
  const costs = rows.flatMap((row) => row.costs)
  const sessions = rows.flatMap((row) => row.sessions).filter((session) => within(session.time.created, range))
  const tokens = { input: 0, cached: 0, output: 0, reasoning: 0 }
  for (const { tokens: used } of sessions) {
    tokens.input += (used.input || 0) + (used.cache?.write || 0)
    tokens.cached += used.cache?.read || 0
    tokens.output += used.output || 0
    tokens.reasoning += used.reasoning || 0
  }
  return {
    spend: spent(costs, range),
    ...(before ? { previous: spent(costs, before) } : {}),
    conversations: sessions.filter(isConversation).length,
    activeMatters: costByMatter(matters, range).length,
    unpriced: sessions.filter(isUnpriced).length,
    tokens,
  }
}

// All time: what an own matter that has spent anything has cost on average.
export function averageMatterCost(matters: MatterActivity[]) {
  const costs = own(matters).map((row) => spent(row.costs)).filter((cost) => cost > 0)
  return costs.length ? { average: costs.reduce((sum, cost) => sum + cost, 0) / costs.length, matters: costs.length } : undefined
}

// What one drafted document costs: every session of each own matter's latest
// completed run over the documents it delivered. The ingest status keeps only
// the latest run, so earlier runs on the same matter are not seen.
export function averageDraftCost(matters: MatterActivity[]) {
  let cost = 0
  let documents = 0
  let runs = 0
  for (const { sessions, drafting } of own(matters)) {
    const count = draftedCount(drafting)
    if (drafting?.status !== "completed" || !count) continue
    cost += totalCost(runSessions(drafting, sessions))
    documents += count
    runs += 1
  }
  return documents ? { average: cost / documents, documents, runs } : undefined
}

export type DraftState = "ready" | "review" | "findings" | "changed"

export type Draft = { matter: Matter; session: string; finished: number; titles: string[]; state: DraftState; findings?: number; cost: number }

// A delivered body's name: its own title, the plan's, or its file name.
const bodyTitle = (status: LeanDraftingStatus, row: LeanDocument) =>
  cleanTitle(row.title ?? status.plan?.documents.find((planned) => planned.id === row.document)?.title ?? row.deliverable!.replace(/\.docx$/i, ""))

// As the run header settles it: a draft edited or removed since the final
// check is no longer what it checked, and unresolved findings come before the
// check's delivery verdict.
function draftState(status: LeanDraftingStatus): DraftState {
  const after = status.changed_after_check
  if (after && (after.documents.length || after.missing.length)) return "changed"
  if (checklistFindingsCount(status) !== undefined) return "findings"
  return status.gate?.delivery === "execution_ready" ? "ready" : "review"
}

// Own matters' delivered drafts, newest first; one per matter, its latest run.
export function drafts(matters: MatterActivity[], limit = Infinity): Draft[] {
  return own(matters)
    .flatMap(({ matter, sessions, drafting }): Draft[] => {
      const bodies = deliveredBodies(drafting)
      if (drafting?.status !== "completed" || !bodies.length) return []
      const state = draftState(drafting)
      // 0 means the count did not reach the status; the widget then says "findings" alone.
      const findings = state === "findings" ? checklistFindingsCount(drafting) : undefined
      return [{
        matter,
        session: drafting.owner_session_id,
        finished: Date.parse(drafting.updated_at),
        titles: bodies.map((row) => bodyTitle(drafting, row)),
        state,
        ...(findings ? { findings } : {}),
        cost: totalCost(runSessions(drafting, sessions)),
      }]
    })
    .sort((a, b) => b.finished - a.finished)
    .slice(0, limit)
}

export type AttentionItem = { matter: Matter; kind: "provider" | "failed" | "questions" | "stalled" | "blocked" | "paused"; count?: number; session: string }

// Drafting runs that wait on the lawyer, most urgent first. The sample matter
// counts: the lawyer may be trying a draft on it.
export function attention(matters: MatterActivity[]): AttentionItem[] {
  const rank = { provider: 0, failed: 1, questions: 2, stalled: 3, blocked: 4, paused: 5 }
  return matters
    .flatMap(({ matter, drafting }): AttentionItem[] => {
      // A run the lawyer stopped settles as failed but needs nothing more.
      if (!drafting || drafting.stopped) return []
      const session = drafting.owner_session_id
      // Answers already queued for the run wait on it, not on the lawyer.
      const open = openQuestions(drafting).length
      if (drafting.status === "failed") return [{ matter, kind: drafting.provider_stop ? "provider" : "failed", session }]
      if (open) return [{ matter, kind: "questions", count: open, session }]
      // Still marked running with nothing working on it (ingest restarted):
      // it moves again only when the lawyer presses Continue.
      if (drafting.status === "running" && !drafting.in_flight) return [{ matter, kind: "stalled", session }]
      if (drafting.status === "blocked") return [{ matter, kind: drafting.paused_at ? "paused" : "blocked", session }]
      return []
    })
    .sort((a, b) => rank[a.kind] - rank[b.kind])
}

// Ingest stores some times as epoch seconds.
const epochMs = (value: number | undefined) => (!value ? 0 : value < 1e11 ? value * 1000 : value)

// When a matter was last worked in: its newest conversation, or its opening.
const lastActive = ({ matter, sessions }: MatterActivity) => Math.max(epochMs(matter.created_at), ...sessions.map((session) => session.time.updated))

export type RecentMatter = { matter: Matter; last: number; cost: number; documents: number; status?: LeanDraftingStatus["status"] }

// The sample matter is listed like any other: it is somewhere to work.
export function recentMatters(matters: MatterActivity[], limit = 5): RecentMatter[] {
  return matters
    .map((row) => ({
      matter: row.matter,
      last: lastActive(row),
      cost: spent(row.costs),
      documents: draftedCount(row.drafting),
      status: row.drafting?.status,
    }))
    .sort((a, b) => b.last - a.last)
    .slice(0, limit)
}

export type Progress = {
  matters: number
  documents: number
  conversations: number
  drafted: number
  // Where the next step happens: the matter last worked in, preferring one
  // with documents.
  target?: Matter
  // A drafting run still under way or waiting on the lawyer (a stalled run
  // waits for Continue).
  run?: { matter: Matter; session: string; waiting: boolean }
}

// How far the lawyer has come in their own matters, for the first-steps track
// and the starting layout. The sample matter does not count.
export function progress(matters: MatterActivity[]): Progress {
  const own = matters.filter((row) => !row.matter.sample).sort((a, b) => lastActive(b) - lastActive(a))
  const live = own.find(({ drafting }) => drafting && !drafting.stopped && drafting.status !== "completed")
  return {
    matters: own.length,
    documents: own.reduce((sum, row) => sum + (row.documents ?? 0), 0),
    conversations: own.reduce((sum, row) => sum + row.sessions.filter(isConversation).length, 0),
    drafted: own.reduce((sum, row) => sum + draftedCount(row.drafting), 0),
    target: (own.find((row) => row.documents) ?? own[0])?.matter,
    run: live?.drafting && { matter: live.matter, session: live.drafting.owner_session_id, waiting: live.drafting.status !== "running" || !live.drafting.in_flight || live.drafting.questions.some((question) => question.status === "open") },
  }
}
