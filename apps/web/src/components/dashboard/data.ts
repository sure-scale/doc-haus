import { useEffect, useState } from "react"
import { getMatter, listMatters, type Matter } from "../../api/ingest/matters"
import { getDraftingStatus } from "../../api/ingest/drafting-status"
import { leanStageSessions, type LeanDraftingStatus } from "../../api/ingest/lean"
import { getPreferences, type DraftingPreferences } from "../../api/ingest/preferences"
import { fetchWordSetup, type WordSetupStatus } from "../../api/ingest/word"
import { listPlaybookLibrary, listTemplates } from "../../api/ingest/library"
import { listClients } from "../../api/ingest/clients"
import { listSessions } from "../../api/opencode/sessions"
import { matterClient } from "../../api/opencode/client"

export type Session = Awaited<ReturnType<typeof listSessions>>[number]

// One matter as the dashboard sees it: every engine session in its folder
// (subagent runs included, each carrying its own cost) and its drafting run.
// `documents` is read only for the newest few of the lawyer's own matters.
export type MatterActivity = { matter: Matter; sessions: Session[]; drafting?: LeanDraftingStatus; documents?: number; failed: boolean }

export type DashboardData = {
  matters: MatterActivity[]
  preferences?: DraftingPreferences
  word?: WordSetupStatus
  templates?: number
  playbooks?: number
  clients?: number
  // Set when the matter list itself could not be read, so an unreachable
  // service is never shown as an empty practice.
  error?: string
}

// Days in the chosen window; 0 is all time.
export type Period = 7 | 30 | 90 | 0

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

async function loadMatter(matter: Matter, withDocuments: boolean): Promise<MatterActivity> {
  const [sessions, drafting, documents] = await Promise.allSettled([
    listSessions(matterClient(matter.dir)),
    getDraftingStatus(matter.id),
    withDocuments ? getMatter(matter.id).then((detail) => detail.documents.length) : Promise.resolve(undefined),
  ])
  return {
    matter,
    sessions: sessions.status === "fulfilled" ? sessions.value : [],
    drafting: drafting.status === "fulfilled" ? drafting.value : undefined,
    documents: documents.status === "fulfilled" ? documents.value : undefined,
    failed: sessions.status === "rejected",
  }
}

export async function loadDashboard(): Promise<DashboardData> {
  const [matters, preferences, word, templates, playbooks, clients] = await Promise.all([
    listMatters().then((list) => ({ list }), (error: unknown) => ({ list: [] as Matter[], error: error instanceof Error ? error.message : "Could not load matters" })),
    settled(getPreferences()),
    settled(fetchWordSetup()),
    settled(listTemplates().then((result) => result.templates.length)),
    settled(listPlaybookLibrary().then((result) => result.playbooks.length)),
    settled(listClients().then((result) => result.length)),
  ])
  // First steps only ask whether documents were added, so the file lists of
  // the newest few own matters are enough.
  const newest = new Set(matters.list.filter((matter) => !matter.sample).sort((a, b) => b.created_at - a.created_at).slice(0, 3).map((matter) => matter.id))
  return {
    matters: await mapLimited(matters.list, 4, (matter) => loadMatter(matter, newest.has(matter.id))),
    preferences, word, templates, playbooks, clients,
    ...("error" in matters ? { error: matters.error } : {}),
  }
}

// Re-reads when `setup` changes (App re-reads it whenever Settings or
// onboarding closes, so saved settings may have changed), when a matter is added
// elsewhere (`refresh`) or on reload(). Keeps the previous figures on screen
// while a refresh runs.
export function useDashboardData(setup: unknown, refresh: number) {
  const [data, setData] = useState<DashboardData>()
  const [loading, setLoading] = useState(true)
  const [reloads, setReloads] = useState(0)
  useEffect(() => {
    let live = true
    setLoading(true)
    void loadDashboard().then((next) => {
      if (!live) return
      setData(next)
      setLoading(false)
    })
    return () => { live = false }
  }, [setup, refresh, reloads])
  return { data, loading, reload: () => setReloads((count) => count + 1) }
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

// The window's first instant: the start of the day `period - 1` days before
// today, so "7 days" is today and the six before it.
export function periodStart(period: Period, now: number) {
  return period === 0 ? -Infinity : addDays(startOfDay(now), 1 - period)
}

const sessionsOf = (matters: MatterActivity[]) => matters.flatMap((row) => row.sessions)

// A conversation the lawyer started, as the sidebar lists them: not a subagent
// run, not a session the runtime owns (drafting stages, the narrator) unless it
// is a drafting run's own conversation, and not a system-titled one such as a
// review grid's cell sessions. Their cost still counts towards spend.
export function isConversation(session: Session) {
  if (session.parentID || /^doc\.haus/i.test(session.title ?? "")) return false
  const role = typeof session.metadata === "object" && session.metadata !== null && "dochaus_role" in session.metadata ? session.metadata.dochaus_role : undefined
  return role === undefined || role === "lean-drafting-owner"
}

export function totalCost(sessions: Session[]) {
  return sessions.reduce((sum, session) => sum + (session.cost || 0), 0)
}

// Spend for each of the last `days` days (oldest first), each conversation
// counted on the day it started.
export function dailySpend(matters: MatterActivity[], days: number, now: number) {
  const first = addDays(startOfDay(now), 1 - days)
  const buckets = Array.from({ length: days }, (_, index) => ({ day: addDays(first, index), cost: 0, count: 0 }))
  for (const session of sessionsOf(matters)) {
    const index = dayIndex(first, session.time.created)
    if (index < 0 || index >= days) continue
    buckets[index].cost += session.cost || 0
    if (isConversation(session)) buckets[index].count += 1
  }
  return buckets
}

// Conversations a lawyer started each day — the activity calendar's cells.
export function dailyConversations(matters: MatterActivity[], first: number, days: number) {
  const counts = new Array<number>(days).fill(0)
  for (const session of sessionsOf(matters)) {
    if (!isConversation(session)) continue
    const index = dayIndex(first, session.time.created)
    if (index >= 0 && index < days) counts[index] += 1
  }
  return counts
}

export type MatterCost = { matter: Matter; cost: number; conversations: number }

export function costByMatter(matters: MatterActivity[], period: Period, now: number): MatterCost[] {
  const start = periodStart(period, now)
  return matters
    .map(({ matter, sessions }) => {
      const inWindow = sessions.filter((session) => session.time.created >= start)
      return { matter, cost: totalCost(inWindow), conversations: inWindow.filter(isConversation).length }
    })
    .filter((row) => row.cost > 0)
    .sort((a, b) => b.cost - a.cost)
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

export type DraftedRun = { matter: Matter; cost: number; documents: number; finished: number }

// Documents a completed run drafted: its delivered bodies, as the run's own
// header counts them. Redlines, notes and combined files are copies of those.
export function draftedCount(drafting: LeanDraftingStatus | undefined) {
  if (drafting?.documents?.status !== "completed") return 0
  return drafting.documents.documents.filter((row) => row.kind === "body" && row.deliverable).length
}

// A matter's latest drafting run, once it has delivered. The ingest status
// keeps only the latest run, so earlier runs on the same matter are not seen.
export function draftedRuns(matters: MatterActivity[], period: Period, now: number): DraftedRun[] {
  const start = periodStart(period, now)
  return matters.flatMap(({ matter, sessions, drafting }) => {
    if (!drafting || drafting.status !== "completed") return []
    const documents = draftedCount(drafting)
    const finished = Date.parse(drafting.updated_at)
    if (!documents || !(finished >= start)) return []
    return [{ matter, cost: totalCost(runSessions(drafting, sessions)), documents, finished }]
  })
}

export type Summary = {
  spend: number
  conversations: number
  activeMatters: number
  averageMatter?: number
  documents: number
  averageDocument?: number
  tokens: { input: number; cached: number; output: number; reasoning: number }
}

export function summarize(matters: MatterActivity[], period: Period, now: number): Summary {
  const start = periodStart(period, now)
  const sessions = sessionsOf(matters).filter((session) => session.time.created >= start)
  const perMatter = costByMatter(matters, period, now)
  const runs = draftedRuns(matters, period, now)
  const documents = runs.reduce((sum, run) => sum + run.documents, 0)
  const spend = totalCost(sessions)
  const tokens = { input: 0, cached: 0, output: 0, reasoning: 0 }
  for (const { tokens: used } of sessions) {
    tokens.input += (used.input || 0) + (used.cache?.write || 0)
    tokens.cached += used.cache?.read || 0
    tokens.output += used.output || 0
    tokens.reasoning += used.reasoning || 0
  }
  return {
    spend,
    conversations: sessions.filter(isConversation).length,
    activeMatters: perMatter.length,
    averageMatter: perMatter.length ? perMatter.reduce((sum, row) => sum + row.cost, 0) / perMatter.length : undefined,
    documents,
    averageDocument: documents ? runs.reduce((sum, run) => sum + run.cost, 0) / documents : undefined,
    tokens,
  }
}

// Where the window's charts begin: all time starts on the first day any
// conversation started.
export function windowStart(matters: MatterActivity[], period: Period, now: number) {
  if (period !== 0) return periodStart(period, now)
  return startOfDay(sessionsOf(matters).reduce((first, session) => Math.min(first, session.time.created), now))
}

// A stat tile's small trend: the window split into even buckets.
export function bucketed(values: { time: number; value: number }[], start: number, now: number, count: number) {
  const span = Math.max(DAY, now - start)
  const buckets = new Array<number>(count).fill(0)
  for (const { time, value } of values) {
    if (time < start || time > now) continue
    buckets[Math.min(count - 1, Math.floor(((time - start) / span) * count))] += value
  }
  return buckets
}

// Spend in the window of the same length just before this one, for the
// change badge; all time has none.
export function previousSpend(matters: MatterActivity[], period: Period, now: number) {
  if (period === 0) return undefined
  const end = periodStart(period, now)
  const start = addDays(end, -period)
  return totalCost(sessionsOf(matters).filter((session) => session.time.created >= start && session.time.created < end))
}

export const sessionCosts = (matters: MatterActivity[]) => sessionsOf(matters).map((session) => ({ time: session.time.created, value: session.cost || 0 }))

export type AttentionItem = { matter: Matter; kind: "questions" | "blocked" | "paused" | "failed" | "provider"; count?: number; session: string }

// Drafting runs that wait on the lawyer, most urgent first.
export function attention(matters: MatterActivity[]): AttentionItem[] {
  const rank = { provider: 0, failed: 1, questions: 2, blocked: 3, paused: 4 }
  return matters
    .flatMap(({ matter, drafting }): AttentionItem[] => {
      // A run the lawyer stopped settles as failed but needs nothing more.
      if (!drafting || drafting.stopped) return []
      const session = drafting.owner_session_id
      const open = drafting.questions.filter((question) => question.status === "open").length
      if (drafting.status === "failed") return [{ matter, kind: drafting.provider_stop ? "provider" : "failed", session }]
      if (open) return [{ matter, kind: "questions", count: open, session }]
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

export function recentMatters(matters: MatterActivity[], limit = 5): RecentMatter[] {
  return matters
    .map((row) => ({
      matter: row.matter,
      last: lastActive(row),
      cost: totalCost(row.sessions),
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
  // A drafting run still under way or waiting on the lawyer.
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
    run: live?.drafting && { matter: live.matter, session: live.drafting.owner_session_id, waiting: live.drafting.status !== "running" || live.drafting.questions.some((question) => question.status === "open") },
  }
}
