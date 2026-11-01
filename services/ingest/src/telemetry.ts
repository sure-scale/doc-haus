import { Database } from "bun:sqlite"
import { existsSync, statSync } from "node:fs"
import path from "node:path"
import { UPKEEP_AGENT, USAGE_GROUP_LABEL, humanize, usageAgent, usageAgentKey, type UsageGroup } from "../../../dochaus/lib/usage-agents"
import { enginePages, generationUsage, sessionUsage, type UsageEntry, type UsageSession } from "./matter-usage"

// The Usage drawer reports provider cost as OpenCode recorded it per model
// step (see matter-usage.ts), so its totals match the Dashboard. telemetry.db
// supplies what live native sessions do not: stateless job and deleted session
// usage, tool calls and the latest context size check.

export type LeanPayload = {
  current_activity: { cost_usd: number; turns: number; since_ts: number } | null
  today: { cost_usd: number; turns: number }
  matter_total: { cost_usd: number; turns: number; sessions: number; since_ts: number }
  current_session: { tokens: number; status: "ok" | "warn" | "fail"; agent: string; label: string; model: string } | null
  // Model calls with tokens but no provider-reported cost.
  unpriced_turns: number
  // Some conversations could not be read; totals may be low.
  partial: boolean
}

export type DetailPayload = {
  by_group: Array<{ group: UsageGroup; label: string; cost_usd: number; turns: number }>
  by_agent: Array<{ agent: string; label: string; group: UsageGroup; cost_usd: number; turns: number; tool_calls: number }>
  // Upkeep has no model; its row carries a plain label instead, so the rows
  // sum to the matter total.
  by_model: Array<{ model: string | null; label?: string; cost_usd: number; turns: number; unpriced_turns: number }>
  by_tool: Array<{ tool: string; label: string; calls: number; total_duration_ms: number; avg_duration_ms: number }>
  recent_turns: Array<{ ts: number; agent: string; label: string; group: UsageGroup; model: string | null; tokens: number; cost_usd: number | null }>
  recent_sessions: Array<{ id: string; title: string; agent: string; label: string; group: UsageGroup; last_activity_at: number; cost_usd: number }>
  sparkline: Array<{ ts: number; cost_usd: number }>
}

export type StoragePayload = {
  size_bytes: number
  matter_count: number
  oldest_ts: number | null
}

type Usage = { sessions: UsageSession[]; entries: UsageEntry[]; unread: boolean }
type CachedSession = { updated: number; cost: number; entries: UsageEntry[] }

// The status-bar Usage button polls every 5 s. Listing sessions is one cheap
// call; reading a session's messages is not, so each matter keeps the last
// read per session and re-reads only sessions whose update time or cost moved.
const usageCache = new Map<string, Map<string, CachedSession>>()

export async function readMatterUsage(matterDir: string, signal: AbortSignal): Promise<Usage> {
  let sessions: UsageSession[]
  try {
    sessions = await enginePages(matterDir, "/api/session", signal) as UsageSession[]
  } catch {
    return { sessions: [], entries: generationUsage(matterDir), unread: true }
  }
  const previous = usageCache.get(matterDir)
  const next = new Map<string, CachedSession>()
  const entries: UsageEntry[] = []
  let unread = false
  for (const session of sessions) {
    const hit = previous?.get(session.id)
    if (hit && hit.updated === session.time.updated && hit.cost === (session.cost || 0)) {
      next.set(session.id, hit)
      entries.push(...hit.entries)
      continue
    }
    const result = await sessionUsage(matterDir, session, signal)
    entries.push(...result.entries)
    if (result.unread) unread = true
    else next.set(session.id, { updated: session.time.updated, cost: session.cost || 0, entries: result.entries })
  }
  usageCache.set(matterDir, next)
  entries.push(...generationUsage(matterDir, sessions))
  return { sessions, entries: entries.filter(used), unread }
}

export function clearUsageCache() {
  usageCache.clear()
}

export async function readLeanPayload(matterDir: string, opts: { signal: AbortSignal; activityWindowMs?: number; session?: string }): Promise<LeanPayload> {
  const usage = await readMatterUsage(matterDir, opts.signal)
  const now = Date.now()
  const activityWindowMs = opts.activityWindowMs ?? 30 * 60_000
  const current = summarize(usage.entries.filter((entry) => entry.time >= now - activityWindowMs))
  const today = summarize(usage.entries.filter((entry) => entry.time >= new Date().setHours(0, 0, 0, 0)))
  const total = summarize(usage.entries)
  return {
    current_activity: current.turns > 0 || current.cost_usd > 0 ? { ...current, since_ts: now - activityWindowMs } : null,
    today,
    matter_total: {
      ...total,
      sessions: usage.sessions.length,
      since_ts: usage.entries.length ? Math.min(...usage.entries.map((entry) => entry.time)) : 0,
    },
    current_session: readCurrentSession(matterDir, usage.sessions, opts.session),
    unpriced_turns: usage.entries.filter(unpriced).length,
    partial: usage.unread,
  }
}

export async function readDetailPayload(matterDir: string, signal: AbortSignal): Promise<DetailPayload> {
  const usage = await readMatterUsage(matterDir, signal)
  const sessions = new Map(usage.sessions.map((session) => [session.id, session]))
  const { byTool, toolCallsByAgent } = readToolCalls(matterDir, sessions)

  const agents = new Map<string, DetailPayload["by_agent"][number]>()
  const agentRow = (key: string) => {
    const existing = agents.get(key)
    if (existing) return existing
    const { label, group } = usageAgent(key)
    const row = { agent: key, label, group, cost_usd: 0, turns: 0, tool_calls: 0 }
    agents.set(key, row)
    return row
  }
  const models = new Map<string | null, DetailPayload["by_model"][number]>()
  let upkeep: DetailPayload["by_model"][number] | undefined
  for (const entry of usage.entries) {
    const agent = agentRow(entry.agent)
    agent.cost_usd += entry.cost
    agent.turns += call(entry) ? 1 : 0
    if (entry.agent === UPKEEP_AGENT) {
      upkeep ??= { model: null, label: usageAgent(UPKEEP_AGENT).label, cost_usd: 0, turns: 0, unpriced_turns: 0 }
      upkeep.cost_usd += entry.cost
      continue
    }
    const model = entry.model ?? null
    const row = models.get(model) ?? { model, cost_usd: 0, turns: 0, unpriced_turns: 0 }
    row.cost_usd += entry.cost
    row.turns += 1
    row.unpriced_turns += unpriced(entry) ? 1 : 0
    models.set(model, row)
  }
  for (const [key, count] of toolCallsByAgent) agentRow(key).tool_calls += count
  const byAgent = [...agents.values()].sort((a, b) => b.cost_usd - a.cost_usd || b.turns - a.turns)

  const byGroup = (["background", "work"] as const).map((group) => {
    const rows = byAgent.filter((row) => row.group === group)
    return {
      group,
      label: USAGE_GROUP_LABEL[group],
      cost_usd: rows.reduce((sum, row) => sum + row.cost_usd, 0),
      turns: rows.reduce((sum, row) => sum + row.turns, 0),
    }
  })

  // A conversation's cost includes the sub-agents it started.
  const root = (id: string) => {
    let current = sessions.get(id)
    for (let depth = 0; current?.parentID && sessions.has(current.parentID) && depth < 50; depth++) current = sessions.get(current.parentID)
    return current?.id ?? id
  }
  const sessionCost = new Map<string, number>()
  for (const entry of usage.entries) {
    if (!entry.session) continue
    const id = root(entry.session)
    sessionCost.set(id, (sessionCost.get(id) ?? 0) + entry.cost)
  }
  const recentSessions = usage.sessions
    .filter((session) => !session.parentID || !sessions.has(session.parentID))
    .sort((a, b) => b.time.updated - a.time.updated)
    .slice(0, 20)
    .map((session) => {
      const agent = usageAgent(session.agent, session)
      return {
        id: session.id,
        title: session.title ?? "",
        agent: agent.key,
        label: agent.label,
        group: agent.group,
        last_activity_at: session.time.updated,
        cost_usd: sessionCost.get(session.id) ?? 0,
      }
    })

  const since = Date.now() - 24 * 60 * 60_000
  const hours = new Map<number, number>()
  for (const entry of usage.entries) {
    if (entry.time < since) continue
    const hour = Math.trunc(entry.time / 3_600_000) * 3_600_000
    hours.set(hour, (hours.get(hour) ?? 0) + entry.cost)
  }

  const recentTurns = usage.entries
    .filter(call)
    .sort((a, b) => b.time - a.time)
    .slice(0, 50)
    .map((entry) => {
      const agent = usageAgent(entry.agent)
      return {
        ts: entry.time,
        agent: agent.key,
        label: agent.label,
        group: agent.group,
        model: entry.model ?? null,
        tokens: tokenTotal(entry),
        cost_usd: unpriced(entry) ? null : entry.cost,
      }
    })

  return {
    by_group: byGroup,
    by_agent: byAgent,
    by_model: [...models.values(), ...(upkeep ? [upkeep] : [])].sort((a, b) => b.cost_usd - a.cost_usd || b.turns - a.turns),
    by_tool: byTool,
    recent_turns: recentTurns,
    recent_sessions: recentSessions,
    sparkline: [...hours.entries()].map(([ts, cost_usd]) => ({ ts, cost_usd })).sort((a, b) => a.ts - b.ts),
  }
}

// Storage covers what Clear removes: tool, context and stateless generation
// records. Conversation cost lives with conversation history.
export function readStoragePayload(matterDirs: string[]): StoragePayload {
  let size = 0
  let oldest: number | null = null
  for (const dir of matterDirs) {
    const files = telemetryDbFiles(dir)
    if (files.length === 0) continue
    size += files.reduce((sum, file) => sum + statSync(file).size, 0)
    const ts = withTelemetryDb(dir, (db) => {
      const tables = new Set((db.query("SELECT name FROM sqlite_master WHERE type = 'table'").all() as { name: string }[]).map((row) => row.name))
      const times = ["turns", "tool_calls", "generations"]
        .filter((table) => tables.has(table))
        .map((table) => (db.query(`SELECT MIN(ts) AS ts FROM ${table}`).get() as { ts: number | null }).ts)
        .filter((value): value is number => typeof value === "number")
      return times.length ? Math.min(...times) : null
    })
    if (typeof ts === "number") oldest = oldest === null ? ts : Math.min(oldest, ts)
  }
  return { size_bytes: size, matter_count: matterDirs.length, oldest_ts: oldest }
}

function readToolCalls(matterDir: string, sessions: Map<string, UsageSession>) {
  const result = withTelemetryDb(matterDir, (db) => {
    if (!hasTable(db, "tool_calls")) return undefined
    const byTool = (db
      .query(`
        SELECT
          tool,
          COUNT(*) AS calls,
          COALESCE(SUM(duration_ms), 0) AS total_duration_ms,
          COALESCE(AVG(duration_ms), 0) AS avg_duration_ms
        FROM tool_calls
        GROUP BY tool
        ORDER BY calls DESC
      `)
      .all() as Omit<DetailPayload["by_tool"][number], "label">[])
      .map((row) => ({ ...row, label: humanize(row.tool) }))
    const toolCallsByAgent = new Map<string, number>()
    const rows = db
      .query("SELECT session_id, agent, COUNT(*) AS calls FROM tool_calls GROUP BY session_id, agent")
      .all() as { session_id: string; agent: string | null; calls: number }[]
    for (const row of rows) {
      const key = usageAgentKey(row.agent ?? undefined, sessions.get(row.session_id))
      toolCallsByAgent.set(key, (toolCallsByAgent.get(key) ?? 0) + row.calls)
    }
    return { byTool, toolCallsByAgent }
  })
  return result ?? { byTool: [], toolCallsByAgent: new Map<string, number>() }
}

// Context size is the plugin's per-turn estimate; native sessions record no
// equivalent. The assistant name comes from the native session when present.
// Without a requested session this is the lawyer's latest conversation, so
// background agents, sub-agents and runtime-owned jobs are passed over.
function readCurrentSession(matterDir: string, sessions: UsageSession[], sessionID?: string): LeanPayload["current_session"] {
  return withTelemetryDb(matterDir, (db) => {
    if (!hasTable(db, "sessions") || !hasTable(db, "turns")) return null
    const rows = (
      sessionID
        ? db.query("SELECT id, parent_id, model, agent FROM sessions WHERE id = ?").all(sessionID)
        : db.query("SELECT id, parent_id, model, agent FROM sessions ORDER BY last_activity_at DESC LIMIT 100").all()
    ) as { id: string; parent_id: string | null; model: string | null; agent: string | null }[]
    const natives = new Map(sessions.map((item) => [item.id, item]))
    const turnQuery = db.query("SELECT agent, model, estimated_input_tokens AS tokens FROM turns WHERE session_id = ? ORDER BY ts DESC LIMIT 1")
    for (const session of rows) {
      const turn = turnQuery.get(session.id) as { agent: string | null; model: string | null; tokens: number } | null
      const native = natives.get(session.id)
      const agent = usageAgent(session.agent ?? turn?.agent ?? native?.agent, native)
      if (!sessionID && !conversation(session.parent_id, native, agent.group)) continue
      const tokens = turn?.tokens ?? 0
      return {
        tokens,
        status: tokenStatus(tokens),
        agent: agent.key,
        label: agent.label,
        model: session.model ?? turn?.model ?? "unknown",
      }
    }
    return null
  }) ?? null
}

// The conversation list's rule: a top-level chat the lawyer started, or a
// drafting run's own conversation; never a background agent or a throwaway
// job session (review table cells, routing), which carry a "doc.haus" title.
function conversation(parentID: string | null, native: UsageSession | undefined, group: UsageGroup) {
  if (group === "background" || parentID || native?.parentID || /^doc\.haus/i.test(native?.title ?? "")) return false
  const role = native?.metadata?.dochaus_role
  return role === undefined || role === "lean-drafting-owner"
}

function tokenStatus(tokens: number): "ok" | "warn" | "fail" {
  if (tokens >= 180_000) return "fail"
  if (tokens >= 90_000) return "warn"
  return "ok"
}

function summarize(entries: UsageEntry[]) {
  return {
    cost_usd: entries.reduce((sum, entry) => sum + entry.cost, 0),
    turns: entries.filter(call).length,
  }
}

function tokenTotal(entry: UsageEntry) {
  const tokens = entry.tokens
  return tokens ? tokens.input + tokens.output + tokens.reasoning + tokens.cache.read + tokens.cache.write : 0
}

// An entry with neither cost nor tokens is an empty conversation.
function used(entry: UsageEntry) {
  return entry.cost > 0 || tokenTotal(entry) > 0
}

// One model call; title and compaction upkeep is cost, not a call.
function call(entry: UsageEntry) {
  return entry.agent !== UPKEEP_AGENT
}

function unpriced(entry: UsageEntry) {
  return entry.cost <= 0 && tokenTotal(entry) > 0
}

function withTelemetryDb<T>(matterDir: string, read: (db: Database) => T): T | undefined {
  const file = telemetryDbPath(matterDir)
  if (!existsSync(file)) return undefined
  let db: Database | undefined
  try {
    db = new Database(file, { readonly: true })
    return read(db)
  } catch {
    return undefined
  } finally {
    db?.close()
  }
}

function hasTable(db: Database, name: string) {
  return Boolean(db.query("SELECT name FROM sqlite_master WHERE type = 'table' AND name = ?").get(name))
}

function telemetryDbPath(matterDir: string) {
  return path.join(matterDir, ".dochaus", "telemetry.db")
}

function telemetryDbFiles(matterDir: string) {
  const file = telemetryDbPath(matterDir)
  return [file, `${file}-wal`, `${file}-shm`].filter((sibling) => existsSync(sibling))
}
