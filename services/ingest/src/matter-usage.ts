import { Database } from "bun:sqlite"
import { existsSync } from "node:fs"
import path from "node:path"
import { engineV2Fetch } from "../../../dochaus/lib/engine-v2"
import { readGenerations } from "../../../dochaus/lib/telemetry"
import { UPKEEP_AGENT, usageAgentKey } from "../../../dochaus/lib/usage-agents"
import { engineURL } from "./matter-session"

// One usage ledger per matter, priced as the provider reported it. Native
// sessions carry each model step's agent, model, tokens and cost; a session's
// cost beyond its messages is title and compaction upkeep. Stateless generation
// jobs have no session and are read from the matter's telemetry database.
export type UsageTokens = { input: number; output: number; reasoning: number; cache: { read: number; write: number } }
export type UsageEntry = { time: number; cost: number; session?: string; agent: string; model?: string; tokens?: UsageTokens }
export type UsageSession = {
  id: string
  parentID?: string
  title?: string
  agent?: string
  metadata?: Record<string, unknown> & { dochaus_role?: unknown }
  cost: number
  tokens: { input: number; output: number; reasoning: number; cache?: { read: number; write: number } }
  time: { created: number; updated: number }
}
type UsageMessage = { info: { cost?: number; agent?: string; providerID?: string; modelID?: string; tokens?: Partial<UsageTokens>; time: { created: number } } }

export async function enginePages(directory: string, pathname: string, signal: AbortSignal): Promise<unknown[]> {
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

const tokens = (value: Partial<UsageTokens> | undefined): UsageTokens => ({
  input: value?.input || 0, output: value?.output || 0, reasoning: value?.reasoning || 0,
  cache: { read: value?.cache?.read || 0, write: value?.cache?.write || 0 },
})

// Usage for one session. On a message read failure the whole session cost is
// kept under the session's own agent and the caller marks the matter partial.
export async function sessionUsage(directory: string, session: UsageSession, signal: AbortSignal): Promise<{ entries: UsageEntry[]; unread: boolean }> {
  const whole: UsageEntry = { time: session.time.created, cost: session.cost || 0, session: session.id, agent: usageAgentKey(session.agent, session), tokens: tokens(session.tokens) }
  if (!session.cost) return { entries: [whole], unread: false }
  try {
    const messages = await enginePages(directory, `/api/session/${encodeURIComponent(session.id)}/message`, signal) as UsageMessage[]
    const entries = messages.flatMap(({ info }): UsageEntry[] => info.cost ? [{
      time: info.time.created, cost: info.cost, session: session.id,
      agent: usageAgentKey(info.agent ?? session.agent, session),
      ...(info.providerID && info.modelID ? { model: `${info.providerID}/${info.modelID}` } : {}),
      tokens: tokens(info.tokens),
    }] : [])
    const rest = session.cost - entries.reduce((sum, entry) => sum + entry.cost, 0)
    if (rest > .0001) entries.push({ time: session.time.updated, cost: rest, session: session.id, agent: UPKEEP_AGENT })
    return { entries, unread: false }
  } catch {
    return { entries: [whole], unread: true }
  }
}

export function generationUsage(directory: string): UsageEntry[] {
  const file = path.join(directory, ".dochaus", "telemetry.db")
  if (!existsSync(file)) return []
  let db: Database | undefined
  try {
    db = new Database(file, { readonly: true })
    const table = db.query("SELECT name FROM sqlite_master WHERE type = 'table' AND name = 'generations'").get()
    if (!table) return []
    return readGenerations(db).map(row => ({ time: row.ts, cost: row.cost ?? 0, agent: row.agent, ...(row.model ? { model: row.model } : {}), tokens: row.tokens }))
  } catch {
    return []
  } finally {
    db?.close()
  }
}

export async function matterUsage(directory: string, signal: AbortSignal): Promise<{ sessions: UsageSession[]; entries: UsageEntry[]; unread: boolean }> {
  const sessions = await enginePages(directory, "/api/session", signal) as UsageSession[]
  const entries: UsageEntry[] = []
  let unread = false
  for (const session of sessions) {
    const result = await sessionUsage(directory, session, signal)
    entries.push(...result.entries)
    unread ||= result.unread
  }
  entries.push(...generationUsage(directory))
  return { sessions, entries, unread }
}
