import { Database } from "bun:sqlite"
import { existsSync } from "node:fs"
import path from "node:path"
import { engineV2Fetch } from "../../../dochaus/lib/engine-v2"
import { readGenerations } from "../../../dochaus/lib/telemetry"
import { sessionUsageEntries, wholeSessionUsage, type UsageEntry, type UsageMessage, type UsageSession } from "../../../dochaus/lib/session-usage"
import { engineURL } from "./matter-session"

// One usage ledger per matter, priced as the provider reported it: live native
// sessions (dochaus/lib/session-usage.ts) plus the matter's telemetry database,
// which keeps stateless generation jobs and sessions deleted after they ran.
export type { UsageEntry, UsageSession, UsageTokens } from "../../../dochaus/lib/session-usage"

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

// Usage for one session. On a message read failure the whole session cost is
// kept under the session's own agent and the caller marks the matter partial.
export async function sessionUsage(directory: string, session: UsageSession, signal: AbortSignal): Promise<{ entries: UsageEntry[]; unread: boolean }> {
  if (!session.cost) return { entries: [wholeSessionUsage(session)], unread: false }
  try {
    const messages = await enginePages(directory, `/api/session/${encodeURIComponent(session.id)}/message`, signal) as UsageMessage[]
    return { entries: sessionUsageEntries(session, messages), unread: false }
  } catch {
    return { entries: [wholeSessionUsage(session)], unread: true }
  }
}

// Stored rows of a session that is still live are skipped: the live read
// already counts it (a delete that failed after the engine kept its usage).
export function generationUsage(directory: string, live: readonly { id: string }[] = []): UsageEntry[] {
  const file = path.join(directory, ".dochaus", "telemetry.db")
  if (!existsSync(file)) return []
  let db: Database | undefined
  try {
    db = new Database(file, { readonly: true })
    const table = db.query("SELECT name FROM sqlite_master WHERE type = 'table' AND name = 'generations'").get()
    if (!table) return []
    const ids = new Set(live.map(session => session.id))
    return readGenerations(db).filter(row => !row.session || !ids.has(row.session)).map(row => ({
      time: row.ts, cost: row.cost ?? 0, ...(row.session ? { session: row.session } : {}), agent: row.agent, ...(row.model ? { model: row.model } : {}), tokens: row.tokens,
    }))
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
  entries.push(...generationUsage(directory, sessions))
  return { sessions, entries, unread }
}
