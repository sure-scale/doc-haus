import path from "node:path"
import { realpathSync } from "node:fs"
import { diagnosticErrorMetadata } from "../../../dochaus/lib/diagnostic-metadata"
import { engineV2Fetch } from "../../../dochaus/lib/engine-v2"
import { sessionUsageEntries, usageTokens, wholeSessionUsage, type UsageEntry, type UsageMessage, type UsageSession } from "../../../dochaus/lib/session-usage"
import { openTelemetryDb, recordGeneration } from "../../../dochaus/lib/telemetry"

// Deleting a native session drops its cost, and the core delete also removes its
// child sessions. Before a matter's session DELETE reaches the core handler, the
// session tree's usage is read through the same routes and entries ingest uses
// and kept in the matter's telemetry database under each session id. A failed
// read or write never blocks the delete.
export async function keepDeletedSessionUsage(request: Request, handler: (request: Request) => Promise<Response>, workspace: string, password?: string) {
  const directory = request.headers.get("x-opencode-directory")
  const id = /^\/api\/session\/([^/]+)$/.exec(new URL(request.url).pathname)?.[1]
  if (!directory || !id) return
  if (password && request.headers.get("authorization") !== `Basic ${btoa(`opencode:${password}`)}`) return
  try {
    const actual = realpathSync(directory), root = realpathSync(workspace)
    if (actual === root || !actual.startsWith(root + path.sep)) return
    const transport = Object.assign((input: Parameters<typeof fetch>[0], init?: RequestInit) =>
      handler(new Request(input instanceof Request ? input : String(input), init)), { preconnect: fetch.preconnect }) as typeof fetch
    const engine = engineV2Fetch(transport)
    const read = async (pathname: string, search?: Record<string, string>) => {
      const url = new URL(pathname, request.url)
      for (const [key, value] of Object.entries(search ?? {})) url.searchParams.set(key, value)
      // engineV2Fetch signs internal reads with the server password checked above.
      const response = await engine(url, { headers: { "x-opencode-directory": directory }, signal: request.signal })
      if (!response.ok) throw new Error(`Engine read failed (${response.status})`)
      return await response.json() as unknown
    }
    const ids: string[] = [], entries: UsageEntry[] = []
    const keep = async (session: UsageSession) => {
      ids.push(session.id)
      const messages = session.cost ? await read(`/api/session/${encodeURIComponent(session.id)}/message`).catch(() => undefined) as UsageMessage[] | undefined : []
      entries.push(...messages ? sessionUsageEntries(session, messages) : [wholeSessionUsage(session)])
      for (const child of await read("/api/session", { parentID: session.id }) as UsageSession[]) await keep(child)
    }
    await keep(await read(`/api/session/${id}`) as UsageSession)
    const used = entries.filter(entry => {
      const tokens = usageTokens(entry.tokens)
      return entry.cost > 0 || tokens.input + tokens.output + tokens.reasoning + tokens.cache.read + tokens.cache.write > 0
    })
    if (!used.length) return
    const db = openTelemetryDb(actual)
    try {
      // A retried delete replaces the usage its earlier attempt kept.
      db.transaction(() => {
        for (const session of ids) db.query("DELETE FROM generations WHERE session_id = ?").run(session)
        for (const entry of used) recordGeneration(db, {
          agent: entry.agent, ...(entry.model ? { model: entry.model } : {}), tokens: usageTokens(entry.tokens), cost: entry.cost, ts: entry.time, session: entry.session,
        })
      })()
    } finally { db.close() }
  } catch (error) {
    console.warn("[session-usage] usage was not kept before session deletion", diagnosticErrorMetadata(error))
  }
}
