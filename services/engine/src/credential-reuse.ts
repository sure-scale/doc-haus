import { Database } from "bun:sqlite"
import { existsSync, readFileSync, realpathSync } from "node:fs"
import os from "node:os"
import path from "node:path"
import { Credential } from "@opencode/core/credential"
import { Integration } from "@opencode/schema/integration"
import { Effect, Logger, Schema } from "effect"
import { isRemovedProvider } from "./effective-settings"

export type ReusableConnection = { id: string; integrationID: string; label: string; authType: "key" | "oauth" }
type Stored = { integrationID: Integration.ID; label: string; value: Credential.Value }
type Row = { id: string; integration_id: string; label: string; authType: string }
const decode = Schema.decodeUnknownSync(Credential.Value)
const integration = (value: unknown): value is string => typeof value === "string" && /^[a-zA-Z0-9][a-zA-Z0-9._-]{0,127}$/.test(value)

// The launcher captures the native XDG root before sealing the engine runtime.
// Direct engine starts retain the native default; no request can supply a path.
export function nativeCredentialDirectory() {
  return process.env.DOCHAUS_NATIVE_OPENCODE_DATA_DIR ?? path.join(process.env.XDG_DATA_HOME || path.join(os.homedir(), ".local", "share"), "opencode")
}

export function credentialReuse(options: {
  directory: string
  targetDirectory: string
  create: (input: Stored) => Promise<{ id: string }>
}) {
  const sources = [
    { id: crypto.randomUUID(), kind: "native", label: "OpenCode connections", file: path.join(options.directory, "opencode.db") },
    { id: crypto.randomUUID(), kind: "native", label: "OpenCode local connections", file: path.join(options.directory, "opencode-local.db") },
    { id: crypto.randomUUID(), kind: "legacy", label: "OpenCode legacy sign-ins", file: path.join(options.directory, "auth.json") },
  ] as const
  const handles = new Map<string, { sourceId: string; record: string; connection: ReusableConnection }>()
  function handle(sourceId: string, record: string, info: Omit<ReusableConnection, "id">) {
    const prior = Array.from(handles.entries()).find(([, value]) => value.sourceId === sourceId && value.record === record)
    const id = prior?.[0] ?? crypto.randomUUID()
    const connection = { id, ...info }
    handles.set(id, { sourceId, record, connection })
    return connection
  }
  function available(file: string) {
    if (!existsSync(file)) return false
    // Never offer the active doc.haus store as an external connection source.
    return !["opencode.db", "opencode-local.db", "auth.json"].some(name => {
      const target = path.join(options.targetDirectory, name)
      return existsSync(target) && realpathSync(file) === realpathSync(target)
    })
  }
  function legacy(file: string): Record<string, unknown> {
    const value: unknown = JSON.parse(readFileSync(file, "utf8"))
    if (!value || typeof value !== "object" || Array.isArray(value)) throw new Error("Invalid credential source")
    return value as Record<string, unknown>
  }
  return {
    list() {
      try {
        return { sources: sources.flatMap(source => {
          if (!available(source.file)) return []
          const connections: ReusableConnection[] = []
          let unavailableCount = 0
          try {
          if (source.kind === "native") {
            // Opening the native service would run migrations. Read only the
            // current credential table, never sessions, config or discovery KV.
            const db = new Database(source.file, { readonly: true })
            try {
              const columns = db.query<{ name: string }, []>("PRAGMA table_info(credential)").all().map(row => row.name)
              if (!["id", "integration_id", "label", "value"].every(column => columns.includes(column))) throw new Error("Unsupported source schema")
              const rows = db.query<Row, []>("SELECT id, integration_id, label, json_extract(value, '$.type') AS authType FROM credential").all()
              for (const row of rows) {
                if (isRemovedProvider(row.integration_id)) continue
                if (!integration(row.integration_id) || !["key", "oauth"].includes(row.authType)) { unavailableCount++; continue }
                connections.push(handle(source.id, row.id, { integrationID: row.integration_id, label: row.label, authType: row.authType as "key" | "oauth" }))
              }
            } finally { db.close() }
          } else {
            for (const [id, raw] of Object.entries(legacy(source.file))) {
              const integrationID = id.replace(/\/+$/, "")
              if (isRemovedProvider(integrationID)) continue
              if (!integration(integrationID) || !raw || typeof raw !== "object" || !("type" in raw) || raw.type !== "api" && raw.type !== "oauth") { unavailableCount++; continue }
              connections.push(handle(source.id, id, { integrationID, label: raw.type === "oauth" ? "OAuth" : "API key", authType: raw.type === "oauth" ? "oauth" : "key" }))
            }
          }
          return [{ id: source.id, label: source.label, status: "ready", unavailableCount, connections }]
          } catch { return [{ id: source.id, label: source.label, status: "unavailable", unavailableCount, connections: [], message: "This OpenCode credential store could not be read or uses an unsupported format." }] }
        }) }
      } catch { throw new Error("Could not read existing OpenCode sign-ins. The original store has not been changed.") }
    },
    async importSelected(input: unknown) {
      if (!input || typeof input !== "object" || Array.isArray(input) || Object.keys(input).sort().join(",") !== "connectionId,sourceId" ||
          !("connectionId" in input) || !("sourceId" in input) || typeof input.connectionId !== "string" || typeof input.sourceId !== "string")
        throw new Error("Select one existing OpenCode connection to reuse.")
      const selected = handles.get(input.connectionId)
      const source = sources.find(source => source.id === input.sourceId)
      if (!selected || !source || selected.sourceId !== source.id) throw new Error("Refresh the sign-in list and select one connection.")
      try {
        if (!available(source.file)) throw new Error("Source unavailable")
        let stored: Stored
        if (source.kind === "native") {
          const db = new Database(source.file, { readonly: true })
          try {
            const row = db.query<{ integration_id: string; label: string; value: string }, [string]>("SELECT integration_id, label, value FROM credential WHERE id = ?").get(selected.record)
            if (!row || !integration(row.integration_id)) throw new Error("Connection unavailable")
            stored = { integrationID: Integration.ID.make(row.integration_id), label: row.label, value: decode(JSON.parse(row.value)) }
          } finally { db.close() }
        } else {
          stored = legacyCredential(selected.connection.integrationID, legacy(source.file)[selected.record])
        }
        // Revalidate the current source row as well as the discovery handle:
        // another process may have changed its integration since listing.
        if (isRemovedProvider(stored.integrationID)) throw new Error("Provider unavailable")
        const result = await options.create(stored)
        return { connection: { id: result.id, integrationID: stored.integrationID, label: stored.label, authType: stored.value.type } }
      } catch { throw new Error("Could not reuse the selected sign-in. The original store has not been changed; connect again or refresh the list.") }
    },
  }
}

function legacyCredential(integrationID: string, input: unknown): Stored {
  if (!input || typeof input !== "object" || !("type" in input)) throw new Error("Invalid legacy credential")
  const value = input as Record<string, unknown>
  const converted = value.type === "api" ? { type: "key", key: value.key, metadata: value.metadata }
    : value.type === "oauth" ? {
      type: "oauth", refresh: value.refresh, access: value.access, expires: value.expires,
      methodID: integrationID === "openai" ? "chatgpt-browser" : ["github-copilot", "opencode", "xai"].includes(integrationID) ? "device" : "oauth",
      metadata: value.accountId || value.enterpriseUrl ? {
        ...(typeof value.accountId === "string" ? { accountID: value.accountId } : {}),
        ...(typeof value.enterpriseUrl === "string" ? { enterpriseUrl: value.enterpriseUrl } : {}),
      } : undefined,
    } : undefined
  return { integrationID: Integration.ID.make(integrationID), label: value.type === "oauth" ? "OAuth" : "API key", value: decode(converted) }
}

export function credentialReuseHandler(service: ReturnType<typeof credentialReuse>, password?: string) {
  return async (request: Request) => {
    const origin = request.headers.get("origin")
    const headers = { "Cache-Control": "no-store", ...(origin && /^https?:\/\/(localhost|127\.0\.0\.1)(:\d+)?$/.test(origin)
      ? { "Access-Control-Allow-Origin": origin, "Access-Control-Allow-Credentials": "true", Vary: "Origin" } : {}) }
    const json = (body: object, status = 200) => Response.json(body, { status, headers })
    if (origin && !/^https?:\/\/(localhost|127\.0\.0\.1)(:\d+)?$/.test(origin)) return json({ error: "Origin is not allowed." }, 403)
    if (request.method === "OPTIONS") return new Response(null, { status: 204, headers: {
      ...headers, "Access-Control-Allow-Methods": "GET, POST", "Access-Control-Allow-Headers": "authorization, content-type",
    } })
    if (password && request.headers.get("authorization") !== `Basic ${btoa(`opencode:${password}`)}`) return json({ error: "Authentication required." }, 401)
    try {
      if (request.method === "GET") return json(service.list())
      if (request.method !== "POST") return json({ error: "Method not allowed." }, 405)
      if (request.headers.get("content-type")?.split(";")[0].trim().toLowerCase() !== "application/json") return json({ error: "Select one connection using JSON." }, 415)
      const text = await readSelectionBody(request)
      return json(await service.importSelected(text === undefined ? undefined : JSON.parse(text)))
    } catch { return json({ error: "Could not read or reuse the selected sign-in. Refresh the list or connect again." }, 400) }
  }
}

async function readSelectionBody(request: Request) {
  const reader = request.body?.getReader()
  if (!reader) return
  const timer = setTimeout(() => { void reader.cancel().catch(() => {}) }, 5000)
  const chunks: Uint8Array[] = []
  let length = 0
  try {
    while (true) {
      const chunk = await reader.read()
      if (chunk.done) return new TextDecoder().decode(Buffer.concat(chunks))
      length += chunk.value.length
      if (length > 1024) { await reader.cancel(); return }
      chunks.push(chunk.value)
    }
  } finally { clearTimeout(timer); reader.releaseLock() }
}

// Drizzle errors can include bound insert values. Suppress native write logging
// inside this one operation; callers return a fixed error without the cause.
export function createReusedCredential(credentials: Credential.Interface, input: Stored) {
  return credentials.create(input).pipe(Effect.provide(Logger.layer([])))
}
