import { engineV2Fetch } from "../../../dochaus/lib/engine-v2"
import { readSessionExecutionState } from "./chat-continuation"
import { boundedEngineRequest, engineURL, serializeMatter, type EngineRequestOptions } from "./matter-session"

export class TemplateDraftSessionDeletionError extends Error {
  constructor(message: string, readonly status: 409 | 503) { super(message) }
}

type DraftSession = { id: string; directory: string; parentID?: string }
const SESSION_ID = /^[A-Za-z0-9_-]{1,512}$/

// The native delete cascades to children and interrupts active execution. Check
// every descendant and its execution state first: a directory header alone does
// not restrict the native globally-addressed session ID.
export function deleteTemplateDraftSessions(directory: string, options: EngineRequestOptions = {}) {
  return serializeMatter(directory, async () => {
    const request = engineV2Fetch(options.fetch)
    const api = async (route: string, method = "GET") => {
      const response = await boundedEngineRequest(request, `${engineURL(options.engineURL)}${route}`, {
        method, headers: { "x-opencode-directory": directory },
      }).catch(() => undefined)
      if (!response || !response.ok) throw unavailable()
      if (response.status === 204) return undefined
      try { return await response.json() as unknown } catch { throw unavailable() }
    }
    const listed = await api("/api/session")
    const sessions = sessionList(listed, directory)
    const known = new Map(sessions.map(session => [session.id, session]))
    if (known.size !== sessions.length) throw unavailable()

    // Child listing deliberately omits a directory filter. Native deletion is
    // recursive, so a moved child in another location must block deletion too.
    for (const session of sessions) {
      requireSession(await api(`/api/session/${encodeURIComponent(session.id)}`), directory, session.id)
      const children = await nativeChildren(directory, session.id, options)
      for (const child of children) {
        requireSession(child, directory, child.id)
        if (child.parentID !== session.id || !known.has(child.id)) throw changed()
      }
      const expected = sessions.filter(child => child.parentID === session.id)
      if (children.length !== expected.length || expected.some(child => !children.some(actual => actual.id === child.id))) throw changed()
    }
    for (const session of sessions) {
      const state = await readSessionExecutionState(directory, session.id, options)
      if (state === "busy" || state === "retry")
        throw new TemplateDraftSessionDeletionError("Stop this template conversation before deleting it, then try again.", 409)
      if (state !== "idle") throw unavailable()
    }
    const roots = sessions.filter(session => !session.parentID || !known.has(session.parentID))
    if (sessions.length && !roots.length) throw changed()
    for (const session of roots) {
      // Recheck the exact owner immediately before the destructive native call.
      requireSession(await api(`/api/session/${encodeURIComponent(session.id)}`), directory, session.id)
      await api(`/api/session/${encodeURIComponent(session.id)}`, "DELETE")
    }
    if (sessionList(await api("/api/session"), directory).length) throw changed()
  })
}

async function nativeChildren(directory: string, parentID: string, options: EngineRequestOptions): Promise<DraftSession[]> {
  const request = options.fetch ?? fetch
  const headers = new Headers({ "x-opencode-directory": directory })
  const password = process.env.OPENCODE_SERVER_PASSWORD
  if (password) headers.set("Authorization", `Basic ${Buffer.from(`opencode:${password}`).toString("base64")}`)
  const url = new URL(`${engineURL(options.engineURL)}/api/session`)
  url.searchParams.set("parentID", parentID)
  const children: DraftSession[] = []
  const seen = new Set<string>()
  for (;;) {
    const response = await boundedEngineRequest(request, url.href, { headers }).catch(() => undefined)
    if (!response?.ok) throw unavailable()
    const envelope: unknown = await response.json().catch(() => undefined)
    if (!record(envelope) || !Array.isArray(envelope.data)) throw unavailable()
    children.push(...envelope.data.map(value => nativeSession(value)))
    const next = record(envelope.cursor) ? envelope.cursor.next : undefined
    if (next === undefined || next === null) return children
    if (typeof next !== "string" || seen.has(next)) throw unavailable()
    seen.add(next)
    url.searchParams.set("cursor", next)
  }
}

function sessionList(value: unknown, directory: string): DraftSession[] {
  if (!Array.isArray(value)) throw unavailable()
  return value.map(value => requireSession(value, directory))
}

function nativeSession(value: unknown): DraftSession {
  if (!record(value) || !record(value.location)) throw unavailable()
  return { ...value, directory: value.location.directory } as DraftSession
}

function requireSession(value: unknown, directory: string, id?: string): DraftSession {
  if (!record(value) || typeof value.id !== "string" || !SESSION_ID.test(value.id) ||
    typeof value.directory !== "string" || (value.parentID !== undefined && typeof value.parentID !== "string")) throw unavailable()
  if (value.directory !== directory || (id !== undefined && value.id !== id)) throw changed()
  return value as DraftSession
}

function record(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value)
}

function changed() {
  return new TemplateDraftSessionDeletionError("The template conversation changed location. Reload Templates before deleting it.", 409)
}

function unavailable() {
  return new TemplateDraftSessionDeletionError("The conversation service could not confirm deletion. The template draft has been kept; try again when the service is available.", 503)
}
