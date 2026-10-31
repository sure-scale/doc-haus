import { OpenCode } from "@opencode/client"
import type { Context, Hono } from "hono"
import { streamSSE } from "hono/streaming"
import { randomUUID } from "node:crypto"
import { engineURL, type EngineRequestOptions } from "./matter-session"
import type { WordChatAuth } from "./word-chat-context"

const ID = /^[A-Za-z0-9_-]{1,160}$/
const BODY_LIMIT = 2048
const KEEPALIVE_MS = 15_000
const record = (value: unknown): value is Record<string, unknown> => Boolean(value && typeof value === "object" && !Array.isArray(value))

async function conversationInput(c: Context): Promise<string | undefined> {
  if (Number(c.req.header("content-length") ?? "0") > BODY_LIMIT) return
  const reader = c.req.raw.body?.getReader()
  if (!reader) return
  const chunks: Uint8Array[] = []
  let size = 0
  try {
    for (;;) {
      const { value, done } = await reader.read()
      if (done) break
      size += value.byteLength
      if (size > BODY_LIMIT) { await reader.cancel(); return }
      chunks.push(value)
    }
    const input: unknown = JSON.parse(Buffer.concat(chunks).toString("utf8"))
    return record(input) && typeof input.conversationId === "string" && ID.test(input.conversationId) ? input.conversationId : undefined
  } catch { return }
  finally { reader.releaseLock() }
}

/** A linked Word capability observes its own conversation, never the engine's
 * global event feed. Native events only invalidate the existing checked poll
 * path; model text, tools, locations and reviewer identities never cross it. */
export function installWordChatEvents(app: Hono, authenticate: (c: Context) => WordChatAuth | undefined, options: EngineRequestOptions = {}) {
  app.post("/word/chat/events", async c => {
    c.header("Cache-Control", "no-store")
    const auth = authenticate(c)
    if (!auth) return c.json({ error: "Reopen this document from doc.haus." }, 401)
    const requestedConversation = await conversationInput(c)
    if (!requestedConversation) return c.json({ error: "Choose a conversation for this document." }, 400)
    const conversationId: string = requestedConversation
    const headers: Record<string, string> = { "x-opencode-directory": auth.matter.dir }
    if (process.env.OPENCODE_SERVER_PASSWORD)
      headers.Authorization = `Basic ${Buffer.from(`opencode:${process.env.OPENCODE_SERVER_PASSWORD}`).toString("base64")}`
    const client = OpenCode.make({ baseUrl: engineURL(options.engineURL), fetch: options.fetch, headers })
    const live = () => {
      const current = authenticate(c)
      return Boolean(current && current.grant.token_hash === auth.grant.token_hash && current.grant.document_id === auth.grant.document_id && current.matter.dir === auth.matter.dir)
    }
    const owned = (session: { location: { directory: string }; metadata?: Record<string, unknown> }) =>
      session.location.directory === auth.matter.dir && session.metadata?.document === auth.grant.document_id
    const read = (sessionID: string, signal?: AbortSignal) => client.session.get({ sessionID }, { signal: signal ? AbortSignal.any([signal, AbortSignal.timeout(10_000)]) : AbortSignal.timeout(10_000) })
    const root = await read(conversationId).catch(() => undefined)
    if (!root || root.id !== conversationId || root.parentID || !owned(root) || root.metadata?.type !== "word-chat")
      return c.json({ error: "Choose a conversation for this document." }, 404)
    if (!live()) return c.json({ error: "Reopen this document from doc.haus." }, 401)

    return streamSSE(c, async stream => {
      c.header("Cache-Control", "no-store")
      const upstream = new AbortController()
      const stop = () => upstream.abort()
      stream.onAbort(stop)
      c.req.raw.signal.addEventListener("abort", stop, { once: true })
      if (c.req.raw.signal.aborted) stop()
      const heartbeat = setInterval(() => {
        if (!live()) { stop(); return }
        void stream.write(": keepalive\n\n")
      }, KEEPALIVE_MS)
      async function invalidate(sessionID: string, status: "busy" | "idle") {
        if (upstream.signal.aborted || !live()) { stop(); return }
        await stream.writeSSE({ data: JSON.stringify({ id: randomUUID(), created: Date.now(), type: "session.status", data: { sessionID, status: { type: status } } }) })
      }
      async function publicTurn(sessionID: string): Promise<string | undefined> {
        if (!ID.test(sessionID)) return
        const currentRoot = await read(conversationId, upstream.signal).catch(() => undefined)
        if (!currentRoot || currentRoot.id !== conversationId || currentRoot.parentID || !owned(currentRoot) || currentRoot.metadata?.type !== "word-chat") return
        if (sessionID === conversationId) return conversationId
        const seen = new Set<string>()
        let current = sessionID
        let publicID: string | undefined
        for (let depth = 0; depth < 8 && current !== conversationId; depth++) {
          if (seen.has(current) || upstream.signal.aborted) return
          seen.add(current)
          const session: Awaited<ReturnType<typeof read>> | undefined = await read(current, upstream.signal).catch(() => undefined)
          if (!session || session.id !== current || !owned(session) || !session.parentID || !ID.test(session.parentID)) return
          if (session.parentID === conversationId) {
            if (session.metadata?.type !== "word-chat-turn") return
            publicID = session.id
          }
          current = session.parentID
        }
        return current === conversationId ? publicID : undefined
      }
      try {
        if (!live() || upstream.signal.aborted) return
        await stream.writeSSE({ data: JSON.stringify({ id: randomUUID(), created: Date.now(), type: "server.connected", data: {} }) })
        await invalidate(conversationId, "idle")
        for await (const event of client.event.subscribe({ signal: upstream.signal })) {
          if (!live()) { stop(); break }
          if (event.type === "server.connected") continue
          const value: unknown = event
          if (!record(value) || !record(value.data) || typeof value.data.sessionID !== "string") continue
          let status: "busy" | "idle"
          if (value.type === "session.status" && record(value.data.status)) status = value.data.status.type === "idle" ? "idle" : "busy"
          else if (value.type === "session.idle") status = "idle"
          else if (value.type === "session.renamed") status = "idle"
          else if (["session.created", "session.step.started", "session.step.ended", "session.step.failed", "session.inbox.enqueued", "session.inbox.delivered"].includes(String(value.type))) status = "busy"
          else continue
          const turn = await publicTurn(value.data.sessionID)
          if (turn) await invalidate(turn, status)
        }
      } catch {
        // Closing causes the shared client lifecycle to reconnect and reconcile
        // through the authenticated read path; never stream an upstream error.
      } finally {
        clearInterval(heartbeat)
        c.req.raw.signal.removeEventListener("abort", stop)
        stop()
      }
    })
  })
}
