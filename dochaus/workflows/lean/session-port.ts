import { randomBytes } from "node:crypto"
import { isDeepStrictEqual } from "node:util"
import { SessionID } from "@opencode/schema/session-id"
import { engineV2Fetch } from "../../lib/engine-v2"
import { isProjectedUserAdmission } from "../../lib/engine-v2-projection"
import { providerStopKind } from "../../lib/provider-stop"
import { nativeStageResult } from "./native-stage-result"
import type { StageSessionPort } from "./stages"

export type NativeSessionWait = { sessionID: string; parentID?: string; phase: string; lastSuccessfulContact: number | null; nextRetry: number }
const connectionWaits = new Map<string, Map<string, NativeSessionWait>>()
/** Display-only process memory. No run state, generation permission or recovery
 * decision is derived from this snapshot; native storage remains authoritative. */
export function nativeSessionWaits(directory: string): NativeSessionWait[] {
  return [...(connectionWaits.get(directory)?.values() ?? [])].map(row => ({ ...row }))
}

// Native engine session port: the production path. Child sessions are created
// under the run's owner session with the run's selected model; completion is the
// native idle record after the admitted prompt.
export function engineSessionPort(options: { engineURL?: string; fetch?: typeof fetch; pollMs?: number; timeoutMs?: number; retryDelayMs?: number } = {}): StageSessionPort {
  const base = (options.engineURL ?? process.env.OPENCODE_URL ?? `http://127.0.0.1:${process.env.DOC_HAUS_ENGINE_PORT ?? "4097"}`).replace(/\/$/, "")
  const transport = options.fetch ?? fetch
  const pollMs = options.pollMs ?? 2_000
  type Context = { directory: string; sessionID?: string; parentID?: string; messageID?: string; signal?: AbortSignal; limit?: AbortSignal; lastSuccessfulContact?: number }
  function clearWait(context: Context) {
    const waiting = connectionWaits.get(context.directory)
    if (context.sessionID) waiting?.delete(context.sessionID)
    if (!waiting?.size) connectionWaits.delete(context.directory)
  }
  function contact(context: Context) {
    context.lastSuccessfulContact = Date.now()
    const waiting = context.sessionID && connectionWaits.get(context.directory)?.get(context.sessionID)
    if (waiting) waiting.lastSuccessfulContact = context.lastSuccessfulContact
  }
  function backoff(attempt: number) { return Math.min(Math.max(1, options.retryDelayMs ?? 500) * 2 ** Math.min(attempt, 14), 10_000) }
  function showWait(context: Context, phase: string, ms: number) {
    if (!context.sessionID) return
    const waiting = connectionWaits.get(context.directory) ?? new Map<string, NativeSessionWait>()
    waiting.set(context.sessionID, { sessionID: context.sessionID, ...(context.parentID ? { parentID: context.parentID } : {}), phase,
      lastSuccessfulContact: context.lastSuccessfulContact ?? null, nextRetry: Date.now() + ms })
    connectionWaits.set(context.directory, waiting)
  }
  function stopped(context: Context, cause?: unknown) {
    if (context.signal?.aborted) { clearWait(context); throw new Error("The lean session was paused.", { cause }) }
    if (context.limit?.aborted) { clearWait(context); throw new Error(`Native session observation reached its explicit ${options.timeoutMs}ms wait limit (session=${context.sessionID ?? "new"}; message=${context.messageID ?? "none"}).`, { cause }) }
  }
  async function wait(ms: number, context: Context, cause?: unknown) {
    stopped(context, cause)
    await new Promise<void>((resolve, reject) => {
      const signals = [context.signal, context.limit].filter((signal): signal is AbortSignal => !!signal)
      const cleanup = () => { clearTimeout(timer); for (const signal of signals) signal.removeEventListener("abort", abort) }
      const abort = () => {
        cleanup()
        try { stopped(context, cause) } catch (error) { reject(error) }
      }
      const timer = setTimeout(() => { cleanup(); resolve() }, Math.max(1, ms))
      for (const signal of signals) signal.addEventListener("abort", abort, { once: true })
    })
  }
  // A transport retry is read-only. Admission is sent once with one durable ID;
  // an uncertain acknowledgement is reconciled against native storage instead.
  async function call(phase: string, route: string, context: Context, init: RequestInit = {}, projected = true, reconnect = false) {
    for (let attempt = 0; ; attempt++) {
      stopped(context)
      const headers = new Headers(init.headers)
      headers.set("x-opencode-directory", context.directory)
      const password = process.env.OPENCODE_SERVER_PASSWORD
      if (password) headers.set("authorization", `Basic ${btoa(`opencode:${password}`)}`)
      const timing = { started: performance.now(), deadlineAt: Date.now() + 30_000 }
      const deadline = AbortSignal.timeout(30_000)
      const signal = AbortSignal.any([deadline, ...[context.signal, context.limit].filter((signal): signal is AbortSignal => !!signal)])
      let currentRoute = `${init.method ?? "GET"} ${new URL(`${base}${route}`).pathname}`
      let nativeHead: Record<string, unknown> | undefined
      const observed = (async (input, options) => {
        const url = new URL(input instanceof Request ? input.url : String(input))
        currentRoute = `${options?.method ?? "GET"} ${url.pathname}`
        let response: Response
        try { response = await transport(input, options) }
        catch (cause) { throw new NativeSessionTransportError(phase, currentRoute, context, timing, cause) }
        // Fetch may resolve its headers before a socket reset rejects json().
        // Classify TypeError only at this native body-consumption boundary;
        // JSON syntax and adapter ownership/schema/projection failures retain
        // their own error types and cannot trigger admission recovery or retry.
        const bodyJSON = response.json.bind(response)
        response.json = async () => {
          try {
            const value = await bodyJSON()
            // Native message pages are newest-first by durable sequence. Keep
            // only the first page's terminal header before projection removes
            // idle markers; ownership reads and older pages cannot replace it.
            if (phase === "result-read" && url.pathname.endsWith("/message") && !url.searchParams.has("cursor")) {
              const head = record(value) && Array.isArray(value.data) ? value.data[0] : undefined
              nativeHead = record(head) ? { id: head.id, type: head.type, outcome: head.outcome, error: head.error } : undefined
            }
            contact(context)
            return value
          }
          catch (cause) {
            if (cause instanceof TypeError || transportFailure(cause))
              throw new NativeSessionTransportError(phase, currentRoute, context, timing, cause)
            throw cause
          }
        }
        return response
      }) as typeof fetch
      try {
        const request = projected ? engineV2Fetch(observed) : observed
        const response = await request(`${base}${route}`, { ...init, headers, signal })
        let value: unknown
        if (response.status !== 204) {
          try { value = await response.json() }
          catch (cause) {
            if (cause instanceof NativeSessionTransportError) throw cause
            if (transportFailure(cause)) throw new NativeSessionTransportError(phase, currentRoute, context, timing, cause)
            // A proxy's temporary availability response may be HTML or plain
            // text. Its status still governs reconnecting; malformed successful
            // native JSON remains a hard failure, as do non-syntax body errors.
            if (response.ok || [502, 503, 504].includes(response.status) && !(cause instanceof SyntaxError)) throw cause
          }
        }
        if ([502, 503, 504].includes(response.status) &&
          !providerStopKind(new Error("Native availability", { cause: record(value) && record(value.error) ? value.error : value })))
          throw new NativeSessionTransportError(phase, currentRoute, context, timing, new Error(`HTTP ${response.status}`))
        clearWait(context)
        return { response, value, nativeHead }
      } catch (cause) {
        const error = cause instanceof NativeSessionTransportError ? cause
          : transportFailure(cause) ? new NativeSessionTransportError(phase, currentRoute, context, timing, cause) : cause
        stopped(context, error)
        if (!(error instanceof NativeSessionTransportError) || !reconnect) throw error
        // Native execution owns provider retries. This only waits to observe
        // the same durable work again after connection loss, never resubmits it.
        const ms = backoff(attempt)
        showWait(context, phase, ms)
        await wait(ms, context, error)
      }
    }
  }
  async function reconcile(input: Context & { sessionID: string; messageID: string; text: string }) {
    const route = `/api/session/${encodeURIComponent(input.sessionID)}`
    const pending = await call("admission-readback", `${route}/inbox`, input, {}, true, true)
    if (!pending.response.ok || !Array.isArray(pending.value)) throw new Error("The engine could not verify uncertain prompt admission.", { cause: pending.value })
    const queued = pending.value.filter(record).find(row => row.id === input.messageID)
    if (queued) {
      if (queued.sessionID !== input.sessionID || queued.type !== "user" || queued.delivery !== "queue" ||
        !record(queued.payload) || queued.payload.text !== input.text) throw new Error("The durable prompt admission does not match this request.")
      return true
    }
    const promoted = await call("admission-readback", `${route}/message/${encodeURIComponent(input.messageID)}`, input, {}, true, true)
    if (promoted.response.status === 404) return false
    if (!promoted.response.ok) throw new Error("The engine could not verify uncertain prompt admission.", { cause: promoted.value })
    const message = promoted.value
    if (!record(message) || !record(message.info) || !isProjectedUserAdmission(message) || message.info.type !== "user" ||
      message.info.id !== input.messageID || message.info.sessionID !== input.sessionID || !Array.isArray(message.parts) ||
      message.parts.length !== 1 || !record(message.parts[0]) || message.parts[0].type !== "text" || message.parts[0].text !== input.text)
      throw new Error("The durable prompt admission does not match this request.")
    return true
  }
  return {
    async createSession(input) {
      const sessionID = SessionID.create()
      const context: Context = { directory: input.directory, sessionID, parentID: input.parentID, signal: input.signal,
        ...(options.timeoutMs !== undefined ? { limit: AbortSignal.timeout(options.timeoutMs) } : {}) }
      const metadata = { dochaus_role: "lean-drafting", ...input.metadata }
      const validate = (session: unknown) => {
        if (!record(session) || session.id !== sessionID || session.directory !== input.directory || session.parentID !== input.parentID ||
          session.agent !== input.agent || !isDeepStrictEqual(session.metadata, metadata) ||
          input.permissions && !isDeepStrictEqual(session.permissions, input.permissions) ||
          (input.model ? !record(session.model) || session.model.providerID !== input.model.providerID || session.model.id !== input.model.modelID : session.model !== undefined))
          throw new Error("The native session does not match its requested identity, ownership or selection.")
        return sessionID
      }
      try {
        for (let attempt = 0; ; attempt++) {
          try {
            const created = await call("session-create", "/api/session", context, {
              method: "POST", headers: { "content-type": "application/json" },
              body: JSON.stringify({ id: sessionID, ...(input.parentID ? { parentID: input.parentID } : {}), title: input.title, agent: input.agent,
                ...(input.model ? { model: { providerID: input.model.providerID, id: input.model.modelID } } : {}), metadata,
                ...(input.permissions ? { permissions: input.permissions } : {}) }),
            })
            if (!created.response.ok) throw new Error(`The engine could not create the ${input.agent} session (HTTP ${created.response.status}).`, { cause: created.value })
            return validate(created.value)
          } catch (error) {
            if (!(error instanceof NativeSessionTransportError)) throw error
            const found = await call("session-create-readback", `/api/session/${encodeURIComponent(sessionID)}`, context, {}, true, true)
            if (found.response.ok) return validate(found.value)
            if (found.response.status !== 404) throw new Error(`The engine could not reconcile native session creation (HTTP ${found.response.status}).`, { cause: found.value })
            // Native Session.create(id) returns an existing exact ID and creates
            // it only when absent. Only this native idempotent mutation repeats.
            const ms = backoff(attempt)
            showWait(context, "session-create", ms)
            await wait(ms, context, error)
          }
        }
      } finally { clearWait(context) }
    },
    async prompt(input) {
      if (input.signal?.aborted) throw new Error(`The ${input.agent} session was paused.`)
      const messageID = ascendingMessageID()
      const context = { ...input, messageID, ...(options.timeoutMs !== undefined ? { limit: AbortSignal.timeout(options.timeoutMs) } : {}) }
      try {
        try {
          const admitted = await call("prompt-admission", `/api/session/${encodeURIComponent(input.sessionID)}/prompt`, context, {
            method: "POST", headers: { "content-type": "application/json" },
            body: JSON.stringify({ messageID, text: input.text, delivery: "queue" }),
          })
          // Typed fields of a refusal remain available to providerStopKind.
          if (!admitted.response.ok) throw new Error(`The engine refused the ${input.agent} prompt (HTTP ${admitted.response.status}).`, { cause: admitted.value })
        } catch (error) {
          if (!(error instanceof NativeSessionTransportError)) throw error
          if (!await reconcile(context)) throw new Error(`${error.message} Exact durable admission remains unproven; the prompt was not resent.`, { cause: error })
        }
        // A pause that arrived while this prompt was being admitted found nothing
        // to interrupt yet; stop the admitted turn here instead.
        if (input.signal?.aborted) throw new Error(`The ${input.agent} session was paused.`)
        for (;;) {
          await wait(pollMs, context)
          const newest = await call("idle-read", `/api/session/${encodeURIComponent(input.sessionID)}/message?order=desc&limit=1`, context, {}, false, true)
          if (!newest.response.ok) throw new Error(`The engine could not read the session outcome (HTTP ${newest.response.status}).`, { cause: newest.value })
          const envelope = newest.value
          if (!record(envelope) || !Array.isArray(envelope.data) || envelope.data.length > 1 ||
            envelope.data.some(row => !record(row) || typeof row.type !== "string")) throw new Error("The engine returned an invalid native idle envelope.")
          const native = envelope.data[0]
          const idle = record(native) && native.type === "idle" ? native : undefined
          if (!idle) continue
          const listed = await call("result-read", `/api/session/${encodeURIComponent(input.sessionID)}/message`, context, {}, true, true)
          if (!listed.response.ok) throw new Error(`The engine could not read the session result (HTTP ${listed.response.status}).`, { cause: listed.value })
          const messages = listed.value
          if (!Array.isArray(messages)) continue
          // The poll is only a hint. A queued prompt may promote between it and
          // transcript collection. Only this transcript snapshot's native head
          // can establish terminality for its exact latest admitted user.
          const terminal = listed.nativeHead
          if (terminal?.type !== "idle") continue
          if (typeof terminal.outcome !== "string" || !["succeeded", "failed", "interrupted"].includes(terminal.outcome))
            throw new Error("The engine returned an invalid native idle outcome.")
          const outcome = terminal.outcome
          const latestUser = messages.findLast(isProjectedUserAdmission)
          if (!record(latestUser) || !record(latestUser.info) || latestUser.info.id !== messageID) {
            if (messages.some(row => record(row) && record(row.info) && row.info.id === messageID))
              throw new Error(`The ${input.agent} session turn ${messageID} was superseded by another durable prompt.`)
            continue
          }
          // The engine resumes an interrupted provider response with a synthetic
          // continuation, so only the turn's last reply decides whether it failed.
          const last = messages.filter(record).findLast(message => record(message.info) && message.info.role === "assistant" &&
            message.info.parentID === messageID)
          // A turn that failed before any reply has its error on the idle record only.
          const error = record(last) && record(last.info) && last.info.error ? last.info.error : outcome === "failed" ? terminal.error : undefined
          // The native error's own message, not its payload; its typed fields stay in the cause.
          if (outcome !== "succeeded" || error)
            throw new Error(`The ${input.agent} session ended ${outcome}${error ? `: ${errorDetail(error)}` : ""}.`, { cause: error })
          return nativeStageResult(messages, messageID, input.agent)
        }
      } catch (error) {
        // Admission can commit before an aborted HTTP response reaches us.
        // Interrupt our existing session with a fresh bounded request, including
        // pauses during admission readback or completed-result collection.
        if (input.signal?.aborted) {
          await interrupt({ directory: input.directory, sessionID: input.sessionID }).catch(() => undefined)
          throw new Error(`The ${input.agent} session was paused.`, { cause: error })
        }
        throw error
      } finally { clearWait(context) }
    },
    interrupt,
  }

  async function interrupt(input: { directory: string; sessionID: string }) {
    const { response } = await call("session-interrupt", `/api/session/${encodeURIComponent(input.sessionID)}/interrupt`, input, { method: "POST" })
    if (!response.ok) throw new Error(`The engine could not interrupt the session (HTTP ${response.status}).`)
  }
}

class NativeSessionTransportError extends Error {
  readonly timeoutMs = 30_000
  readonly elapsedMs: number
  readonly deadlineAt: number
  readonly sessionID: string | undefined
  readonly messageID: string | undefined
  constructor(readonly phase: string, readonly route: string, context: { sessionID?: string; messageID?: string }, timing: { started: number; deadlineAt: number }, cause: unknown) {
    const elapsed = Math.round(performance.now() - timing.started)
    super(`Native session transport failed during ${phase} (${route}; session=${context.sessionID ?? "new"}; message=${context.messageID ?? "none"}; timeout=30000ms; elapsed=${elapsed}ms; deadline=${new Date(timing.deadlineAt).toISOString()}).`, { cause })
    this.name = "NativeSessionTransportError"
    this.elapsedMs = elapsed
    this.deadlineAt = timing.deadlineAt
    this.sessionID = context.sessionID
    this.messageID = context.messageID
  }
}

function transportFailure(error: unknown) {
  return error instanceof Error && ["AbortError", "TimeoutError"].includes(error.name)
}

// Native message ids sort by time: 12 hex of (ms * 0x1000 + counter), then 14 base62.
const sequence = { next: 0 }
function ascendingMessageID() {
  const time = (BigInt(Date.now()) * 0x1000n + BigInt(sequence.next++ % 0x1000)).toString(16).padStart(12, "0").slice(-12)
  const alphabet = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
  return `msg_${time}${[...randomBytes(14)].map(byte => alphabet[byte % alphabet.length]).join("")}`
}

function errorDetail(error: unknown) {
  const text = record(error) && typeof error.message === "string" ? error.message.trim().replace(/\.+$/, "") : ""
  return text || JSON.stringify(error)
}

function record(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value)
}
