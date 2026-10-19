import { randomBytes } from "node:crypto"
import { engineV2Fetch } from "../../lib/engine-v2"
import { isProjectedUserAdmission } from "../../lib/engine-v2-projection"
import { nativeStageResult } from "./native-stage-result"
import type { StageSessionPort } from "./stages"

// Native engine session port: the production path. Child sessions are created
// under the run's owner session with the run's selected model; completion is the
// native idle record after the admitted prompt.
export function engineSessionPort(options: { engineURL?: string; fetch?: typeof fetch; pollMs?: number; timeoutMs?: number } = {}): StageSessionPort {
  const base = (options.engineURL ?? process.env.OPENCODE_URL ?? `http://127.0.0.1:${process.env.DOC_HAUS_ENGINE_PORT ?? "4097"}`).replace(/\/$/, "")
  const transport = options.fetch ?? fetch
  const request = engineV2Fetch(transport)
  const pollMs = options.pollMs ?? 2_000
  const timeoutMs = options.timeoutMs ?? 90 * 60_000
  return {
    async createSession(input) {
      const response = await request(`${base}/api/session`, {
        method: "POST",
        headers: { "content-type": "application/json", "x-opencode-directory": input.directory },
        body: JSON.stringify({
          ...(input.parentID ? { parentID: input.parentID } : {}),
          title: input.title,
          agent: input.agent,
          ...(input.model ? { model: { providerID: input.model.providerID, id: input.model.modelID } } : {}),
          metadata: { dochaus_role: "lean-drafting", ...input.metadata },
          ...(input.permissions ? { permissions: input.permissions } : {}),
        }),
        signal: AbortSignal.timeout(30_000),
      })
      if (!response.ok) throw new Error(`The engine could not create the ${input.agent} session (HTTP ${response.status}).`)
      const session: unknown = await response.json()
      if (!record(session) || typeof session.id !== "string") throw new Error("The engine returned an invalid session.")
      return session.id
    },
    async prompt(input) {
      const messageID = ascendingMessageID()
      const admitted = await request(`${base}/api/session/${encodeURIComponent(input.sessionID)}/prompt`, {
        method: "POST",
        headers: { "content-type": "application/json", "x-opencode-directory": input.directory },
        body: JSON.stringify({ messageID, text: input.text, delivery: "queue" }),
        signal: AbortSignal.timeout(30_000),
      })
      // Typed fields of a failure travel as the error's cause (providerStopKind).
      if (!admitted.ok) throw new Error(`The engine refused the ${input.agent} prompt (HTTP ${admitted.status}).`, { cause: await admitted.json().catch(() => undefined) })
      // A pause that arrived while this prompt was being admitted found nothing
      // to interrupt yet; stop the admitted turn here instead.
      if (input.signal?.aborted) {
        await interrupt(input).catch(() => undefined)
        throw new Error(`The ${input.agent} session was paused.`)
      }
      const deadline = Date.now() + timeoutMs
      while (Date.now() < deadline) {
        await Bun.sleep(pollMs)
        if (input.signal?.aborted) throw new Error(`The ${input.agent} session was paused.`)
        const idle = await idleOutcome(transport, base, input.directory, input.sessionID)
        if (!idle) continue
        const { outcome } = idle
        const listed = await request(`${base}/api/session/${encodeURIComponent(input.sessionID)}/message`, {
          headers: { "x-opencode-directory": input.directory }, signal: AbortSignal.timeout(30_000),
        })
        if (!listed.ok) continue
        const messages: unknown = await listed.json()
        if (!Array.isArray(messages)) continue
        // An idle record from before this prompt was promoted is not its completion.
        const latestUser = messages.findLast(isProjectedUserAdmission)
        if (!record(latestUser) || !record(latestUser.info) || latestUser.info.id !== messageID) continue
        // The engine resumes an interrupted provider response with a synthetic
        // continuation, so only the turn's last reply decides whether it failed.
        const last = messages.filter(record).findLast(message => record(message.info) && message.info.role === "assistant" &&
          message.info.parentID === messageID)
        // A turn that failed before any reply has its error on the idle record only.
        const error = record(last) && record(last.info) && last.info.error ? last.info.error : outcome === "failed" ? idle.error : undefined
        if (outcome !== "succeeded" || error)
          throw new Error(`The ${input.agent} session ended ${outcome}${error ? `: ${JSON.stringify(error)}` : ""}.`, { cause: error })
        return nativeStageResult(messages, messageID, input.agent)
      }
      throw new Error(`The ${input.agent} session did not finish within ${Math.round(timeoutMs / 60_000)} minutes.`)
    },
    interrupt,
  }

  async function interrupt(input: { directory: string; sessionID: string }) {
    const response = await request(`${base}/api/session/${encodeURIComponent(input.sessionID)}/interrupt`, {
      method: "POST", headers: { "x-opencode-directory": input.directory }, signal: AbortSignal.timeout(30_000),
    })
    if (!response.ok) throw new Error(`The engine could not interrupt the session (HTTP ${response.status}).`)
  }
}

async function idleOutcome(transport: typeof fetch, base: string, directory: string, sessionID: string) {
  const password = process.env.OPENCODE_SERVER_PASSWORD
  const response = await transport(`${base}/api/session/${encodeURIComponent(sessionID)}/message?order=desc&limit=1`, {
    headers: { "x-opencode-directory": directory, ...(password ? { authorization: `Basic ${btoa(`opencode:${password}`)}` } : {}) },
    signal: AbortSignal.timeout(30_000),
  }).catch(() => undefined)
  if (!response?.ok) return
  const envelope: unknown = await response.json().catch(() => undefined)
  const newest = record(envelope) && Array.isArray(envelope.data) ? envelope.data[0] : undefined
  if (!record(newest) || newest.type !== "idle") return
  return { outcome: typeof newest.outcome === "string" ? newest.outcome : "unknown", error: newest.error }
}

// Native message ids sort by time: 12 hex of (ms * 0x1000 + counter), then 14 base62.
const sequence = { next: 0 }
function ascendingMessageID() {
  const time = (BigInt(Date.now()) * 0x1000n + BigInt(sequence.next++ % 0x1000)).toString(16).padStart(12, "0").slice(-12)
  const alphabet = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
  return `msg_${time}${[...randomBytes(14)].map(byte => alphabet[byte % alphabet.length]).join("")}`
}


function record(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value)
}
