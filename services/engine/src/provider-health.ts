import { LLM, LLMClient } from "@opencode/ai"
import { Model } from "@opencode/core/model"
import { ModelResolver } from "@opencode/core/model-resolver"
import { SessionID } from "@opencode/schema/session-id"
import { Effect } from "effect"

export type ProviderTestResult = { ok: true } | {
  ok: false
  kind: "auth" | "model" | "unavailable"
  expired?: boolean
  error: string
}

export type ProviderTestInput = { providerID: string; modelID: string; force?: boolean }

export function providerTestInput(input: unknown): input is ProviderTestInput {
  if (!input || typeof input !== "object" || Array.isArray(input)) return false
  const value = input as Record<string, unknown>
  return Object.keys(value).every(key => ["providerID", "modelID", "force"].includes(key)) &&
    typeof value.providerID === "string" && /^[a-zA-Z0-9][a-zA-Z0-9._-]{0,127}$/.test(value.providerID) &&
    typeof value.modelID === "string" && value.modelID.length > 0 && value.modelID.length <= 256 && !/[\s#]/.test(value.modelID) &&
    (value.force === undefined || typeof value.force === "boolean")
}

export const unavailableModel: ProviderTestResult = {
  ok: false, kind: "model", error: "This model is unavailable or disabled by provider policy. Choose an available model in Settings.",
}

// Keep the existing stateless Generate request shape, native resolution and
// credentials, but retain typed failures. An added output limit can invalidate
// ChatGPT requests or configured thinking budgets; the server deadline bounds
// this fixed probe instead.
// No matter text, tools or persisted conversation enters this probe.
export function probeProvider(ref: Model.Ref) {
  return Effect.gen(function* () {
    const model = yield* Model.Service.use(models => models.get(ref.providerID, ref.id))
    if (!model?.enabled) return unavailableModel
    const resolved = yield* ModelResolver.Service.use(resolver => resolver.resolve(ref))
    if (!resolved) return unavailableModel
    yield* LLMClient.Service.use(client => client.generate(LLM.request({
      model: resolved.model,
      prompt: "Reply with OK.",
      http: { headers: { "x-opencode-session": SessionID.create() } },
    })))
    return { ok: true } as const
  }).pipe(Effect.catchCause(cause => Effect.succeed(providerTestFailure(cause))))
}

// Inspect only bounded error structure. In particular, quota/transport failures
// must never make a saved account appear disconnected. Raw provider text stays
// inside the engine and is never returned to the browser.
export function providerTestFailure(error: unknown): ProviderTestResult {
  const queue: unknown[] = [error]
  const visited = new Set<unknown>()
  let auth = false
  let expired = false
  let model = false
  for (let index = 0; index < queue.length && index < 32; index++) {
    const value = queue[index]
    if (!value || typeof value !== "object" || visited.has(value)) continue
    visited.add(value)
    const item = value as Record<string, unknown>
    if (item._tag === "Authentication" || item._tag === "Integration.Authorization") {
      auth = true
      expired ||= [item.message, item.body].some(text => typeof text === "string" &&
        /\bkey\b[\s\S]{0,80}\bexpir(?:ed|y|ation)\b|\bexpir(?:ed|y|ation)\b[\s\S]{0,80}\bkey\b/i.test(text.slice(0, 8192)))
    }
    if (item._tag === "Generate.ModelSelectionError" || typeof item._tag === "string" && item._tag.startsWith("SessionRunnerModel.")) model = true
    for (const key of ["reason", "cause", "error", "defect"]) if (item[key] !== undefined) queue.push(item[key])
    if (Array.isArray(item.reasons)) queue.push(...item.reasons.slice(0, 8))
  }
  if (auth) return { ok: false, kind: "auth", ...(expired ? { expired: true } : {}),
    error: expired ? "The provider API key has expired. Replace it in Settings; signing in again will not renew an expired key."
      : "The provider rejected the saved connection. Check or reconnect the provider in Settings." }
  if (model) return unavailableModel
  return { ok: false, kind: "unavailable", error: "The provider could not be checked. Check the connection, model access or account allowance, then retry." }
}

// Process-local only: cache keys are credential/settings fingerprints, never
// secrets. Force retries bypass completed results but still join a current test.
export function providerTestCache(now = Date.now, ttl = 60_000) {
  const entries = new Map<string, { pending: boolean; expires: number; result: Promise<ProviderTestResult> }>()
  return (key: string, force: boolean, check: () => Promise<ProviderTestResult>) => {
    const previous = entries.get(key)
    if (previous && (previous.pending || !force && previous.expires > now())) return previous.result
    for (const [entryKey, entry] of entries) if (!entry.pending && entry.expires <= now()) entries.delete(entryKey)
    if (entries.size >= 128) for (const [entryKey, entry] of entries) {
      if (!entry.pending) { entries.delete(entryKey); break }
    }
    const entry = { pending: true, expires: Infinity, result: Promise.resolve<ProviderTestResult>({ ok: true }) }
    entry.result = Promise.resolve().then(check).catch(providerTestFailure).then(result => {
      entry.pending = false
      entry.expires = now() + ttl
      return result
    })
    entries.set(key, entry)
    return entry.result
  }
}
