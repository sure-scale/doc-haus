import { LLM, LLMClient } from "@opencode/ai"
import { Model } from "@opencode/core/model"
import { ModelResolver } from "@opencode/core/model-resolver"
import { SessionID } from "@opencode/schema/session-id"
import { Effect } from "effect"
import { providerTestFailure } from "./provider-health"

export type ConcurrencyFailure = "auth" | "model" | "rate_limit" | "unavailable" | "empty" | "timeout"
export type ConcurrencyProbe = { ok: true } | { ok: false; kind: ConcurrencyFailure }

// One stateless prompt, no matter data or tools. An empty model response is not
// a successful test, even if the provider returned a successful HTTP response.
export function concurrencyProbe(ref: Model.Ref) {
  return Effect.gen(function* () {
    const model = yield* Model.Service.use(models => models.get(ref.providerID, ref.id))
    if (!model?.enabled) return { ok: false, kind: "model" } as const
    const resolved = yield* ModelResolver.Service.use(resolver => resolver.resolve(ref))
    if (!resolved) return { ok: false, kind: "model" } as const
    const result = yield* LLMClient.Service.use(client => client.generate(LLM.request({
      model: resolved.model,
      prompt: "Write eight short friendly sentences, each on a separate line, about everyday activities.",
      http: { headers: { "x-opencode-session": SessionID.create() } },
    })))
    return result.text.trim() ? { ok: true } as const : { ok: false, kind: "empty" } as const
  }).pipe(Effect.catchCause(cause => {
    if (isRateLimit(cause)) return Effect.succeed<ConcurrencyProbe>({ ok: false, kind: "rate_limit" })
    const failure = providerTestFailure(cause)
    return Effect.succeed<ConcurrencyProbe>({ ok: false, kind: failure.ok ? "unavailable" : failure.kind })
  }))
}

// Inspect structured status only; never surface a provider response body.
function isRateLimit(error: unknown): boolean {
  const queue = [error]
  const visited = new Set<unknown>()
  for (let index = 0; index < queue.length && index < 64; index++) {
    const value = queue[index]
    if (!value || typeof value !== "object" || visited.has(value)) continue
    visited.add(value)
    const item = value as Record<string, unknown>
    if (item.status === 429 || item.statusCode === 429 || item._tag === "RateLimit") return true
    for (const key of ["http", "reason", "cause", "error", "defect"]) if (item[key] !== undefined) queue.push(item[key])
    if (Array.isArray(item.reasons)) queue.push(...item.reasons.slice(0, 8))
  }
  return false
}

export type ConcurrencyTestResult = {
  model: string
  requested: number
  succeeded: number
  failed: number
  peak: number
  elapsed_ms: number
  errors: { kind: ConcurrencyFailure; count: number }[]
}

// Observer callbacks are called by the shared model admission while a test
// request holds a permit. This measures our dispatch overlap, not the provider's.
export async function runConcurrencyTest(model: string, count: number, run: (observer: { start(): void; end(): void }) => Promise<ConcurrencyProbe>): Promise<ConcurrencyTestResult> {
  const started = performance.now()
  let active = 0, peak = 0
  const observer = { start() { active++; peak = Math.max(peak, active) }, end() { active-- } }
  const results = await Promise.all(Array.from({ length: count }, () => run(observer).catch((): ConcurrencyProbe => ({ ok: false, kind: "unavailable" }))))
  const failures = results.filter((result): result is Extract<ConcurrencyProbe, { ok: false }> => !result.ok)
  const errors = new Map<ConcurrencyFailure, number>()
  for (const result of failures) errors.set(result.kind, (errors.get(result.kind) ?? 0) + 1)
  return { model, requested: count, succeeded: count - failures.length, failed: failures.length, peak,
    elapsed_ms: Math.round(performance.now() - started), errors: [...errors].map(([kind, count]) => ({ kind, count })) }
}
