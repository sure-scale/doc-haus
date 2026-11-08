import { AIError, isRetryable, LLMClient, type LLMRequest } from "@opencode/ai/effect"
import { llmClient } from "@opencode/core/effect/app-node-platform"
import { MODEL_CAPACITY_SCOPE } from "@opencode/core/model-resolver"
import { Context, Effect, Layer, Stream } from "effect"
import { modelConcurrency, readEffectiveSettings } from "./effective-settings"

export type ModelAdmissionObserver = { start: () => void; end: () => void }
type Waiter = { resume: (effect: Effect.Effect<() => void>) => void; observer?: ModelAdmissionObserver; scope?: string }
type Capacity = { until: number; next: number; remaining?: number }
const CAPACITY_DELAY_MAX = 15 * 60 * 1000
const CAPACITY_DELAY_DEFAULT = 2000
const REOPEN_GAP = 250

/** One engine-wide queue. Permits cover provider operations, never an agent's
 * entire lifetime: a parent waiting for a child does not occupy a slot. */
export class ModelAdmission {
  private active = 0
  private waiting: Waiter[] = []
  private capacityWaits = new Map<string, Capacity>()
  private timer: ReturnType<typeof setInterval> | undefined
  constructor(private readonly limit: () => number) {}

  /** A provider refusal delays admission, never repeats the operation. Native
   * ModelRetry retains retry ownership and cancellation for the failed call. */
  capacity(scope: string | undefined, error: unknown) {
    if (!scope || !(error instanceof AIError) || error.reason._tag !== "RateLimit" || !isRetryable(error)) return false
    const hint = error.reason.retryAfterMs ?? error.reason.rateLimit?.retryAfterMs
    if (hint !== undefined && (!Number.isFinite(hint) || hint < 0)) return false
    const delay = hint === undefined ? CAPACITY_DELAY_DEFAULT : Math.min(Math.ceil(hint), CAPACITY_DELAY_MAX)
    const until = Math.max(this.capacityWaits.get(scope)?.until ?? 0, Date.now() + delay)
    this.capacityWaits.set(scope, { until, next: until })
    this.pump()
    return true
  }

  private pump() {
    let maximum: number
    try { maximum = this.limit() } catch (error) {
      // Invalid saved settings must not silently disable the limit or strand jobs.
      for (const waiter of this.waiting.splice(0)) waiter.resume(Effect.die(error))
      this.stopTimer()
      return
    }
    const now = Date.now()
    // Native backoff or other active routes can outlast the provider's hint.
    // End recovery after one actual admission wave, not elapsed idle time.
    while (this.waiting.length && this.active < maximum) {
      const index = this.waiting.findIndex(waiter => !waiter.scope || !this.capacityWaits.has(waiter.scope) || this.capacityWaits.get(waiter.scope)!.next <= now)
      if (index < 0) break
      const [waiter] = this.waiting.splice(index, 1)
      if (waiter!.scope) {
        const capacity = this.capacityWaits.get(waiter!.scope)
        if (capacity) {
          capacity.remaining ??= maximum
          capacity.remaining--
          capacity.next = now + REOPEN_GAP
          if (!capacity.remaining) this.capacityWaits.delete(waiter!.scope)
        }
      }
      this.active++
      try { waiter!.observer?.start() } catch { /* Observability never owns admission. */ }
      let released = false
      waiter!.resume(Effect.succeed(() => {
        if (released) return
        released = true
        this.active--
        try { waiter!.observer?.end() } catch { /* Observability never owns admission. */ }
        this.pump()
      }))
    }
    if (!this.waiting.length) this.stopTimer()
  }

  private stopTimer() {
    if (this.timer) clearInterval(this.timer)
    this.timer = undefined
  }

  permit(observer?: ModelAdmissionObserver, scope?: string) {
    const acquire = Effect.callback<() => void>(resume => {
      const waiter: Waiter = { resume, observer, scope }
      this.waiting.push(waiter)
      this.pump()
      // Settings are written by ingest, a separate process. Poll only while jobs
      // wait so raising the limit admits new work without waiting for a release.
      if (this.waiting.length && !this.timer) {
        this.timer = setInterval(() => this.pump(), 250)
        this.timer.unref?.()
      }
      return Effect.sync(() => {
        const index = this.waiting.indexOf(waiter)
        if (index >= 0) this.waiting.splice(index, 1)
        if (!this.waiting.length) this.stopTimer()
      })
    })
    return Effect.acquireRelease(acquire, release => Effect.sync(release), { interruptible: true })
  }
}

export const modelAdmission = new ModelAdmission(() => modelConcurrency(readEffectiveSettings()))

/** The resolver's access identity contains only the selected connection id or
 * environment name. No credential, request header or endpoint query is read. */
function requestScope(request: LLMRequest) {
  const model = request.model
  const auth = model?.route?.auth
  if (!auth) return
  const access = Object.getOwnPropertyDescriptor(auth, MODEL_CAPACITY_SCOPE)?.value
  if (typeof access !== "string" || !access) return
  return JSON.stringify([access, model.provider, model.id, model.route.id, model.route.protocol])
}

/** Wrap exactly once. In particular the separately built health/test client
 * must be wrapped from its raw instance, not from another wrapped client. */
export function limitModelClient(client: Context.Service.Shape<typeof LLMClient.Service>, observer?: ModelAdmissionObserver, admission = modelAdmission): Context.Service.Shape<typeof LLMClient.Service> {
  const observe = (scope: string | undefined) => (error: unknown) => Effect.sync(() => { admission.capacity(scope, error) })
  const compact: typeof client.compact = ((request: never, options: never) =>
    Effect.scoped(Effect.flatMap(admission.permit(observer, requestScope(request)), () => client.compact(request, options).pipe(Effect.tapError(observe(requestScope(request))))))) as typeof client.compact
  return {
    stream: (request, options) => {
      const scope = requestScope(request)
      return Stream.unwrap(Effect.map(admission.permit(observer, scope), release => client.stream(request, options).pipe(
        Stream.tapError(observe(scope)), Stream.ensuring(Effect.sync(release)),
      )))
    },
    generate: (request, options) => Effect.scoped(Effect.flatMap(admission.permit(observer, requestScope(request)), () => client.generate(request, options).pipe(Effect.tapError(observe(requestScope(request)))))),
    compact,
  }
}

// Replace the engine's global LLM client node: all location/session graphs share
// one queue without changing upstream packages or serializing unrelated HTTP.
export const makeModelConcurrencyNode = (admission = modelAdmission) => llmClient.replace(llmClient.mapLayer(layer => layer.pipe(
  Layer.flatMap(context => Layer.succeed(LLMClient.Service, limitModelClient(Context.get(context, LLMClient.Service), undefined, admission))),
)))
export const modelConcurrencyNode = makeModelConcurrencyNode()
