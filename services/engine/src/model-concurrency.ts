import { LLMClient } from "@opencode/ai"
import { llmClient } from "@opencode/core/effect/app-node-platform"
import { Context, Effect, Layer, Stream } from "effect"
import { modelConcurrency, readEffectiveSettings } from "./effective-settings"

export type ModelAdmissionObserver = { start: () => void; end: () => void }
type Waiter = { resume: (effect: Effect.Effect<() => void>) => void; observer?: ModelAdmissionObserver }

/** One engine-wide queue. Permits cover provider operations, never an agent's
 * entire lifetime: a parent waiting for a child does not occupy a slot. */
export class ModelAdmission {
  private active = 0
  private waiting: Waiter[] = []
  private timer: ReturnType<typeof setInterval> | undefined
  constructor(private readonly limit: () => number) {}

  private pump() {
    let maximum: number
    try { maximum = this.limit() } catch (error) {
      // Invalid saved settings must not silently disable the limit or strand jobs.
      for (const waiter of this.waiting.splice(0)) waiter.resume(Effect.die(error))
      this.stopTimer()
      return
    }
    while (this.waiting.length && this.active < maximum) {
      const waiter = this.waiting.shift()!
      this.active++
      try { waiter.observer?.start() } catch { /* Observability never owns admission. */ }
      let released = false
      waiter.resume(Effect.succeed(() => {
        if (released) return
        released = true
        this.active--
        try { waiter.observer?.end() } catch { /* Observability never owns admission. */ }
        this.pump()
      }))
    }
    if (!this.waiting.length) this.stopTimer()
  }

  private stopTimer() {
    if (this.timer) clearInterval(this.timer)
    this.timer = undefined
  }

  permit(observer?: ModelAdmissionObserver) {
    const acquire = Effect.callback<() => void>(resume => {
      const waiter: Waiter = { resume, observer }
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

/** Wrap exactly once. In particular the separately built health/test client
 * must be wrapped from its raw instance, not from another wrapped client. */
export function limitModelClient(client: Context.Service.Shape<typeof LLMClient.Service>, observer?: ModelAdmissionObserver, admission = modelAdmission): Context.Service.Shape<typeof LLMClient.Service> {
  const compact: typeof client.compact = ((request: never, options: never) =>
    Effect.scoped(Effect.flatMap(admission.permit(observer), () => client.compact(request, options)))) as typeof client.compact
  return {
    stream: (request, options) => Stream.unwrap(Effect.map(admission.permit(observer), () => client.stream(request, options))),
    generate: (request, options) => Effect.scoped(Effect.flatMap(admission.permit(observer), () => client.generate(request, options))),
    compact,
  }
}

// Replace the engine's global LLM client node: all location/session graphs share
// one queue without changing upstream packages or serializing unrelated HTTP.
export const makeModelConcurrencyNode = (admission = modelAdmission) => llmClient.replace(llmClient.mapLayer(layer => layer.pipe(
  Layer.flatMap(context => Layer.succeed(LLMClient.Service, limitModelClient(Context.get(context, LLMClient.Service), undefined, admission))),
)))
export const modelConcurrencyNode = makeModelConcurrencyNode()
