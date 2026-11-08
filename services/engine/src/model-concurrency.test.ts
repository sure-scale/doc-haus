import { describe, expect, test } from "bun:test"
import { Effect, Stream } from "effect"
import { ai, AIError, AuthenticationError, HttpContext, LLMClient, QuotaExceededError, RateLimitError } from "@opencode/ai/effect"
import { configure } from "@opencode/ai/providers/openai"
import { llmClient } from "@opencode/core/effect/app-node-platform"
import { httpClient } from "@opencode/util/effect/app-node-platform"
import { LayerNode } from "@opencode/util/effect/layer-node"
import { withCapacityScope } from "@opencode/core/model-resolver"
import { FetchHttpClient } from "effect/http"
import { ModelAdmission, limitModelClient, makeModelConcurrencyNode } from "./model-concurrency"
import type { Context } from "effect"

const until = async (condition: () => boolean) => {
  const deadline = Date.now() + 2000
  while (!condition()) {
    if (Date.now() > deadline) throw new Error("Timed out waiting for admission")
    await Bun.sleep(10)
  }
}
const latch = () => {
  let resolve!: () => void
  const promise = new Promise<void>(done => { resolve = done })
  return { promise, resolve }
}

describe("model request admission", () => {
  test("limits combined streamed and generated requests and releases after completion", async () => {
    const admission = new ModelAdmission(() => 2)
    const controls = Array.from({ length: 3 }, latch)
    let started = 0
    let active = 0
    let peak = 0
    const observer = { start: () => { active++; peak = Math.max(peak, active) }, end: () => { active-- } }
    const fake = {
      stream: () => Stream.fromEffect(Effect.promise(async () => { const current = started++; await controls[current]!.promise; return "result" })),
      generate: () => Effect.promise(async () => { const current = started++; await controls[current]!.promise; return "result" }),
      compact: () => Effect.die("unused"),
    } as unknown as Context.Service.Shape<typeof LLMClient.Service>
    const client = limitModelClient(fake, observer, admission)
    const jobs = [
      Effect.runPromise(Stream.runCollect(client.stream({} as never))),
      Effect.runPromise(client.generate({} as never)),
      Effect.runPromise(client.generate({} as never)),
    ]
    await until(() => started === 2)
    expect(active).toBe(2)
    controls[0]!.resolve()
    await until(() => started === 3)
    controls[1]!.resolve()
    controls[2]!.resolve()
    await Promise.all(jobs)
    expect(peak).toBe(2)
    expect(active).toBe(0)
  })

  test("independent native model graphs share the replacement gate for streams and generation", async () => {
    const admission = new ModelAdmission(() => 1)
    const controls = Array.from({ length: 2 }, latch)
    let started = 0
    const response = () => new Response([
      { type: "response.output_item.added", item: { type: "message", id: "message" } },
      { type: "response.output_text.delta", item_id: "message", delta: "Hello" },
      { type: "response.completed", response: { id: "response" } },
    ].map(event => `data: ${JSON.stringify(event)}\n\n`).join(""), { headers: { "content-type": "text/event-stream" } })
    const fake = Object.assign(async () => { const index = started++; await controls[index]!.promise; return response() }, { preconnect: fetch.preconnect })
    const request = ai.request({ model: configure({ apiKey: "fixture", baseURL: "https://model.invalid/v1" }).responses("fixture"), prompt: "Hello" })
    const graph = LayerNode.group([llmClient, httpClient])
    const run = (stream: boolean) => Effect.runPromise(Effect.scoped(
      (stream ? LLMClient.stream(request).pipe(Stream.runCollect, Effect.asVoid) : LLMClient.generate(request).pipe(Effect.asVoid)).pipe(
        Effect.provide(LayerNode.compile(graph, { replacements: [makeModelConcurrencyNode(admission)] })),
        Effect.provideService(FetchHttpClient.Fetch, fake),
      ),
    ))
    const first = run(true)
    const second = run(false)
    await until(() => started === 1)
    await Bun.sleep(50)
    expect(started).toBe(1)
    controls[0]!.resolve()
    await until(() => started === 2)
    controls[1]!.resolve()
    await Promise.all([first, second])
  })

  test("raising a limit wakes waiting jobs; lowering it does not interrupt active jobs", async () => {
    let limit = 1
    const admission = new ModelAdmission(() => limit)
    const releases = Array.from({ length: 3 }, latch)
    let started = 0
    const run = () => Effect.runPromise(Effect.scoped(Effect.flatMap(admission.permit(), () => Effect.promise(async () => { const at = started++; await releases[at]!.promise }))))
    const jobs = [run(), run(), run()]
    await until(() => started === 1)
    limit = 2
    await until(() => started === 2)
    limit = 1
    releases[0]!.resolve()
    await Bun.sleep(300)
    expect(started).toBe(2)
    releases[1]!.resolve()
    await until(() => started === 3)
    releases[2]!.resolve()
    await Promise.all(jobs)
  })

  test("interrupting an active stream releases its permit for the next request", async () => {
    const admission = new ModelAdmission(() => 1)
    let active = 0
    const observer = { start: () => { active++ }, end: () => { active-- } }
    const fake = {
      stream: () => Stream.fromEffect(Effect.never),
      generate: () => Effect.succeed("result"),
      compact: () => Effect.die("unused"),
    } as unknown as Context.Service.Shape<typeof LLMClient.Service>
    const client = limitModelClient(fake, observer, admission)
    const controller = new AbortController()
    const first = Effect.runPromise(Stream.runCollect(client.stream({} as never)), { signal: controller.signal }).catch(() => undefined)
    await until(() => active === 1)
    const second = Effect.runPromise(client.generate({} as never))
    controller.abort()
    await Promise.all([first, second])
    expect(active).toBe(0)
  })

  test("cancelled queued work is removed and failed work releases its slot", async () => {
    const admission = new ModelAdmission(() => 1)
    const first = latch()
    const running = Effect.runPromise(Effect.scoped(Effect.flatMap(admission.permit(), () => Effect.promise(() => first.promise))))
    const controller = new AbortController()
    let cancelledStarted = false
    const cancelled = Effect.runPromise(Effect.scoped(Effect.flatMap(admission.permit(), () => Effect.sync(() => { cancelledStarted = true }))), { signal: controller.signal }).catch(() => undefined)
    controller.abort()
    await cancelled
    let thirdStarted = false
    const third = Effect.runPromise(Effect.scoped(Effect.flatMap(admission.permit(), () => Effect.sync(() => { thirdStarted = true }))))
    first.resolve()
    await Promise.all([running, third])
    expect(cancelledStarted).toBe(false)
    expect(thirdStarted).toBe(true)
    await Effect.runPromise(Effect.scoped(Effect.flatMap(admission.permit(), () => Effect.fail("expected")))).catch(() => undefined)
    await Effect.runPromise(Effect.scoped(Effect.flatMap(admission.permit(), () => Effect.void)))
  })

  test("one scoped cooldown covers streamed, generated and compact requests while other connections progress", async () => {
    const admission = new ModelAdmission(() => 3)
    const raw = () => configure({ apiKey: "fixture", baseURL: "https://provider.invalid/v1" }).responses("fixture")
    const request = (id: string) => ai.request({ model: withCapacityScope(raw(), "provider", { type: "credential", id }), prompt: "Hello" })
    const starts: { operation: string; at: number }[] = []
    let refuse = true
    let active = 0
    const fake = {
      stream: () => {
        if (refuse) { refuse = false; return Stream.fail(new AIError({ reason: new RateLimitError({ message: "Temporary capacity", retryAfterMs: 30 }) })) }
        return Stream.fromEffect(Effect.sync(() => { starts.push({ operation: "stream", at: Date.now() }); return "streamed" }))
      },
      generate: () => Effect.sync(() => { starts.push({ operation: "generate", at: Date.now() }); return "generated" }),
      compact: () => Effect.sync(() => { starts.push({ operation: "compact", at: Date.now() }); return "compacted" }),
    } as unknown as Context.Service.Shape<typeof LLMClient.Service>
    const client = limitModelClient(fake, { start: () => { active++ }, end: () => { active-- } }, admission)
    await Effect.runPromise(Stream.runCollect(client.stream(request("a")))).catch(() => undefined)
    expect(active).toBe(0)
    expect(starts).toEqual([])
    const queued = [
      Effect.runPromise(Stream.runCollect(client.stream(request("a")))),
      Effect.runPromise(client.generate(request("a"))),
      Effect.runPromise(client.compact(request("a") as never)),
    ]
    await Effect.runPromise(client.generate(request("b")))
    expect(starts.map(row => row.operation)).toEqual(["generate"])
    expect(active).toBe(0)
    await Promise.all(queued)
    expect(starts.slice(1).map(row => row.operation)).toEqual(["stream", "generate", "compact"])
    for (let index = 2; index < starts.length; index++) expect(starts[index]!.at - starts[index - 1]!.at).toBeGreaterThanOrEqual(200)
    expect(active).toBe(0)
  })

  test("fresh and retry work share scoped FIFO and full concurrency returns after one staggered wave", async () => {
    const admission = new ModelAdmission(() => 3)
    admission.capacity("a", new AIError({ reason: new RateLimitError({ message: "Temporary capacity", retryAfterMs: 25 }) }))
    const order: { name: string; at: number }[] = []
    const run = (name: string, scope = "a") => Effect.runPromise(Effect.scoped(Effect.flatMap(admission.permit(undefined, scope), () => Effect.sync(() => { order.push({ name, at: Date.now() }) }))))
    const waiting = [run("older fresh"), run("retry"), run("newer fresh")]
    await run("unrelated", "b")
    expect(order.map(row => row.name)).toEqual(["unrelated"])
    await Promise.all(waiting)
    expect(order.slice(1).map(row => row.name)).toEqual(["older fresh", "retry", "newer fresh"])
    for (let index = 2; index < order.length; index++) expect(order[index]!.at - order[index - 1]!.at).toBeGreaterThanOrEqual(200)
    let running = 0
    const release = latch()
    const burst = Array.from({ length: 3 }, () => Effect.runPromise(Effect.scoped(Effect.flatMap(admission.permit(undefined, "a"), () => Effect.promise(async () => { running++; await release.promise })))) )
    await until(() => running === 3)
    release.resolve()
    await Promise.all(burst)
  })

  for (const occupied of [false, true]) test(`recovery wave begins on actual admission after ${occupied ? "another scope occupied the slots" : "native backoff outlasted the hint"}`, async () => {
    const admission = new ModelAdmission(() => 2)
    const release = latch()
    let active = 0
    const busy = occupied ? Array.from({ length: 2 }, () => Effect.runPromise(Effect.scoped(Effect.flatMap(admission.permit(undefined, "b"), () => Effect.promise(async () => { active++; await release.promise }))))) : []
    await until(() => active === busy.length)
    admission.capacity("a", new AIError({ reason: new RateLimitError({ message: "Temporary capacity", retryAfterMs: 10 }) }))
    const order: { name: string; at: number }[] = []
    const run = (name: string, scope = "a") => Effect.runPromise(Effect.scoped(Effect.flatMap(admission.permit(undefined, scope), () => Effect.sync(() => { order.push({ name, at: Date.now() }) }))))
    let waiting = occupied ? [run("older fresh"), run("retry")] : []
    // Exceed the old time-only expiry before either recovering operation can
    // acquire a slot. Native retry backoff and other routes can both do this.
    await Bun.sleep(600)
    if (!occupied) waiting = [run("older fresh"), run("retry")]
    release.resolve()
    await until(() => order.some(row => row.name === "older fresh"))
    await run("unrelated", "b")
    await Promise.all([...waiting, ...busy])
    expect(order.map(row => row.name)).toEqual(["older fresh", "unrelated", "retry"])
    expect(order[2]!.at - order[0]!.at).toBeGreaterThanOrEqual(200)
  })

  test("a typed no-hint refusal uses the native initial gap without holding an active permit", async () => {
    const admission = new ModelAdmission(() => 1)
    expect(admission.capacity("a", new AIError({ reason: new RateLimitError({ message: "Temporary capacity" }) }))).toBe(true)
    const controller = new AbortController()
    let delayedStarted = false
    const delayed = Effect.runPromise(Effect.scoped(Effect.flatMap(admission.permit(undefined, "a"), () => Effect.sync(() => { delayedStarted = true }))), { signal: controller.signal }).catch(() => undefined)
    let unrelatedStarted = false
    await Effect.runPromise(Effect.scoped(Effect.flatMap(admission.permit(undefined, "b"), () => Effect.sync(() => { unrelatedStarted = true }))))
    await Bun.sleep(40)
    expect(delayedStarted).toBe(false)
    expect(unrelatedStarted).toBe(true)
    controller.abort()
    await delayed
  })

  test("cancelling queued and first recovery work cannot strand its scope or unrelated work", async () => {
    const admission = new ModelAdmission(() => 2)
    admission.capacity("a", new AIError({ reason: new RateLimitError({ message: "Temporary capacity", retryAfterMs: 20 }) }))
    const probe = new AbortController(), removed = new AbortController()
    let probeStarted = false, removedStarted = false, laterStarted = false
    const first = Effect.runPromise(Effect.scoped(Effect.flatMap(admission.permit(undefined, "a"), () => Effect.andThen(Effect.sync(() => { probeStarted = true }), Effect.never))), { signal: probe.signal }).catch(() => undefined)
    const cancelled = Effect.runPromise(Effect.scoped(Effect.flatMap(admission.permit(undefined, "a"), () => Effect.sync(() => { removedStarted = true }))), { signal: removed.signal }).catch(() => undefined)
    removed.abort()
    await cancelled
    await until(() => probeStarted)
    await Effect.runPromise(Effect.scoped(Effect.flatMap(admission.permit(undefined, "b"), () => Effect.void)))
    probe.abort()
    await first
    await Effect.runPromise(Effect.scoped(Effect.flatMap(admission.permit(undefined, "a"), () => Effect.sync(() => { laterStarted = true }))))
    expect(removedStarted).toBe(false)
    expect(laterStarted).toBe(true)
  })

  test("invalid delays, unknown errors, terminal quota/auth and retry=false cannot poison admission", async () => {
    const errors: unknown[] = [
      ...[-1, Infinity, NaN].map(retryAfterMs => {
        const error = new AIError({ reason: new RateLimitError({ message: "Invalid hint" }) })
        // Constructors reject these values; also defend against subsequently
        // corrupted exception metadata from an untyped client implementation.
        Object.defineProperty(error.reason, "retryAfterMs", { value: retryAfterMs })
        return error
      }),
      { _tag: "RateLimit", retryAfterMs: 10000 },
      new AIError({ reason: new QuotaExceededError({ message: "Hard quota", retryAfterMs: 10000 }) }),
      new AIError({ reason: new AuthenticationError({ message: "Sign in required" }) }),
      new AIError({ reason: new RateLimitError({ message: "Do not retry", retryAfterMs: 10000, http: new HttpContext({ url: "https://provider.invalid", status: 429, headers: { "x-should-retry": "false" } }) }) }),
    ]
    for (const error of errors) {
      const admission = new ModelAdmission(() => 1)
      expect(admission.capacity("a", error)).toBe(false)
      const controller = new AbortController()
      const accepted = Effect.runPromise(Effect.scoped(Effect.flatMap(admission.permit(undefined, "a"), () => Effect.succeed(true))), { signal: controller.signal }).catch(() => false)
      expect(await Promise.race([accepted, Bun.sleep(40).then(() => false)])).toBe(true)
      controller.abort()
    }
  })
})
