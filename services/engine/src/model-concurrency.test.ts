import { describe, expect, test } from "bun:test"
import { Effect, Stream } from "effect"
import { LLM, LLMClient } from "@opencode/ai"
import { configure } from "@opencode/ai/providers/openai"
import { llmClient } from "@opencode/core/effect/app-node-platform"
import { httpClient } from "@opencode/util/effect/app-node-platform"
import { LayerNode } from "@opencode/util/effect/layer-node"
import { FetchHttpClient } from "effect/unstable/http"
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
    const request = LLM.request({ model: configure({ apiKey: "fixture", baseURL: "https://model.invalid/v1" }).responses("fixture"), prompt: "Hello" })
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
})
