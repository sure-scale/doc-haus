import { expect, test } from "bun:test"
import { LLM, LLMClient } from "@opencode/ai"
import { configure } from "@opencode/ai/providers/openai"
import { llmClient } from "@opencode/core/effect/app-node-platform"
import { httpClient } from "@opencode/util/effect/app-node-platform"
import { LayerNode } from "@opencode/util/effect/layer-node"
import { Effect } from "effect"
import { FetchHttpClient, HttpClient } from "effect/unstable/http"
import { modelHttp } from "./model-http"

type FetchOptions = RequestInit & { timeout?: number | boolean }

const graph = LayerNode.group([llmClient, httpClient])
const request = LLM.request({
  model: configure({ apiKey: "fixture", baseURL: "https://model.invalid/v1" }).responses("fixture"),
  prompt: "Hello",
})
const response = () => new Response([
  { type: "response.output_item.added", item: { type: "message", id: "msg_local" } },
  { type: "response.output_text.delta", item_id: "msg_local", delta: "Hello" },
  { type: "response.completed", response: { id: "resp_local" } },
].map(event => `data: ${JSON.stringify(event)}\n\n`).join(""), {
  headers: { "content-type": "text/event-stream" },
})

for (const patched of [false, true]) {
  test(`${patched ? "engine" : "unmodified"} native graph fetch options stay scoped to model execution`, async () => {
    const calls: Array<{ url: string; options?: FetchOptions }> = []
    const fake = Object.assign(async (input: string | URL | Request, options?: FetchOptions) => {
      calls.push({ url: String(input), options })
      return response()
    }, { preconnect: fetch.preconnect })
    await Effect.runPromise(Effect.scoped(Effect.gen(function* () {
      const result = yield* LLMClient.generate(request)
      expect(result.text).toBe("Hello")
      const http = yield* HttpClient.HttpClient
      yield* http.get("https://shared.invalid/catalog")
    }).pipe(
      Effect.provide(LayerNode.compile(graph, { replacements: patched ? [modelHttp] : [] })),
      Effect.provideService(FetchHttpClient.Fetch, fake),
      Effect.provideService(FetchHttpClient.RequestInit, { redirect: "error", cache: "no-store" }),
    )))
    expect(calls.map(call => call.url)).toEqual(["https://model.invalid/v1/responses", "https://shared.invalid/catalog"])
    expect(calls[0].options?.timeout).toBe(patched ? false : undefined)
    expect(calls[0].options).toMatchObject({ redirect: "error", cache: "no-store", method: "POST" })
    expect(calls[0].options?.signal).toBeInstanceOf(AbortSignal)
    expect(calls[1].options?.timeout).toBeUndefined()
    expect(calls[1].options?.redirect).toBe("error")
  })
}

test("caller abort still cancels a quiet native model fetch", async () => {
  const started = Promise.withResolvers<void>()
  const controller = new AbortController()
  let options: FetchOptions | undefined
  const fake = Object.assign((_input: string | URL | Request, init?: FetchOptions) => {
    options = init
    return new Promise<Response>((_resolve, reject) => {
      init?.signal?.addEventListener("abort", () => reject(init.signal?.reason), { once: true })
      started.resolve()
    })
  }, { preconnect: fetch.preconnect })
  const pending = Effect.runPromise(Effect.scoped(LLMClient.generate(request).pipe(
    Effect.provide(LayerNode.compile(graph, { replacements: [modelHttp] })),
    Effect.provideService(FetchHttpClient.Fetch, fake),
  )), { signal: controller.signal })
  await started.promise
  expect(options?.timeout).toBe(false)
  expect(options?.signal?.aborted).toBe(false)
  controller.abort()
  await expect(pending).rejects.toThrow()
  expect(options?.signal?.aborted).toBe(true)
}, 1_000)
