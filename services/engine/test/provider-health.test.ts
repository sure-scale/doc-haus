import { expect, test } from "bun:test"
import { AIError, AuthenticationError, LanguageModel, LLMClient, QuotaExceededError, TransportError } from "@opencode/ai"
import { OpenAIChat } from "@opencode/ai/protocols"
import { OpenAICompatible } from "@opencode/ai/providers/openai-compatible"
import { RequestExecutor } from "@opencode/ai/route"
import { TestLLM } from "@opencode/ai/testing"
import { Model } from "@opencode/core/model"
import { ModelResolver } from "@opencode/core/model-resolver"
import { Cause, Effect, Layer } from "effect"
import { HttpClient, HttpClientRequest, HttpClientResponse } from "effect/unstable/http"
import { probeProvider, providerTestCache, providerTestFailure, providerTestInput } from "../src/provider-health"

test("native stateless probe preserves provider output limits and sends only its fixed prompt", async () => {
  for (const providerID of ["synthetic", "openai"]) {
    const ref = Model.Ref.parse(`${providerID}/note`)
    const selected = Model.Info.default(ref.providerID, ref.id)
    const runtime = LanguageModel.make({ id: "note", provider: providerID, route: OpenAIChat.route })
    await Effect.runPromise(Effect.gen(function* () {
      expect(yield* probeProvider(ref)).toEqual({ ok: true })
      const [request] = yield* TestLLM.Test.use(test => test.requests())
      expect(request.generation?.maxTokens).toBeUndefined()
      expect(request.tools).toEqual([])
      expect(request.messages).toHaveLength(1)
      expect(JSON.stringify(request.messages)).toContain("Reply with OK.")
      expect(request.http?.headers?.["x-opencode-session"]).toMatch(/^ses_/)
    }).pipe(Effect.provide(Layer.mergeAll(
      Layer.mock(Model.Service, { get: () => Effect.succeed(selected) }),
      Layer.mock(ModelResolver.Service, { resolve: () => Effect.succeed({ model: runtime, ref, capabilities: selected.capabilities, cost: selected.cost, limit: selected.limit }) }),
      TestLLM.testLayer({ fallback: TestLLM.text("OK", "health") }),
    ))))
  }
})

test("provider failures distinguish expired authentication from account allowance and transport", () => {
  const expired = new AIError({ reason: new AuthenticationError({ message: "Synthetic API key has expired: synthetic-private-value" }) })
  const rejected = providerTestFailure(Cause.fail(expired))
  expect(rejected).toMatchObject({ ok: false, kind: "auth", expired: true })
  expect(rejected).toHaveProperty("error", "The provider API key has expired. Replace it in Settings; signing in again will not renew an expired key.")
  expect(JSON.stringify(rejected)).not.toContain("synthetic-private-value")
  expect(providerTestFailure(Cause.fail(new AIError({ reason: new AuthenticationError({ message: "Invalid API key" }) })))).toMatchObject({ kind: "auth" })
  expect(providerTestFailure(Cause.fail(new AIError({ reason: new AuthenticationError({ message: "OAuth access token expired" }) })))).toMatchObject({ kind: "auth" })
  expect(providerTestFailure(Cause.fail(new AIError({ reason: new AuthenticationError({ message: "OAuth access token expired" }) })))).not.toHaveProperty("expired")
  expect(providerTestFailure(Cause.fail(new AIError({ reason: new AuthenticationError({ message: "Authentication rejected", body: JSON.stringify({ error: { message: "API key has expired" } }) }) })))).toMatchObject({ kind: "auth", expired: true })
  const quota = new AIError({ reason: new QuotaExceededError({ message: "Account credits expired; verify authentication allowance" }) })
  expect(providerTestFailure(Cause.fail(quota))).toMatchObject({ kind: "unavailable" })
  const network = new AIError({ reason: new TransportError({ message: "Authentication server timeout", transport: "http", operation: "request" }) })
  expect(providerTestFailure(Cause.fail(network))).toMatchObject({ kind: "unavailable" })
  expect(providerTestFailure(Cause.interrupt())).toMatchObject({ kind: "unavailable" })
  expect(providerTestFailure(Cause.fail({ _tag: "Generate.ModelSelectionError" }))).toMatchObject({ kind: "model" })
})

test("native HTTP access restrictions are model failures, while identical 401 and generic 403 remain authentication failures", async () => {
  const ref = Model.Ref.parse("synthetic/note")
  const selected = Model.Info.default(ref.providerID, ref.id)
  const runtime = OpenAICompatible.configure({ provider: "synthetic", baseURL: "https://provider.invalid/v1", apiKey: "synthetic-test-key" }).model("note")
  const calls: string[] = []
  async function probe(status: number, body: string) {
    const http = Layer.succeed(HttpClient.HttpClient, HttpClient.make(request => Effect.gen(function* () {
      const web = yield* HttpClientRequest.toWeb(request).pipe(Effect.orDie)
      calls.push(web.url)
      expect(web.method).toBe("POST")
      expect(web.headers.get("x-opencode-session")).toMatch(/^ses_/)
      const input = JSON.parse(yield* Effect.promise(() => web.text()))
      expect(input.messages).toEqual([{ role: "user", content: "Reply with OK." }])
      return HttpClientResponse.fromWeb(request, new Response(body, { status, headers: { "content-type": status === 200 ? "text/event-stream" : "application/json" } }))
    })))
    const executor = RequestExecutor.layer.pipe(Layer.provide(http))
    return Effect.runPromise(probeProvider(ref).pipe(Effect.provide(Layer.mergeAll(
      Layer.mock(Model.Service, { get: () => Effect.succeed(selected) }),
      Layer.mock(ModelResolver.Service, { resolve: () => Effect.succeed({ model: runtime, ref, capabilities: selected.capabilities, cost: selected.cost, limit: selected.limit }) }),
      LLMClient.layer.pipe(Layer.provide(executor)),
    ))))
  }
  // Establish a passing native route before changing only the HTTP response.
  const success = `data: ${JSON.stringify({ id: "synthetic", object: "chat.completion.chunk", created: 1, model: "note", choices: [{ index: 0, delta: { role: "assistant", content: "OK" }, finish_reason: null }] })}\n\ndata: ${JSON.stringify({ choices: [{ index: 0, delta: {}, finish_reason: "stop" }] })}\n\ndata: [DONE]\n\n`
  expect(await probe(200, success)).toEqual({ ok: true })
  for (const [type, restriction] of [["FreeTierError", "free-tier"], ["RegionError", "region"]] as const) {
    const body = JSON.stringify({ error: { type, message: "synthetic-private-diagnostic" } })
    const failure = await probe(403, body)
    expect(failure).toMatchObject({ ok: false, kind: "model", restriction })
    expect(JSON.stringify(failure)).not.toContain("synthetic-private-diagnostic")
    expect(await probe(401, body)).toMatchObject({ ok: false, kind: "auth" })
  }
  expect(await probe(403, JSON.stringify({ error: { type: "InvalidApiKeyError", message: "Invalid key" } }))).toMatchObject({ kind: "auth" })
  expect(calls).toHaveLength(6)
  expect(calls.every(url => url === "https://provider.invalid/v1/chat/completions")).toBe(true)
})

test("access classification accepts only bounded JSON codes at HTTP 403, never message hints", () => {
  const failure = (status: number, body: string, message = "Provider rejected request") => providerTestFailure(Cause.fail(new AIError({
    reason: new AuthenticationError({ message, body, http: { url: "https://provider.invalid", status, headers: {} } }),
  })))
  for (const body of ["FreeTierError", "{", JSON.stringify({ error: { message: "FreeTierError" } }), JSON.stringify({ error: { type: "OtherError" } }), JSON.stringify({ error: { type: "FreeTierError" }, padding: "x".repeat(8192) })]) {
    expect(failure(403, body, "FreeTierError: API key expired")).toMatchObject({ kind: "auth", expired: true })
    expect(failure(403, body)).not.toHaveProperty("restriction")
  }
  expect(failure(401, JSON.stringify({ error: { type: "FreeTierError" } }))).toMatchObject({ kind: "auth" })
})

test("startup probes share pending and recent results, while retry and changed connection keys recheck", async () => {
  let now = 100
  let calls = 0
  const cache = providerTestCache(() => now)
  let release!: () => void
  const waiting = new Promise<void>(resolve => { release = resolve })
  const check = async () => { calls++; await waiting; return { ok: true } as const }
  const first = cache("model-credential-one", false, check)
  const second = cache("model-credential-one", false, check)
  const forcedWhilePending = cache("model-credential-one", true, check)
  expect(first).toBe(second)
  expect(first).toBe(forcedWhilePending)
  release()
  await first
  await cache("model-credential-one", false, check)
  expect(calls).toBe(1)
  await cache("model-credential-one", true, check)
  expect(calls).toBe(2)
  await cache("model-credential-two", false, check)
  expect(calls).toBe(3)
  now += 60_001
  await cache("model-credential-one", false, check)
  expect(calls).toBe(4)
})

test("the fixed probe contract rejects extra generation options and malformed retry flags", () => {
  expect(providerTestInput({ providerID: "synthetic", modelID: "note" })).toBe(true)
  expect(providerTestInput({ providerID: "synthetic", modelID: "note", force: true })).toBe(true)
  expect(providerTestInput({ providerID: "synthetic", modelID: "note", prompt: "Caller text" })).toBe(false)
  expect(providerTestInput({ providerID: "synthetic", modelID: "note", force: "true" })).toBe(false)
  expect(providerTestInput({ providerID: "synthetic", modelID: "note#variant" })).toBe(false)
})
