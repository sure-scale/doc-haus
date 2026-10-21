import { expect, test } from "bun:test"
import { AIError, AuthenticationError, LanguageModel, QuotaExceededError, TransportError } from "@opencode/ai"
import { OpenAIChat } from "@opencode/ai/protocols"
import { TestLLM } from "@opencode/ai/testing"
import { Model } from "@opencode/core/model"
import { ModelResolver } from "@opencode/core/model-resolver"
import { Cause, Effect, Layer } from "effect"
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
