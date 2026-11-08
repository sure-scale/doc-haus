import { expect, test } from "bun:test"
import { ai, LanguageModel } from "@opencode/ai/effect"
import { configure } from "@opencode/ai/providers/openai"
import { compileRequest } from "@opencode/ai/testing"
import { MODEL_CAPACITY_SCOPE, ModelResolver, withCapacityScope } from "@opencode/core/model-resolver"
import { Model } from "@opencode/core/model"
import { Provider } from "@opencode/core/provider"
import { Integration } from "@opencode/core/integration"
import { Credential } from "@opencode/core/credential"
import { AISDK } from "@opencode/core/aisdk"
import { Npm } from "@opencode/util/npm"
import { Effect, Layer, Option } from "effect"
import { Headers } from "effect/http"

const model = () => configure({ apiKey: "synthetic-api-key", baseURL: "https://provider.invalid/v1" }).responses("synthetic-model")
const scope = (selected: LanguageModel) => Object.getOwnPropertyDescriptor(selected.route.auth, MODEL_CAPACITY_SCOPE)?.value

test("the actual resolver carries the active environment connection across separately constructed runtime models", async () => {
  const providerID = Provider.ID.make("synthetic-provider")
  const selected = Model.Info.make({ ...Model.Info.default(providerID, Model.ID.make("fixture")),
    package: "@opencode/ai/providers/openai", canonical: Provider.ID.make("openai"), settings: { baseURL: "https://provider.invalid/v1" },
  })
  const provider = Provider.Info.make({ ...Provider.Info.empty(providerID), activation: "enabled", package: selected.package!, settings: selected.settings })
  let name = "SYNTHETIC_PROVIDER_KEY"
  const unused = () => Effect.die("Unused synthetic service")
  const integrations = Layer.mock(Integration.Service, {
    revision: () => 0,
    connection: {
      active: () => Effect.succeed({ type: "env" as const, name }),
      resolve: () => Effect.succeed(Credential.Key.make({ type: "key", key: "synthetic-api-key" })),
      key: unused, external: unused, activate: unused, update: unused, remove: unused, status: unused,
    },
    oauth: { connect: unused, status: unused, complete: unused, cancel: unused },
    command: { connect: unused, status: unused, cancel: unused },
  })
  const layer = ModelResolver.layer.pipe(Layer.provide(Layer.mergeAll(
    Layer.mock(Provider.Service, { get: () => Effect.succeed(provider) }), Layer.mock(Model.Service, { get: () => Effect.succeed(selected) }),
    integrations, Layer.mock(Npm.Service, { add: unused, which: unused }), Layer.mock(AISDK.Service, { model: unused, hook: { sdk: unused, language: unused } }),
  )))
  await Effect.runPromise(Effect.gen(function* () {
    const resolver = yield* ModelResolver.Service
    const first = yield* resolver.resolveModel(selected)
    const second = yield* resolver.resolveModel(selected)
    expect(first.model.route.auth).not.toBe(second.model.route.auth)
    expect(scope(first.model)).toBe(scope(second.model))
    expect(scope(first.model)).toContain("env:SYNTHETIC_PROVIDER_KEY")
    name = "ANOTHER_SYNTHETIC_PROVIDER_KEY"
    const changed = yield* resolver.resolveModel(selected)
    expect(scope(changed.model)).not.toBe(scope(first.model))
  }).pipe(Effect.provide(layer)))
})

test("independent model resolutions share only the same selected provider and nonsecret access identity", () => {
  const first = withCapacityScope(model(), "provider-a", { type: "credential", id: "saved-a" })
  const second = withCapacityScope(model(), "provider-a", { type: "credential", id: "saved-a" })
  expect(first.route.auth).not.toBe(second.route.auth)
  expect(scope(first)).toBe(scope(second))
  expect(scope(first)).not.toBe(scope(withCapacityScope(model(), "provider-a", { type: "credential", id: "saved-b" })))
  expect(scope(first)).not.toBe(scope(withCapacityScope(model(), "provider-b", { type: "credential", id: "saved-a" })))
  const environment = withCapacityScope(model(), "provider-a", { type: "env", name: "SYNTHETIC_PROVIDER_KEY" })
  expect(scope(environment)).toContain("env:SYNTHETIC_PROVIDER_KEY")
  expect(scope(environment)).not.toContain("synthetic-api-key")
})

test("model and endpoint cloning preserve the scope without mutating shared original auth", () => {
  const original = model()
  const selected = withCapacityScope(original, "provider-a", { type: "env", name: "SYNTHETIC_PROVIDER_KEY" })
  expect(scope(original)).toBeUndefined()
  expect(selected.route.auth).not.toBe(original.route.auth)
  expect(Object.getPrototypeOf(selected.route.auth)).toBe(Object.getPrototypeOf(original.route.auth))
  const modelClone = LanguageModel.update(selected, { defaults: { generation: { maxTokens: 32 } } })
  const routeClone = LanguageModel.update(modelClone, { route: modelClone.route.with({ endpoint: { baseURL: "https://proxy.invalid/v1" } }) })
  expect(scope(routeClone)).toBe(scope(selected))
  const noAccess = model()
  expect(withCapacityScope(noAccess, "provider-a", undefined)).toBe(noAccess)
  expect(scope(noAccess)).toBeUndefined()
})

test("the scope is absent from prompt, provider JSON, headers and ordinary serialization", async () => {
  const original = model()
  const selected = withCapacityScope(original, "provider-a", { type: "credential", id: "private-connection-id" })
  const request = ai.request({ model: selected, prompt: "Describe the sample." })
  expect(Object.getOwnPropertyDescriptor(selected.route.auth, MODEL_CAPACITY_SCOPE)?.enumerable).toBe(false)
  expect(JSON.stringify(request)).not.toContain("private-connection-id")
  expect(JSON.stringify(selected.route.auth)).toBe(JSON.stringify(original.route.auth))
  const prepared = await Effect.runPromise(compileRequest(request))
  expect(JSON.stringify(prepared.body)).not.toContain("private-connection-id")
  const headers = await Effect.runPromise(selected.route.auth.apply({ request, method: "POST", url: "https://provider.invalid/v1/responses", body: "{}", headers: Headers.empty }))
  expect(JSON.stringify(headers)).not.toContain("private-connection-id")
  expect(Headers.get(headers, "authorization")).toEqual(Option.some("Bearer synthetic-api-key"))
  expect(JSON.stringify(prepared.body)).toContain("Describe the sample.")
})
