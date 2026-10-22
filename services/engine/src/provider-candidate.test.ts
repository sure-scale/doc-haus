import { expect, test } from "bun:test"
import { mkdtempSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { LLMClient } from "@opencode/ai"
import { RequestExecutor } from "@opencode/ai/route"
import { Integration } from "@opencode/core/integration"
import { Context, Effect, Layer, Logger } from "effect"
import { FetchHttpClient } from "effect/unstable/http"
import { providerCandidates, candidateInput, candidateSettings } from "./provider-candidate"
import { createProviderSetupSession } from "./provider-setup-session"
import { credentialReuse } from "./credential-reuse"
import { mergeSettings } from "./effective-settings"
import { probeProvider } from "./provider-health"
import { Model } from "@opencode/core/model"
import { Provider } from "@opencode/core/provider"

const quiet = Logger.layer([])
test("candidate fields reject executable settings and preserve policy and current model pair", () => {
  const input = {
    providerID: "synthetic-local",
    connection: { type: "host" },
    local: { name: "Synthetic", baseURL: "http://localhost:1234/v1", modelIDs: [] },
  }
  expect(candidateInput(input)).toBe(true)
  expect(candidateInput({ ...input, prompt: "User matter text" })).toBe(false)
  expect(candidateInput({ ...input, options: { apiKey: "synthetic" } })).toBe(false)
  const provider = {
    ...Provider.Info.empty(Provider.ID.make(input.providerID)),
    package: "@opencode/ai/providers/openai-compatible",
  }
  const policy = { model: "old/main", small_model: "old/fast", disabled_providers: [input.providerID] }
  expect(() => candidateSettings(policy, input as never, provider)).toThrow("disabled")
  const allowed = { ...policy, disabled_providers: [] }
  const settings = candidateSettings(allowed, input as never, provider)
  expect(settings.model).toBe("old/main")
  expect(settings.small_model).toBe("old/fast")
  expect("provider" in allowed).toBe(false)
})

test("native draft tests exact connection and both models before deferred save; rollback restores the old credential", async () => {
  const legal = mkdtempSync(path.join(tmpdir(), "dochaus-candidate-test-"))
  const calls: { authorization: string | null; model: string; messages: unknown[]; tools?: unknown[] }[] = []
  const provider = Bun.serve({
    port: 0,
    async fetch(request) {
      const body = await request.json()
      calls.push({ authorization: request.headers.get("authorization"), ...body })
      if (
        !(body.model === "noauth" && request.headers.get("authorization") === null) &&
        request.headers.get("authorization") !== "Bearer synthetic-new"
      )
        return Response.json({ error: { message: "Synthetic rejection" } }, { status: 401 })
      if (body.stream !== false)
        return new Response(
          `data: ${JSON.stringify({ id: "synthetic-result", object: "chat.completion.chunk", created: 1, model: body.model, choices: [{ index: 0, delta: { role: "assistant", content: "OK" }, finish_reason: "stop" }] })}\n\ndata: [DONE]\n\n`,
          { headers: { "content-type": "text/event-stream" } },
        )
      return Response.json({
        id: "synthetic-result",
        object: "chat.completion",
        created: 1,
        model: body.model,
        choices: [{ index: 0, message: { role: "assistant", content: "OK" }, finish_reason: "stop" }],
      })
    },
  })
  writeFileSync(path.join(legal, "opencode.json"), "{}")
  try {
    await Effect.runPromise(
      Effect.scoped(
        Effect.gen(function* () {
          const integrationID = Integration.ID.make("synthetic-local")
          const baseline = yield* createProviderSetupSession({
            legal,
            settings: {},
            credential: { integrationID, label: "Old connection", value: { type: "key", key: "synthetic-old" } },
          })
          const clientContext = yield* Layer.build(
            LLMClient.layer.pipe(Layer.provide(RequestExecutor.layer.pipe(Layer.provide(FetchHttpClient.layer)))),
          )
          const old = yield* baseline.credentials.all()
          let holdCreate: Promise<void> | undefined
          let onCreate = () => {}
          let createCount = 0
          const installationCredentials = {
            ...baseline.credentials,
            create: (input: Parameters<typeof baseline.credentials.create>[0]) =>
              Effect.gen(function* () {
                createCount++
                onCreate()
                if (holdCreate) yield* Effect.promise(() => holdCreate!)
                return yield* baseline.credentials.create(input)
              }),
          }
          const policy = {
            provider: {
              [integrationID]: {
                options: {
                  apiKey: "synthetic-legacy-config-key",
                  authToken: "synthetic-legacy-auth",
                  accessToken: "synthetic-legacy-access",
                },
              },
            },
          }
          const manager = providerCandidates({
            legal,
            credentials: installationCredentials,
            client: Context.get(clientContext, LLMClient.Service),
            settings: () => policy,
            inspect: (effect) => Effect.runPromise(baseline.provide(effect)),
            reuse: credentialReuse({
              directory: path.join(legal, "absent"),
              targetDirectory: legal,
              create: async () => {
                throw new Error("Unexpected import")
              },
            }),
          })
          yield* Effect.promise(async () => {
            const create = (key: string) =>
              manager.handle("session", {
                providerID: integrationID,
                connection: { type: "key", key },
                local: { name: "Synthetic local", baseURL: `http://127.0.0.1:${provider.port}/v1`, modelIDs: [] },
              })
            const session = (await create("synthetic-new")) as { sessionID: string }
            try {
              const models = (await manager.handle("models", { ...session, modelIDs: ["primary", "fast"] })) as {
                data: { id: string }[]
                customModels: boolean
              }
              expect(models.customModels).toBe(true)
              expect(models.data.map((model) => model.id)).toEqual(["primary", "fast"])
              expect(await manager.handle("test", { ...session, modelID: "primary" })).toEqual({ ok: true })
              await expect(manager.handle("prepare", { ...session, primary: "primary", fast: "fast" })).rejects.toThrow(
                "both selected",
              )
              expect(await manager.handle("test", { ...session, modelID: "fast" })).toEqual({ ok: true })
              expect(await Effect.runPromise(baseline.credentials.all())).toEqual(old)
              const prepared = (await manager.handle("prepare", { ...session, primary: "primary", fast: "fast" })) as {
                patch: object
              }
              expect(prepared.patch).toMatchObject({
                model: "synthetic-local/primary",
                small_model: "synthetic-local/fast",
                provider: { "synthetic-local": { models: { primary: { name: "primary" }, fast: { name: "fast" } } } },
              })
              expect(JSON.stringify(prepared)).not.toContain("synthetic-new")
              expect(calls).toHaveLength(2)
              expect(
                calls.every(
                  (call) =>
                    call.authorization === "Bearer synthetic-new" &&
                    JSON.stringify(call.messages).includes("Reply with OK.") &&
                    !call.tools?.length,
                ),
              ).toBe(true)
              await manager.handle("save", { ...session, primary: "primary", fast: "fast" })
              const saved = await Effect.runPromise(baseline.credentials.list(integrationID))
              expect(saved).toHaveLength(2)
              const finalSettings = mergeSettings(policy, prepared.patch as Record<string, unknown>)
              expect((finalSettings.provider?.[integrationID]?.options as Record<string, unknown>).apiKey).toBe(
                "synthetic-legacy-config-key",
              )
              const activeSaved = saved.at(-1)!
              const reloaded = await Effect.runPromise(
                Effect.scoped(
                  Effect.gen(function* () {
                    const runtime = yield* createProviderSetupSession({
                      legal,
                      settings: finalSettings,
                      credential: { integrationID, label: activeSaved.label, value: activeSaved.value },
                    })
                    return yield* runtime.provide(
                      probeProvider(Model.Ref.parse(`${integrationID}/primary`)).pipe(
                        Effect.provideService(LLMClient.Service, Context.get(clientContext, LLMClient.Service)),
                      ),
                    )
                  }),
                ).pipe(Effect.provide(quiet)),
              )
              expect(reloaded).toEqual({ ok: true })
              expect(calls.at(-1)?.authorization).toBe("Bearer synthetic-new")
              await manager.handle("save", { ...session, primary: "primary", fast: "fast" })
              await manager.handle("rollback", session)
              expect(await Effect.runPromise(baseline.credentials.all())).toEqual(old)
              let releaseCreate = () => {}
              let enteredCreate = () => {}
              holdCreate = new Promise<void>((resolve) => {
                releaseCreate = resolve
              })
              const entered = new Promise<void>((resolve) => {
                enteredCreate = resolve
              })
              onCreate = enteredCreate
              const delayedSave = manager.handle("save", { ...session, primary: "primary", fast: "fast" })
              await entered
              const duplicateSave = manager.handle("save", { ...session, primary: "primary", fast: "fast" })
              let rollbackSettled = false
              const rollback = manager.handle("rollback", session).then((result) => {
                rollbackSettled = true
                return result
              })
              await Promise.resolve()
              expect(rollbackSettled).toBe(false)
              expect(await Effect.runPromise(baseline.credentials.all())).toEqual(old)
              releaseCreate()
              await Promise.all([delayedSave, duplicateSave, rollback])
              expect(createCount).toBe(2)
              expect(await Effect.runPromise(baseline.credentials.all())).toEqual(old)
              holdCreate = undefined
              await expect(
                manager.handle("session", {
                  providerID: integrationID,
                  connection: { type: "host" },
                  local: { name: "Conflicting host", baseURL: `http://127.0.0.1:${provider.port}/v1`, modelIDs: [] },
                }),
              ).rejects.toThrow("existing connection")
              const host = (await manager.handle("session", {
                providerID: "synthetic-no-auth",
                connection: { type: "host" },
                local: {
                  name: "No auth endpoint",
                  baseURL: `http://127.0.0.1:${provider.port}/v1`,
                  modelIDs: ["noauth"],
                },
              })) as { sessionID: string }
              try {
                expect(await manager.handle("test", { ...host, modelID: "noauth" })).toEqual({ ok: true })
                expect(await manager.handle("save", { ...host, primary: "noauth", fast: "noauth" })).toEqual({})
                expect(await Effect.runPromise(baseline.credentials.all())).toEqual(old)
              } finally {
                await manager.close(host.sessionID)
              }
              const rejected = (await create("synthetic-rejected")) as { sessionID: string }
              try {
                await manager.handle("models", { ...rejected, modelIDs: ["primary"] })
                expect(await manager.handle("test", { ...rejected, modelID: "primary" })).toMatchObject({
                  ok: false,
                  kind: "auth",
                })
                expect(await Effect.runPromise(baseline.credentials.all())).toEqual(old)
              } finally {
                await manager.close(rejected.sessionID)
              }
            } finally {
              await manager.close(session.sessionID)
            }
          })
        }),
      ).pipe(Effect.provide(quiet)),
    )
  } finally {
    provider.stop(true)
    rmSync(legal, { recursive: true, force: true })
  }
}, 60_000)

test("Vertex Anthropic draft project and location determine the same native endpoint that is saved", () => {
  const provider = {
    ...Provider.Info.empty(Provider.ID.make("google-vertex-anthropic")),
    package: "@opencode/ai/providers/google-vertex-messages",
  }
  for (const location of ["global", "europe-west1"]) {
    const settings = candidateSettings(
      {},
      { providerID: provider.id, connection: { type: "host" }, options: { project: " synthetic-project ", location } },
      provider,
    )
    expect(settings.provider?.[provider.id]?.options).toEqual({
      project: "synthetic-project",
      location,
      baseURL: `https://${location === "global" ? "" : location + "-"}aiplatform.googleapis.com/v1/projects/synthetic-project/locations/${location}/publishers/anthropic/models`,
    })
  }
})

test("reopening a saved compatible endpoint supports no key, replacement key and exact saved variants", async () => {
  const legal = mkdtempSync(path.join(tmpdir(), "dochaus-candidate-reopen-"))
  const calls: { authorization: string | null; variant: string | undefined }[] = []
  const endpoint = Bun.serve({
    port: 0,
    async fetch(request) {
      const body = await request.json()
      calls.push({
        authorization: request.headers.get("authorization"),
        variant: body.reasoning_effort,
      })
      return new Response(
        `data: ${JSON.stringify({ id: "synthetic", object: "chat.completion.chunk", created: 1, model: body.model, choices: [{ index: 0, delta: { content: "OK" }, finish_reason: "stop" }] })}\n\ndata: [DONE]\n\n`,
        { headers: { "content-type": "text/event-stream" } },
      )
    },
  })
  writeFileSync(path.join(legal, "opencode.json"), "{}")
  try {
    await Effect.runPromise(
      Effect.scoped(
        Effect.gen(function* () {
          const providerID = "synthetic-existing"
          const policy = {
            provider: {
              [providerID]: {
                npm: "@ai-sdk/openai-compatible",
                name: "Saved synthetic endpoint",
                options: { baseURL: "http://127.0.0.1:1/v1" },
                models: {
                  primary: {
                    name: "Primary",
                    variants: { precise: { reasoningEffort: "high" } },
                  },
                },
              },
            },
          }
          const baseline = yield* createProviderSetupSession({ legal, settings: policy })
          const client = yield* Layer.build(
            LLMClient.layer.pipe(Layer.provide(RequestExecutor.layer.pipe(Layer.provide(FetchHttpClient.layer)))),
          )
          const manager = providerCandidates({
            legal,
            credentials: baseline.credentials,
            client: Context.get(client, LLMClient.Service),
            settings: () => policy,
            inspect: (effect) => Effect.runPromise(baseline.provide(effect)),
            reuse: credentialReuse({
              directory: path.join(legal, "absent"),
              targetDirectory: legal,
              create: async () => {
                throw new Error("Unexpected import")
              },
            }),
          })
          yield* Effect.promise(async () => {
            for (const connection of [{ type: "host" }, { type: "key", key: "synthetic-replacement" }] as const) {
              const session = (await manager.handle("session", {
                providerID,
                options: { baseURL: `http://127.0.0.1:${endpoint.port}/v1` },
                connection,
              })) as { sessionID: string }
              try {
                const catalog = (await manager.handle("models", session)) as {
                  data: { id: string; variants: { id: string; name: string }[] }[]
                }
                expect(catalog.data.find((model) => model.id === "primary")?.variants).toContainEqual({
                  id: "precise",
                  name: "precise",
                })
                await expect(manager.handle("models", { ...session, modelIDs: ["primary#precise"] })).rejects.toThrow(
                  "model IDs",
                )
                expect(await manager.handle("test", { ...session, modelID: "primary#precise" })).toEqual({ ok: true })
                expect(calls.at(-1)).toEqual({
                  authorization: connection.type === "key" ? "Bearer synthetic-replacement" : null,
                  variant: "high",
                })
                expect(await manager.handle("test", { ...session, modelID: "primary#missing" })).toMatchObject({
                  ok: false,
                  kind: "model",
                })
                await expect(
                  manager.handle("prepare", { ...session, primary: "primary#missing", fast: "primary#precise" }),
                ).rejects.toThrow("both selected")
                const pair = { ...session, primary: "primary#precise", fast: "primary#precise" }
                expect(await manager.handle("prepare", pair)).toMatchObject({
                  patch: {
                    model: `${providerID}/primary#precise`,
                    small_model: `${providerID}/primary#precise`,
                    provider: { [providerID]: { options: { baseURL: `http://127.0.0.1:${endpoint.port}/v1` } } },
                  },
                })
                expect(await Effect.runPromise(baseline.credentials.all())).toHaveLength(0)
                await manager.handle("save", pair)
                expect(await Effect.runPromise(baseline.credentials.all())).toHaveLength(
                  connection.type === "key" ? 1 : 0,
                )
                await manager.handle("rollback", session)
                expect(await Effect.runPromise(baseline.credentials.all())).toHaveLength(0)
              } finally {
                await manager.close(session.sessionID)
              }
            }
          })
        }),
      ).pipe(Effect.provide(quiet)),
    )
  } finally {
    endpoint.stop(true)
    rmSync(legal, { recursive: true, force: true })
  }
}, 60_000)

test("clearing an optional cloud profile removes the override and preserves shipped fallback", () => {
  const provider = {
    ...Provider.Info.empty(Provider.ID.make("amazon-bedrock")),
    package: "@opencode/ai/providers/bedrock",
  }
  const policy = {
    provider: { "amazon-bedrock": { options: { region: "us-east-1", profile: "synthetic-old-profile" } } },
  }
  const input = { providerID: provider.id, connection: { type: "host" as const }, options: { profile: null } }
  expect(candidateInput(input)).toBe(true)
  expect(candidateSettings(policy, input, provider).provider?.[provider.id]?.options).toEqual({ region: "us-east-1" })
  expect(
    candidateSettings(policy, input, provider, {
      provider: { "amazon-bedrock": { options: { profile: "synthetic-shipped-profile" } } },
    }).provider?.[provider.id]?.options,
  ).toEqual({ region: "us-east-1", profile: "synthetic-shipped-profile" })
  expect(policy.provider["amazon-bedrock"].options.profile).toBe("synthetic-old-profile")
})
