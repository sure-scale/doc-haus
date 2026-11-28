import { afterEach, describe, expect, test } from "bun:test"
import { chmodSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { LLMClient } from "@opencode/ai/effect"
import { RequestExecutor } from "@opencode/ai/route"
import { Credential } from "@opencode/core/credential"
import { Integration } from "@opencode/core/integration"
import { Context, Effect, Layer, Logger } from "effect"
import { FetchHttpClient } from "effect/http"
import { providerCandidates, candidateInput, candidateSettings } from "./provider-candidate"
import { createProviderSetupSession, type SetupCredential } from "./provider-setup-session"
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
              // Adding a provider can save only its connection, leaving the
              // already-tested model roles untouched.
              const connectionOnly = (await manager.handle("prepare", { sessionID: session.sessionID })) as {
                patch: Record<string, unknown>
              }
              expect(Object.keys(connectionOnly.patch)).toEqual(["provider"])
              expect(connectionOnly.patch).not.toHaveProperty("model")
              expect(connectionOnly.patch).not.toHaveProperty("small_model")
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

test("a new setup replaces the least recently used abandoned draft, never the one being saved", async () => {
  const legal = mkdtempSync(path.join(tmpdir(), "dochaus-candidate-abandoned-"))
  const endpoint = Bun.serve({
    port: 0,
    async fetch(request) {
      const body = await request.json()
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
          const baseline = yield* createProviderSetupSession({ legal, settings: {} })
          const client = yield* Layer.build(
            LLMClient.layer.pipe(Layer.provide(RequestExecutor.layer.pipe(Layer.provide(FetchHttpClient.layer)))),
          )
          const manager = providerCandidates({
            legal,
            credentials: baseline.credentials,
            client: Context.get(client, LLMClient.Service),
            settings: () => ({}),
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
            const opened: string[] = []
            const start = async (name: string) => {
              const session = (await manager.handle("session", {
                providerID: `synthetic-${name}`,
                connection: { type: "host" },
                local: { name, baseURL: `http://127.0.0.1:${endpoint.port}/v1`, modelIDs: ["primary"] },
              })) as { sessionID: string }
              opened.push(session.sessionID)
              return session
            }
            const alive = (session: { sessionID: string }) =>
              manager.handle("models", session).then(
                () => true,
                (error: Error) => {
                  expect(error.message).toBe("Provider setup expired. Start again.")
                  return false
                },
              )
            try {
              // Pages reloaded or closed mid-setup leave drafts nobody discards.
              const abandoned = []
              for (const name of ["reload-1", "reload-2", "reload-3", "reload-4"]) abandoned.push(await start(name))
              const current = await start("current")
              expect(await manager.handle("test", { ...current, modelID: "primary" })).toEqual({ ok: true })
              expect(await manager.handle("save", { ...current, primary: "primary", fast: "primary" })).toEqual({})
              // A concurrent setup keeps its own draft but cannot save over one in progress.
              const second = await start("second")
              expect(await manager.handle("test", { ...second, modelID: "primary" })).toEqual({ ok: true })
              await expect(manager.handle("save", { ...second, primary: "primary", fast: "primary" })).rejects.toThrow(
                "Another provider setup is being saved",
              )
              // Two more setups make the draft being saved the least recently used.
              const later = [await start("third"), await start("fourth")]
              const last = await start("last")
              expect(await manager.handle("rollback", current)).toEqual({ ok: true })
              expect(await Promise.all(abandoned.map(alive))).toEqual([false, false, false, false])
              expect(await alive(second)).toBe(false)
              expect(await Promise.all([current, ...later, last].map(alive))).toEqual([true, true, true, true])
            } finally {
              for (const id of opened) await manager.close(id)
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

describe("external provider connections", () => {
  const integrationID = Integration.ID.make("claude-agent")
  const projectVariables = ["GOOGLE_VERTEX_PROJECT", "GOOGLE_CLOUD_PROJECT", "GCP_PROJECT", "GCLOUD_PROJECT"]
  const saved = Object.fromEntries(
    ["PATH", "OPENCODE_CLAUDE_PATH", ...projectVariables].map((name) => [name, process.env[name]]),
  )
  const directories: string[] = []
  afterEach(() => {
    for (const [name, value] of Object.entries(saved)) {
      if (value === undefined) delete process.env[name]
      else process.env[name] = value
    }
    for (const directory of directories.splice(0)) rmSync(directory, { recursive: true, force: true })
  })

  function temporary() {
    const directory = mkdtempSync(path.join(tmpdir(), "dochaus-external-test-"))
    directories.push(directory)
    return directory
  }

  function fakeClaude(account: Record<string, unknown>) {
    const directory = temporary()
    const log = path.join(directory, "log.jsonl")
    const executable = path.join(directory, "claude")
    const reply = {
      commands: [],
      agents: [],
      models: [
        {
          value: "default",
          resolvedModel: "synthetic-main",
          displayName: "Default",
          description: "Synthetic default",
          supportedEffortLevels: ["low", "high"],
        },
        {
          value: "synthetic-main",
          resolvedModel: "synthetic-main",
          displayName: "Synthetic Main",
          description: "Synthetic model",
          supportedEffortLevels: ["low", "high"],
        },
      ],
      account,
      claude_code_version: "2.1.250",
    }
    writeFileSync(
      executable,
      `#!${process.execPath}
const fs = require("node:fs")
const log = ${JSON.stringify(log)}
const reply = ${JSON.stringify(reply)}
fs.appendFileSync(log, JSON.stringify({ args: process.argv.slice(2), memory: process.env.CLAUDE_CODE_DISABLE_AUTO_MEMORY, instructions: process.env.CLAUDE_CODE_DISABLE_CLAUDE_MDS }) + "\\n")
let buffer = ""
process.stdin.setEncoding("utf8")
process.stdin.on("data", (chunk) => {
  buffer += chunk
  for (let index = buffer.indexOf("\\n"); index >= 0; index = buffer.indexOf("\\n")) {
    const line = buffer.slice(0, index)
    buffer = buffer.slice(index + 1)
    if (!line.trim()) continue
    const message = JSON.parse(line)
    fs.appendFileSync(log, JSON.stringify({ received: message.type, subtype: message.request?.subtype }) + "\\n")
    if (message.type === "control_request" && message.request?.subtype === "initialize")
      process.stdout.write(JSON.stringify({ type: "control_response", response: { subtype: "success", request_id: message.request_id, response: reply } }) + "\\n")
  }
})
process.stdin.on("end", () => process.exit(0))
`,
    )
    chmodSync(executable, 0o755)
    return {
      executable,
      log: () =>
        readFileSync(log, "utf8")
          .trim()
          .split("\n")
          .map((line) => JSON.parse(line)),
    }
  }

  async function withManager<T>(
    credential: SetupCredential | undefined,
    use: (manager: ReturnType<typeof providerCandidates>, ids: Credential.ID[]) => Promise<T>,
  ) {
    const legal = temporary()
    writeFileSync(path.join(legal, "opencode.json"), "{}")
    return Effect.runPromise(
      Effect.scoped(
        Effect.gen(function* () {
          const baseline = yield* createProviderSetupSession({ legal, settings: {}, credential })
          const ids = (yield* baseline.credentials.all()).map((item) => item.id)
          const manager = providerCandidates({
            legal,
            credentials: baseline.credentials,
            client: undefined as never,
            settings: () => ({}),
            inspect: (effect) => Effect.runPromise(baseline.provide(effect)),
            reuse: credentialReuse({
              directory: path.join(legal, "absent"),
              targetDirectory: legal,
              create: async () => {
                throw new Error("Unexpected import")
              },
            }),
          })
          return yield* Effect.promise(() => use(manager, ids))
        }),
      ).pipe(Effect.provide(quiet)),
    )
  }

  test("external connections accept only a method identifier", () => {
    const input = { providerID: "claude-agent", connection: { type: "external", methodID: "claude-login" } }
    expect(candidateInput(input)).toBe(true)
    expect(candidateInput({ ...input, connection: { type: "external" } })).toBe(false)
    expect(candidateInput({ ...input, connection: { type: "external", methodID: "../claude" } })).toBe(false)
    expect(candidateInput({ ...input, connection: { type: "external", methodID: "claude-login", key: "x" } })).toBe(false)
    expect(
      candidateInput({ ...input, connection: { type: "external", methodID: "adc", answer: { project: "p" } } }),
    ).toBe(true)
    expect(candidateInput({ ...input, connection: { type: "external", methodID: "adc", answer: "p" } })).toBe(false)
  })

  test("a missing Claude Code passes the native needs_auth message through unchanged", async () => {
    const empty = temporary()
    mkdirSync(path.join(empty, "bin"))
    process.env.PATH = path.join(empty, "bin")
    delete process.env.OPENCODE_CLAUDE_PATH
    await withManager(undefined, async (manager) => {
      await expect(
        manager.handle("session", { providerID: "claude-agent", connection: { type: "external", methodID: "unknown" } }),
      ).rejects.toThrow("Select an available sign-in for this provider.")
      const created = (await manager.handle("session", {
        providerID: "claude-agent",
        connection: { type: "external", methodID: "claude-login" },
      })) as { sessionID: string; external: unknown }
      expect(created.external).toEqual({
        status: "needs_auth",
        message: "Claude Code is not installed. Install it, then run `claude auth login`.",
        url: "https://code.claude.com/docs/en/setup",
      })
      await manager.close(created.sessionID)
    })
  })

  test("a signed-in Claude Code is ready and lists models without a prompt or user settings", async () => {
    const claude = fakeClaude({ email: "user@example.com", subscriptionType: "Synthetic", apiProvider: "firstParty" })
    process.env.OPENCODE_CLAUDE_PATH = claude.executable
    await withManager(undefined, async (manager) => {
      const created = (await manager.handle("session", {
        providerID: "claude-agent",
        connection: { type: "external", methodID: "claude-login" },
      })) as { sessionID: string; external: unknown }
      expect(created.external).toEqual({ status: "ready" })
      const models = (await manager.handle("models", { sessionID: created.sessionID })) as { data: { id: string }[] }
      expect(models.data.map((model) => model.id)).toEqual(["synthetic-main"])
      await manager.close(created.sessionID)
    })
    const log = claude.log()
    const spawns = log.filter((line) => line.args)
    expect(spawns.length).toBeGreaterThan(0)
    for (const spawn of spawns) {
      expect(spawn.args).toContain("--setting-sources=")
      expect(spawn.args).toContain("--strict-mcp-config")
      expect(spawn.instructions).toBe("1")
      expect(spawn.memory).toBe("1")
    }
    expect(log.some((line) => line.received === "user")).toBe(false)
  })

  test("a saved external connection is checked again and reports a signed-out Claude Code", async () => {
    const claude = fakeClaude({ tokenSource: "none", apiProvider: "firstParty" })
    process.env.OPENCODE_CLAUDE_PATH = claude.executable
    const credential = {
      integrationID,
      label: "Claude Agent",
      value: Credential.External.make({ type: "external", methodID: Integration.MethodID.make("claude-login") }),
    }
    await withManager(credential, async (manager, [id]) => {
      const created = (await manager.handle("session", {
        providerID: "claude-agent",
        connection: { type: "credential", id: id! },
      })) as { sessionID: string; external: unknown }
      expect(created.external).toEqual({
        status: "needs_auth",
        message: "Claude Code is not logged in. Run `claude auth login` in a terminal.",
      })
      await manager.close(created.sessionID)
    })
    expect(claude.log().some((line) => line.received === "user")).toBe(false)
  })

  test("an external method with fields connects with its answers and rejects incomplete ones", async () => {
    for (const name of projectVariables) delete process.env[name]
    await withManager(undefined, async (manager) => {
      const connect = (answer?: Record<string, string>) =>
        manager.handle("session", {
          providerID: "google-vertex",
          connection: { type: "external", methodID: "google-adc", ...(answer ? { answer } : {}) },
        }) as Promise<{ sessionID: string; external: unknown }>
      await expect(connect()).rejects.toThrow("Complete the provider connection fields.")
      await expect(connect({ location: "global" })).rejects.toThrow("Complete the provider connection fields.")
      await expect(connect({ project: "synthetic-project", location: "global", extra: "x" })).rejects.toThrow(
        "Complete the provider connection fields.",
      )
      const created = await connect({ project: "synthetic-project", location: "global" })
      expect(created.external).toEqual({ status: "ready" })
      const models = (await manager.handle("models", { sessionID: created.sessionID })) as { data: unknown[] }
      expect(models.data.length).toBeGreaterThan(0)
      await manager.close(created.sessionID)
    })
  })
})
