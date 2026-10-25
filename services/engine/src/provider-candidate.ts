import { LLMClient } from "@opencode/ai"
import { Credential } from "@opencode/core/credential"
import { Integration } from "@opencode/core/integration"
import { Instance } from "@opencode/core/instance"
import { Model } from "@opencode/core/model"
import { Provider } from "@opencode/core/provider"
import { Form } from "@opencode/core/form"
import { Form as FormSchema } from "@opencode/schema/form"
import { Context, Effect, Exit, Logger, Schema, Scope } from "effect"
import {
  mergeSettings,
  readEffectiveSettings,
  readSettingsFile,
  settingsPaths,
  type EngineSettings,
} from "./effective-settings"
import { credentialReuse, createReusedCredential, type Stored } from "./credential-reuse"
import { probeProvider, providerTestFailure } from "./provider-health"
import { createProviderSetupSession } from "./provider-setup-session"

type Connection =
  | { type: "key"; key: string; answer?: FormSchema.Answer }
  | { type: "credential"; id: string }
  | { type: "env"; name: string }
  | { type: "host" }
  | { type: "new" }
  | { type: "reuse"; sourceId: string; connectionId: string }
export type CandidateInput = {
  providerID: string
  options?: Record<string, string | null>
  connection: Connection
  local?: { name: string; baseURL: string; modelIDs: string[] }
}
const object = (value: unknown): value is Record<string, unknown> =>
  !!value && typeof value === "object" && !Array.isArray(value)
const identifier = (value: unknown): value is string =>
  typeof value === "string" && /^[a-zA-Z0-9][a-zA-Z0-9._-]{0,127}$/.test(value)
const modelIdentifier = (value: unknown): value is string =>
  typeof value === "string" && value.length > 0 && value.length <= 256 && !/[\s#]/.test(value)
const modelReference = (value: unknown): value is string => {
  if (typeof value !== "string") return false
  const [id, variant, extra] = value.split("#")
  return modelIdentifier(id) && extra === undefined && (variant === undefined || modelIdentifier(variant))
}
const keys = (value: Record<string, unknown>, allowed: string[]) =>
  Object.keys(value).every((key) => allowed.includes(key))
export class ProviderCandidateError extends Error {}
const compatible = (provider: Provider.Info) =>
  /(?:openai-compatible|ollama)$/.test(provider.package) || provider.id === "ollama"

export function candidateInput(value: unknown): value is CandidateInput {
  if (
    !object(value) ||
    !keys(value, ["providerID", "options", "connection", "local"]) ||
    !identifier(value.providerID) ||
    !object(value.connection)
  )
    return false
  if (
    value.local !== undefined &&
    (!object(value.local) ||
      !keys(value.local, ["name", "baseURL", "modelIDs"]) ||
      typeof value.local.name !== "string" ||
      !value.local.name.trim() ||
      value.local.name.length > 128 ||
      typeof value.local.baseURL !== "string" ||
      value.local.baseURL.length > 2048 ||
      !Array.isArray(value.local.modelIDs) ||
      value.local.modelIDs.length > 16 ||
      !value.local.modelIDs.every(modelIdentifier))
  )
    return false
  const options = value.options
  if (
    options !== undefined &&
    (!object(options) ||
      Object.entries(options).some(
        ([key, value]) =>
          !["project", "location", "region", "profile", "baseURL", "endpoint", "resourceName", "apiVersion"].includes(
            key,
          ) ||
          (value !== null && (typeof value !== "string" || !value.trim() || value.length > 2048)),
      ))
  )
    return false
  const connection = value.connection
  switch (connection.type) {
    case "key":
      return (
        keys(connection, ["type", "key", "answer"]) &&
        typeof connection.key === "string" &&
        connection.key.trim().length > 0 &&
        connection.key.length <= 16384 &&
        (connection.answer === undefined || Schema.is(FormSchema.Answer)(connection.answer))
      )
    case "credential":
      return keys(connection, ["type", "id"]) && typeof connection.id === "string" && connection.id.length <= 256
    case "env":
      return (
        keys(connection, ["type", "name"]) &&
        typeof connection.name === "string" &&
        /^[A-Z][A-Z0-9_]{0,127}$/.test(connection.name)
      )
    case "reuse":
      return (
        keys(connection, ["type", "sourceId", "connectionId"]) &&
        typeof connection.sourceId === "string" &&
        typeof connection.connectionId === "string" &&
        connection.sourceId.length <= 256 &&
        connection.connectionId.length <= 256
      )
    case "host":
    case "new":
      return keys(connection, ["type"])
    default:
      return false
  }
}

function candidateOptions(input: CandidateInput) {
  const options = Object.fromEntries(
    Object.entries(input.options ?? {}).map(([key, value]) => [key, value === null ? null : value.trim()]),
  )
  if (input.providerID === "google-vertex-anthropic" && options.project) {
    const location = options.location || "global"
    const host = location === "global" ? "aiplatform.googleapis.com" : `${location}-aiplatform.googleapis.com`
    options.baseURL = `https://${host}/v1/projects/${encodeURIComponent(options.project)}/locations/${encodeURIComponent(location)}/publishers/anthropic/models`
  }
  return options
}

export function candidateSettings(
  policy: EngineSettings,
  input: CandidateInput,
  provider: Provider.Info,
  fallback: EngineSettings = {},
) {
  if (
    provider.activation === "disabled" ||
    policy.disabled_providers?.includes(input.providerID) ||
    (Array.isArray(policy.enabled_providers) && !policy.enabled_providers.includes(input.providerID))
  )
    throw new ProviderCandidateError("This provider is disabled by installation policy.")
  const allowed = ["google-vertex", "google-vertex-anthropic"].includes(input.providerID)
    ? ["project", "location", "baseURL"]
    : input.providerID === "amazon-bedrock"
      ? ["region", "profile", "endpoint", "baseURL"]
      : ["azure", "azure-openai"].includes(input.providerID)
        ? ["resourceName", "baseURL", "apiVersion"]
        : compatible(provider)
          ? ["baseURL"]
          : []
  if (Object.keys(input.options ?? {}).some((key) => !allowed.includes(key)))
    throw new ProviderCandidateError("These options are not supported by this provider.")
  const options = candidateOptions(input)
  const inherited = fallback.provider?.[input.providerID]?.options
  for (const [key, value] of Object.entries(options))
    if (value === null && object(inherited) && typeof inherited[key] === "string") options[key] = inherited[key]
  if (input.local) options.baseURL = input.local.baseURL.trim()
  for (const endpoint of [options.baseURL, options.endpoint].filter(
    (value): value is string => typeof value === "string",
  )) {
    const url = new URL(endpoint)
    if (!["http:", "https:"].includes(url.protocol) || url.username || url.password || url.search || url.hash)
      throw new ProviderCandidateError("Enter an HTTP or HTTPS model endpoint without credentials or a fragment.")
  }
  return mergeSettings(policy, {
    provider: {
      [input.providerID]: {
        ...(input.local
          ? {
              npm: "@ai-sdk/openai-compatible",
              name: input.local.name.trim(),
              models: Object.fromEntries(input.local.modelIDs.map((id) => [id, { name: id }])),
            }
          : {}),
        options: { ...options },
      },
    },
  })
}

type SetupSession = Effect.Success<ReturnType<typeof createProviderSetupSession>>
type Entry = {
  session: SetupSession
  scope: Scope.Closeable
  input: CandidateInput
  integrationID: Integration.ID
  touched: number
  tested: Set<string>
  saved?: { connection?: { id: string } }
  created?: string
  previous?: string
  settings: EngineSettings
  modelIDs: Set<string>
  pending: Promise<unknown>
  customModels: boolean
}

// Drafts have their own native graph and memory credentials. Only /save writes a
// credential to the installation. Discard and failed probes leave it untouched.
export function providerCandidates(options: {
  legal: string
  credentials: Credential.Interface
  client: Context.Service.Shape<typeof LLMClient.Service>
  settings?: () => EngineSettings
  scope?: Scope.Scope
  reuse: ReturnType<typeof credentialReuse>
  inspect: <A, E>(effect: Effect.Effect<A, E, Provider.Service | Integration.Service>) => Promise<A>
}) {
  const sessions = new Map<string, Entry>()
  let saving: string | undefined
  let creating = 0
  const run = <A, E>(effect: Effect.Effect<A, E>) => Effect.runPromise(effect.pipe(Effect.provide(Logger.layer([]))))
  const close = async (id: string) => {
    const entry = sessions.get(id)
    if (!entry) return
    while (true) {
      const pending = entry.pending
      await pending.catch(() => {})
      if (entry.pending === pending) break
    }
    sessions.delete(id)
    if (saving === id) saving = undefined
    await run(Scope.close(entry.scope, Exit.void))
  }
  const within = <A, E>(
    entry: Entry,
    effect: Effect.Effect<A, E, Instance.Services | Credential.Service | Scope.Scope>,
  ) => run(entry.session.provide(effect).pipe(Effect.provideService(Scope.Scope, entry.scope)))
  async function create(input: CandidateInput) {
    input = { ...input, options: candidateOptions(input) }
    for (const [id, entry] of sessions) if (entry.touched < Date.now() - 30 * 60_000) await close(id)
    // Each draft holds a native instance graph, so at most four exist. A page
    // reloaded or closed mid-setup cannot discard its draft, so a new setup
    // replaces the least recently used draft that is not being saved.
    while (sessions.size + creating >= 4) {
      let oldest: [string, Entry] | undefined
      for (const item of sessions)
        if (item[0] !== saving && (!oldest || item[1].touched < oldest[1].touched)) oldest = item
      if (!oldest) throw new ProviderCandidateError("Close an unfinished setup before starting another.")
      await close(oldest[0])
    }
    creating++
    try {
      const { provider, integration } = await options.inspect(
        Effect.gen(function* () {
          const provider = yield* Provider.Service.use((service) => service.get(Provider.ID.make(input.providerID)))
          if (input.local && provider) throw new ProviderCandidateError("Choose a unique name for the local provider.")
          if (!provider && !input.local) throw new ProviderCandidateError("Select an available provider.")
          if (!provider)
            return {
              provider: {
                ...Provider.Info.empty(Provider.ID.make(input.providerID)),
                package: "@opencode/ai/providers/openai-compatible",
                integrationID: Integration.ID.make(input.providerID),
              },
              integration: undefined,
            }
          const integration = yield* Integration.Service.use((service) =>
            service.get(provider.integrationID ?? Integration.ID.make(provider.id)),
          )
          return { provider, integration }
        }),
      )
      const policy = (options.settings ?? readEffectiveSettings)()
      const settings = candidateSettings(
        policy,
        input,
        provider,
        options.settings ? {} : readSettingsFile(settingsPaths().shipped),
      )
      const integrationID = provider.integrationID ?? Integration.ID.make(input.providerID)
      let credential: Stored | undefined
      const selected = input.connection
      if (selected.type === "key") {
        const answer = selected.answer ?? {}
        credential = {
          integrationID,
          label: "API key",
          value: Credential.Key.make({
            type: "key",
            key: selected.key.trim(),
            ...(Object.keys(answer).length ? { configuration: answer } : {}),
          }),
        }
      } else if (selected.type === "credential") {
        const saved = await run(options.credentials.get(Credential.ID.make(selected.id)))
        if (!saved || saved.integrationID !== integrationID)
          throw new ProviderCandidateError("Select a saved connection for this provider.")
        credential = { integrationID, label: saved.label, value: saved.value }
      } else if (selected.type === "env") {
        if (
          !integration?.methods.some((method) => method.type === "env" && method.names.includes(selected.name)) ||
          !process.env[selected.name]
        )
          throw new ProviderCandidateError("Select an available provider environment connection.")
        credential = {
          integrationID,
          label: selected.name,
          value: Credential.Key.make({ type: "key", key: process.env[selected.name]! }),
        }
      } else if (selected.type === "reuse") {
        await options.reuse.importSelected(
          { sourceId: selected.sourceId, connectionId: selected.connectionId },
          async (stored) => {
            credential = stored
            return { id: crypto.randomUUID() }
          },
        )
        if (credential?.integrationID !== integrationID)
          throw new ProviderCandidateError("Select an existing connection for this provider.")
      } else if (
        selected.type === "host" &&
        !["google-vertex", "google-vertex-anthropic", "amazon-bedrock", "ollama"].includes(provider.id) &&
        !compatible(provider)
      )
        throw new ProviderCandidateError("This provider requires a connection.")
      if (
        selected.type === "host" &&
        (integration?.connections.length ||
          (await run(options.credentials.list(integrationID))).length ||
          ["apiKey", "accessToken", "authToken"].some((key) => {
            const values = policy.provider?.[provider.id]?.options
            return object(values) && values[key]
          }))
      )
        throw new ProviderCandidateError(
          "Host credentials cannot replace this provider’s existing connection here. Choose a saved connection or enter a new key.",
        )
      const scope = await run(options.scope ? Scope.fork(options.scope) : Scope.make())
      try {
        const session = await run(
          createProviderSetupSession({ legal: options.legal, settings, credential }).pipe(
            Effect.provideService(Scope.Scope, scope),
          ),
        )
        if (selected.type === "key" && !input.local) {
          const definition = await run(
            session.provide(Integration.Service.use((service) => service.get(integrationID))),
          )
          const method = definition?.methods.find((method) => method.type === "key")
          const answer = selected.answer ?? {}
          if (
            (!method && !compatible(provider)) ||
            (method?.form
              ? Form.validateFields(method.form) || Form.validateAnswer(method.form, answer)
              : Object.keys(answer).length)
          )
            throw new ProviderCandidateError("Complete the provider connection fields.")
        }
        const id = crypto.randomUUID()
        const entry: Entry = {
          session,
          scope,
          input,
          integrationID,
          touched: Date.now(),
          tested: new Set(),
          settings,
          modelIDs: new Set(input.local?.modelIDs ?? []),
          pending: Promise.resolve(),
          customModels: !!input.local || (!!input.options?.baseURL && compatible(provider)),
        }
        sessions.set(id, entry)
        return { sessionID: id }
      } catch (error) {
        await run(Scope.close(scope, Exit.void))
        throw error
      }
    } finally {
      creating--
    }
  }
  return {
    close,
    async handle(action: string, input: unknown) {
      if (action === "session") {
        if (!candidateInput(input)) throw new ProviderCandidateError("Choose a provider and connection.")
        return create(input)
      }
      if (!object(input) || typeof input.sessionID !== "string")
        throw new ProviderCandidateError("Start provider setup again.")
      const entry = sessions.get(input.sessionID)
      if (!entry || entry.touched < Date.now() - 30 * 60_000) {
        if (entry) await close(input.sessionID)
        throw new ProviderCandidateError("Provider setup expired. Start again.")
      }
      entry.touched = Date.now()
      // Cancellation must reach the native attempt while a code exchange is
      // pending. All credential saves and rollbacks still share the entry queue.
      if (action === "auth" && ["oauth.cancel", "command.cancel"].includes(String(input.action))) {
        if (
          !keys(input, ["sessionID", "action", "attemptID"]) ||
          typeof input.attemptID !== "string" ||
          input.attemptID.length > 256
        )
          throw new ProviderCandidateError("Choose the pending connection to cancel.")
        entry.tested.clear()
        const attempt = { integrationID: entry.integrationID, attemptID: Integration.AttemptID.make(input.attemptID) }
        await within(
          entry,
          Integration.Service.use((service) =>
            input.action === "oauth.cancel" ? service.oauth.cancel(attempt) : service.command.cancel(attempt),
          ).pipe(Effect.timeout("20 seconds")),
        )
        return { data: undefined }
      }
      const job = entry.pending
        .catch(() => {})
        .then(async () => {
          if (action === "models") {
            if (!keys(input, ["sessionID", "modelIDs"]))
              throw new ProviderCandidateError("Choose catalog models or explicit endpoint model IDs.")
            if (input.modelIDs !== undefined) {
              if (
                !Array.isArray(input.modelIDs) ||
                input.modelIDs.length > 16 ||
                !input.modelIDs.every(modelIdentifier) ||
                !entry.customModels
              )
                throw new ProviderCandidateError("Enter model IDs for a compatible endpoint.")
              for (const id of input.modelIDs as string[]) entry.modelIDs.add(id)
              entry.settings = mergeSettings(entry.settings, {
                provider: {
                  [entry.input.providerID]: {
                    models: Object.fromEntries(Array.from(entry.modelIDs, (id) => [id, { name: id }])),
                  },
                },
              })
              if (entry.input.local)
                entry.input.local.modelIDs = [...new Set([...entry.input.local.modelIDs, ...input.modelIDs])]
              await within(
                entry,
                Provider.Service.use((service) =>
                  service.transform((providers) => {
                    const providerID = Provider.ID.make(entry.input.providerID)
                    const existing = providers.get(providerID)
                    if (!existing || !compatible(existing.provider)) return
                    for (const id of input.modelIDs as string[])
                      if (!existing.models.has(Model.ID.make(id)))
                        providers.models.update(providerID, Model.ID.make(id), (model) => {
                          model.name = id
                          model.package = existing.provider.package
                        })
                  }),
                ),
              )
            }
            const data = await within(
              entry,
              Effect.gen(function* () {
                const snapshot = yield* Provider.Service.use((providers) => providers.snapshot())
                const definitions = snapshot.records.get(Provider.ID.make(entry.input.providerID))
                const available = yield* Model.Service.use((models) => models.all())
                return Array.from(definitions?.models.values() ?? [])
                  .map(
                    (model) =>
                      available.find((current) => current.providerID === model.providerID && current.id === model.id) ??
                      model,
                  )
                  .filter((model) => model.enabled && model.status !== "deprecated")
                  .map(({ id, providerID, name, limit, capabilities, cost, variants }) => ({
                    id,
                    providerID,
                    name,
                    limit,
                    capabilities,
                    cost,
                    variants: variants.map(({ id }) => ({ id, name: id })),
                  }))
              }),
            )
            return { data, customModels: entry.customModels }
          }
          if (action === "test") {
            if (!modelReference(input.modelID) || !keys(input, ["sessionID", "modelID"]))
              throw new ProviderCandidateError("Choose a model to test.")
            entry.tested.delete(input.modelID)
            const ref = Model.Ref.parse(`${entry.input.providerID}/${input.modelID}`)
            if (entry.input.connection.type === "new" && !entry.session.credential(entry.integrationID))
              throw new ProviderCandidateError("Finish connecting this provider before testing.")
            const result = await within(
              entry,
              probeProvider(ref).pipe(
                Effect.provideService(LLMClient.Service, options.client),
                Effect.timeout("20 seconds"),
              ),
            ).catch(providerTestFailure)
            if (result.ok) entry.tested.add(input.modelID)
            return result
          }
          if (action === "auth") {
            if (
              !keys(input, ["sessionID", "action", "methodID", "attemptID", "answer", "code"]) ||
              (input.answer !== undefined && !Schema.is(FormSchema.Answer)(input.answer)) ||
              (input.methodID !== undefined && !identifier(input.methodID)) ||
              (input.attemptID !== undefined &&
                (typeof input.attemptID !== "string" || input.attemptID.length > 256)) ||
              (input.code !== undefined && (typeof input.code !== "string" || input.code.length > 8192))
            )
              throw new ProviderCandidateError("Complete the provider connection fields.")
            entry.tested.clear()
            const integrationID = entry.integrationID
            const methodID = typeof input.methodID === "string" ? Integration.MethodID.make(input.methodID) : undefined
            const attemptID =
              typeof input.attemptID === "string" ? Integration.AttemptID.make(input.attemptID) : undefined
            const data = await within(
              entry,
              Integration.Service.use((service) => {
                switch (input.action) {
                  case "oauth.connect":
                    if (methodID)
                      return service.oauth.connect({
                        integrationID,
                        methodID,
                        ...(Schema.is(FormSchema.Answer)(input.answer) ? { answer: input.answer } : {}),
                      })
                    break
                  case "oauth.status":
                    if (attemptID) return service.oauth.status({ integrationID, attemptID })
                    break
                  case "oauth.complete":
                    if (attemptID)
                      return service.oauth
                        .complete({
                          integrationID,
                          attemptID,
                          ...(typeof input.code === "string" ? { code: input.code } : {}),
                        })
                        .pipe(Effect.andThen(service.oauth.status({ integrationID, attemptID })))
                    break
                  case "oauth.cancel":
                    if (attemptID) return service.oauth.cancel({ integrationID, attemptID })
                    break
                  case "command.connect":
                    if (methodID) return service.command.connect({ integrationID, methodID })
                    break
                  case "command.status":
                    if (attemptID) return service.command.status({ integrationID, attemptID })
                    break
                  case "command.cancel":
                    if (attemptID) return service.command.cancel({ integrationID, attemptID })
                    break
                }
                return Effect.fail(new Error("Choose a supported connection action."))
              }).pipe(Effect.timeout("20 seconds")),
            )
            if (object(data) && data.status === "complete") {
              const credential = entry.session.credential(entry.integrationID)
              if (credential) {
                const scope = await run(options.scope ? Scope.fork(options.scope) : Scope.make())
                try {
                  const session = await run(
                    createProviderSetupSession({ legal: options.legal, settings: entry.settings, credential }).pipe(
                      Effect.provideService(Scope.Scope, scope),
                    ),
                  )
                  await run(Scope.close(entry.scope, Exit.void))
                  entry.scope = scope
                  entry.session = session
                } catch (error) {
                  await run(Scope.close(scope, Exit.void))
                  throw error
                }
              }
            }
            if (object(data) && typeof data.message === "string")
              data.message =
                data.status === "failed"
                  ? "The provider could not complete sign-in. Check the connection and retry."
                  : data.message.slice(0, 8192)
            return { data }
          }
          if (["prepare", "save"].includes(action)) {
            if (!keys(input, ["sessionID", "primary", "fast"]))
              throw new ProviderCandidateError("Choose both models before saving.")
            if (
              !modelReference(input.primary) ||
              !modelReference(input.fast) ||
              !entry.tested.has(input.primary) ||
              !entry.tested.has(input.fast)
            )
              throw new ProviderCandidateError("Test both selected models before saving.")
            const patch = {
              model: `${entry.input.providerID}/${input.primary}`,
              small_model: `${entry.input.providerID}/${input.fast}`,
              provider: {
                [entry.input.providerID]: {
                  ...(entry.input.local
                    ? { npm: "@ai-sdk/openai-compatible", name: entry.input.local.name.trim() }
                    : {}),
                  ...(entry.modelIDs.size
                    ? { models: Object.fromEntries(Array.from(entry.modelIDs, (id) => [id, { name: id }])) }
                    : {}),
                  options: {
                    ...entry.input.options,
                    ...(entry.input.local ? { baseURL: entry.input.local.baseURL.trim() } : {}),
                  },
                },
              },
            }
            const policy = (options.settings ?? readEffectiveSettings)()
            if (
              policy.disabled_providers?.includes(entry.input.providerID) ||
              (Array.isArray(policy.enabled_providers) && !policy.enabled_providers.includes(entry.input.providerID))
            )
              throw new ProviderCandidateError("This provider is disabled by installation policy.")
            const selectedModels = await within(
              entry,
              Model.Service.use((models) =>
                Effect.all([
                  ...[input.primary as string, input.fast as string].map((selected) => {
                    const ref = Model.Ref.parse(`${entry.input.providerID}/${selected}`)
                    return models
                      .get(ref.providerID, ref.id)
                      .pipe(
                        Effect.map((model) =>
                          model &&
                          (ref.variant === undefined || model.variants.some((variant) => variant.id === ref.variant))
                            ? model
                            : undefined,
                        ),
                      )
                  }),
                ]),
              ),
            )
            if (
              selectedModels.some(
                (model) =>
                  !model?.enabled ||
                  !model.capabilities.input.includes("text") ||
                  !model.capabilities.output.includes("text"),
              ) ||
              !selectedModels[0]?.capabilities.tools
            )
              throw new ProviderCandidateError("Choose available text models and a primary model that supports tools.")
            if (action === "prepare") return { patch }
            if (saving && saving !== input.sessionID)
              throw new ProviderCandidateError(
                "Another provider setup is being saved. Finish it before saving this setup.",
              )
            saving = input.sessionID as string
            if (entry.saved) return entry.saved
            const active = await options.inspect(
              Integration.Service.use((service) => service.connection.active(entry.integrationID)),
            )
            entry.previous = active?.type === "credential" ? active.id : undefined
            if (!entry.tested.size) throw new ProviderCandidateError("Test the selected models before saving.")
            const selected = entry.input.connection
            const snapshot = entry.session.credential(entry.integrationID)
            const original =
              selected.type === "credential"
                ? await run(options.credentials.get(Credential.ID.make(selected.id)))
                : undefined
            if (selected.type === "credential" && JSON.stringify(snapshot?.value) === JSON.stringify(original?.value)) {
              await run(options.credentials.activate(Credential.ID.make(selected.id)))
              return (entry.saved = { connection: { id: selected.id } })
            }
            const credential = entry.session.credential(entry.integrationID)
            if (!credential || selected.type === "host") return (entry.saved = {})
            const saved = await run(createReusedCredential(options.credentials, credential))
            entry.created = saved.id
            return (entry.saved = { connection: { id: saved.id } })
          }
          if (action === "rollback") {
            if (!keys(input, ["sessionID"]))
              throw new ProviderCandidateError("Use the current provider setup to restore the connection.")
            if (!entry.saved && !entry.previous && !entry.created && saving !== input.sessionID) return { ok: true }
            if (saving && saving !== input.sessionID)
              throw new ProviderCandidateError("Another provider setup is being saved.")
            if (entry.previous) await run(options.credentials.activate(Credential.ID.make(entry.previous)))
            if (entry.created) await run(options.credentials.remove(Credential.ID.make(entry.created)))
            entry.saved = undefined
            entry.created = undefined
            entry.previous = undefined
            if (saving === input.sessionID) saving = undefined
            return { ok: true }
          }
          throw new ProviderCandidateError("Unknown setup action.")
        })
      entry.pending = job
      return job
    },
  }
}
