import { startupObservation, startupObservationEnabled, startupObservationSpan } from "../../../dochaus/lib/startup-observation"
import { diagnosticErrorMetadata } from "../../../dochaus/lib/diagnostic-metadata"
import { appendTrace, describeError, ensureTraceRoot, resolveTraceRoot, sha256, TRACE_ROOT_ENV } from "../../../dochaus/lib/trace-sink"
import { NodeServices } from "@effect/platform-node"
import { LLMClient } from "@opencode/ai"
import { RequestExecutor } from "@opencode/ai/route"
import { Config } from "@opencode/core/config"
import { AbsolutePath } from "@opencode/core/schema"
import { Bus } from "@opencode/core/bus"
import { V1Migration } from "@opencode/core/database/v1-migration"
import { Database } from "@opencode/core/database/database"
import { importLegacyCredentials } from "@opencode/core/database/migration/20260805200742_import_legacy_credentials"
import { Credential } from "@opencode/core/credential"
import { InstructionDiscovery } from "@opencode/core/instruction-discovery"
import { Session } from "@opencode/core/session"
import { SessionRestart } from "@opencode/core/session/execution/restart"
import { Location } from "@opencode/core/location"
import { LocationServiceMap } from "@opencode/core/location-service-map"
import { ManagedPolicy } from "@opencode/core/managed-policy"
import { WellKnown } from "@opencode/core/wellknown"
import { Watcher } from "@opencode/core/filesystem/watcher"
import { nativeFailureEvidenceLogger } from "./native-failure-evidence"
import { generateMatterText } from "./generation"
import { createRoutes } from "@opencode/server/routes"
import { requestRef } from "@opencode/server/location"
import { authorizedRequest } from "@opencode/server/middleware/authorization"
import { Plugin } from "@opencode/core/plugin"
import { Agent } from "@opencode/schema/agent"
import { Model } from "@opencode/core/model"
import { Provider } from "@opencode/core/provider"
import { Permission } from "@opencode/schema/permission"
import { Session as SessionSchema } from "@opencode/schema/session"
import { makeLocationNode } from "@opencode/util/effect/app-node"
import { FSUtil } from "@opencode/util/fs-util"
import { Global } from "@opencode/util/global"
import { Context, Deferred, Effect, Layer, Logger, Option, References, Schema, Tracer } from "effect"
import { HttpClient, HttpEffect, HttpMiddleware, HttpRouter, HttpServer, HttpServerRequest } from "effect/unstable/http"
import path from "node:path"
import { existsSync, mkdirSync, openSync, closeSync, readFileSync, unlinkSync, writeFileSync, realpathSync } from "node:fs"
import { dataPaths } from "./data-paths"
import { legalConfigLayer, legalManagedPolicy } from "./config-layer"
import { generationAdmissionError } from "./generation-admission"
import { readEffectiveSettings } from "./effective-settings"
import { createReusedCredential, credentialReuse, credentialReuseHandler, nativeCredentialDirectory } from "./credential-reuse"
import { researchConsentStatus } from "../../../dochaus/lib/research-consent-status"
import { providerCatalog } from "./provider-catalog"
import { modelHttp } from "./model-http"
import { probeProvider, providerTestCache, providerTestFailure, providerTestInput, unavailableModel } from "./provider-health"

const repository = path.resolve(import.meta.dir, "../../..")
const legal = path.join(repository, "dochaus")
const paths = dataPaths(repository)
const workspace = paths.workspace
const runtime = paths.runtime
const hostname = process.env.OPENCODE_HOSTNAME ?? "127.0.0.1"
const port = Number(process.env.DOC_HAUS_ENGINE_PORT ?? "4097")
const password = process.env.OPENCODE_SERVER_PASSWORD
const managedRestart = process.env.DOCHAUS_MANAGED_RESTART === "1"
const verbosity = process.env.DOCHAUS_LOG_LEVEL?.trim().toLowerCase()
const minimumLogLevel = verbosity === "debug" ? "Debug" : verbosity === "info" ? "Info" : "Warn"
// Request helpers start independent fibers too; share the threshold with startup.
const runEngine = Effect.runPromiseWith(Context.make(References.MinimumLogLevel, minimumLogLevel))
const cachedProviderTest = providerTestCache()
process.env.OPENCODE_CONFIG_DIR = legal
process.env.WORKSPACE_ROOT = workspace
process.env.DOCHAUS_RUNTIME_DIR = runtime

if (!["127.0.0.1", "localhost", "::1"].includes(hostname) && !password)
  throw new Error("OPENCODE_SERVER_PASSWORD is required for non-loopback binding")
if (!Number.isInteger(port) || port < 0 || port > 65_535) throw new Error("Invalid DOC_HAUS_ENGINE_PORT")

// Capture is opt-in. An explicit unsafe capture path still stops startup.
const requestedTrace = process.env[TRACE_ROOT_ENV]?.trim() ? resolveTraceRoot() : undefined
const tracing = requestedTrace ? await ensureTraceRoot(requestedTrace, { repository, protectedPaths: [workspace, legal] }) : undefined
if (tracing) process.env[TRACE_ROOT_ENV] = tracing

// Only the shipped legal configuration and its agent/tool/skill directory are
// trusted. The user's global config and all ancestor project/compatibility roots
// stay outside the location graph, even when a matter lives beneath $HOME.
const legalConfig = makeLocationNode({
  service: Config.Service,
  layer: legalConfigLayer(legal),
  deps: [Watcher.node, Bus.node, FSUtil.node, Global.node, Location.node, Credential.node, WellKnown.node],
})

const ChildInput = Schema.Struct({
  parentID: SessionSchema.ID,
  id: SessionSchema.ID.pipe(Schema.optional),
  title: Schema.String.pipe(Schema.optional),
  agent: Agent.ID.pipe(Schema.optional),
  model: Model.Ref.pipe(Schema.optional),
  metadata: SessionSchema.Metadata.pipe(Schema.optional),
  permissions: Permission.Ruleset.pipe(Schema.optional),
})

// Ownership belongs to the engine process itself, including direct starts. Fail
// closed on an abandoned lock: deleting it is safe only after checking its PID.
mkdirSync(runtime, { recursive: true })
const ownerFile = path.join(runtime, "engine-process.lock")
try {
  const descriptor = openSync(ownerFile, "wx", 0o600)
  writeFileSync(descriptor, String(process.pid))
  closeSync(descriptor)
} catch {
  throw new Error(`Engine runtime ownership is already recorded at ${ownerFile}; verify that its PID has exited before removing an abandoned lock`)
}
process.once("exit", () => {
  if (readFileSync(ownerFile, "utf8") === String(process.pid)) unlinkSync(ownerFile)
})
const databasePath = path.join(runtime, "data", "opencode", "opencode-local.db")
const legacyAuthPath = path.join(runtime, "data", "opencode", "auth.json")
const pendingAuthImport = path.join(runtime, "legacy-auth-import.pending")
// Fresh native database bootstrap skips historical migration bodies. Record
// this specific transition before opening it so an interrupted import resumes.
if (!existsSync(databasePath) && existsSync(legacyAuthPath)) writeFileSync(pendingAuthImport, "pending", { mode: 0o600 })
process.chdir(workspace)

// Replace, rather than merge alongside, raw default drain-failure logging.
const failureEvidence = Logger.layer([Effect.gen(function* () {
  const loggers = yield* Logger.CurrentLoggers
  return nativeFailureEvidenceLogger({ repository, runtime, loggers })
})])

await runEngine(
  Effect.scoped(
    Effect.gen(function* () {
      const loggers = yield* Logger.CurrentLoggers
      const options = {
          app: { name: "doc.haus", version: "2.0.19" },
          password,
          config: { directory: path.join(runtime, "config"), project: false },
          database: { path: databasePath },
          models: { fetch: process.env.OPENCODE_DISABLE_MODELS_FETCH !== "1" },
      }
      startupObservation("graph", "enter")
      const tracer = Tracer.make({
        span(options) {
          const span = new Tracer.NativeSpan(options)
          startupObservationSpan(options.name, "enter", span.spanId)
          const end = span.end.bind(span)
          span.end = (...args) => {
            end(...args)
            startupObservationSpan(options.name, "exit", span.spanId, args[1]._tag)
          }
          return span
        },
      })
      const building = Layer.build(
        createRoutes(options, () => [], [
            modelHttp,
            Global.node.replace(Global.layerWith({
              data: path.join(runtime, "data", "opencode"),
              config: path.join(runtime, "config", "opencode"),
              state: path.join(runtime, "state", "opencode"),
              cache: path.join(runtime, "cache", "opencode"),
              bin: path.join(runtime, "cache", "opencode", "bin"),
              log: path.join(runtime, "data", "opencode", "log"),
              repos: path.join(runtime, "data", "opencode", "repos"),
            })),
            WellKnown.node.replace(Layer.succeed(WellKnown.Service, {
              entries: () => Effect.succeed([]),
              snapshot: () => [],
              refresh: () => Effect.succeed(false),
              add: () => Effect.fail(new Error("Remote configuration discovery is disabled in the legal engine")),
              remove: () => Effect.void,
              resolve: () => Effect.succeed([]),
            })),
            Config.node.replace(legalConfig),
            ManagedPolicy.node.replace(legalManagedPolicy),
            InstructionDiscovery.node.replace(InstructionDiscovery.configured({ global: false, project: false })),
        ]).pipe(Layer.provide(HttpServer.layerServices)),
      )
      const context = yield* (startupObservationEnabled ? building.pipe(Effect.provideService(Tracer.Tracer, tracer)) : building)
      startupObservation("graph", "exit")
      const httpContext = (startupObservationEnabled ? Context.add(context, Tracer.Tracer, tracer) : context).pipe(
        Context.add(Logger.CurrentLoggers, loggers),
        Context.add(References.MinimumLogLevel, minimumLogLevel),
      )
      if (existsSync(pendingAuthImport)) {
        yield* Context.get(context, Database.Service).db.transaction((tx) => Effect.gen(function* () {
          const imported = yield* tx.get("SELECT key FROM kv WHERE key = 'dochaus:legacy-auth-import'")
          if (imported) return
          yield* importLegacyCredentials(tx, legacyAuthPath)
          yield* tx.run("INSERT INTO kv (key, value, time_created, time_updated) VALUES ('dochaus:legacy-auth-import', 'true', 0, 0)")
        }))
        unlinkSync(pendingAuthImport)
      }
      // Native routes start migration in the background. Join its serialized
      // migration before any restart sweep or HTTP request can see partial history.
      startupObservation("migration", "enter")
      yield* V1Migration.run().pipe(Effect.provide(context))
      startupObservation("migration", "exit")
      const handler = Context.get(context, HttpRouter.HttpRouter)
        .asHttpEffect()
        .pipe(HttpMiddleware.cors({ allowedOrigins: origin => /^https?:\/\/(localhost|127\.0\.0\.1)(:\d+)?$/.test(origin), credentials: true, maxAge: 86_400 }), HttpEffect.toWebHandlerWith(httpContext))
      const sessions = Context.get(context, Session.Service)
      const locations = Context.get(context, LocationServiceMap.Service)
      const credentials = Context.get(context, Credential.Service)
      // Stateless health requests use the same native executor and HTTP client;
      // models still resolve in the sealed installation Location graph.
      const probeClient = Context.get(yield* Layer.build(LLMClient.layer.pipe(Layer.provide(
        RequestExecutor.layer.pipe(Layer.provide(Layer.succeed(HttpClient.HttpClient, Context.get(context, HttpClient.HttpClient)))),
      ))), LLMClient.Service)
      const restartRelease = yield* Deferred.make<void>()
      const shutdown = yield* Deferred.make<void>()
      let shuttingDown = false
      const stop = () => {
        shuttingDown = true
        Effect.runSync(Deferred.succeed(shutdown, undefined))
      }
      process.on("SIGINT", stop)
      process.on("SIGTERM", stop)
      if (managedRestart) {
        // Keep this process-lifetime handler through teardown: duplicates must
        // neither repeat recovery nor invoke the signal's default termination.
        const releaseRestart = () => {
          if (shuttingDown) return
          Effect.runSync(Deferred.succeed(restartRelease, undefined))
        }
        process.on("SIGUSR2", releaseRestart)
      }
      const reuseCredentials = credentialReuseHandler(credentialReuse({
        directory: nativeCredentialDirectory(), targetDirectory: path.dirname(databasePath),
        create: input => runEngine(createReusedCredential(Context.get(context, Credential.Service), input)),
      }), password)
      const server = Bun.serve({
        // SSE heartbeats and native session waits can exceed Bun's idle default.
        idleTimeout: 0,
        hostname,
        port,
        fetch: (request) => {
          const pathname = new URL(request.url).pathname
          if (pathname === "/api/dochaus/credential-reuse") return reuseCredentials(request)
          if (pathname === "/api/dochaus/ready" && request.method === "GET")
            return managedReadinessResponse(request, locations)
          if (pathname === "/api/dochaus/generate" && request.method === "POST")
            return generateText(request, locations)
          if (pathname === "/api/dochaus/provider-test" && request.method === "POST")
            return testProvider(request, locations, credentials, probeClient)
          if (pathname === "/api/dochaus/provider-catalog" && request.method === "GET")
            return listProviderCatalog(request, locations)
          if (request.method === "POST" && (/^\/api\/session\/[^/]+\/(prompt|generate)$/.test(pathname) ||
              pathname === "/api/experimental/generate" || pathname === "/api/generate"))
            return admitGeneration(request, locations, sessions, handler)
          // /api/plugin stays ungated so it can report activation while it runs.
          if (request.method === "GET" && /^\/api\/(agent|skill|command|model|provider|reference|mcp)(\/|$)/.test(pathname))
            return managedReadinessResponse(request, locations).then(response => response.ok ? handler(request) : response)
          if (!startupObservationEnabled) return pathname === "/api/dochaus/session" && request.method === "POST" ? createChild(request, sessions) : handler(request)
          const phase = pathname === "/api/plugin" ? "api.plugin" : pathname === "/api/config" ? "api.config" : undefined
          const directory = request.headers.get("x-opencode-directory") ?? workspace
          if (phase) startupObservation(phase, "enter", directory)
          const responding = pathname === "/api/dochaus/session" && request.method === "POST"
            ? createChild(request, sessions) : handler(request)
          return Promise.resolve(responding).then(response => {
            if (phase) startupObservation(phase, "exit", directory, { http_status: response.status })
            return response
          }, error => {
            if (phase) startupObservation(phase, "failed", directory)
            throw error
          })
        },
      })
      startupObservation("listener", "exit")
      process.env.OPENCODE_URL = `http://${hostname === "0.0.0.0" || hostname === "::" ? "127.0.0.1" : hostname.includes(":") ? `[${hostname}]` : hostname}:${server.port}`
      yield* Effect.addFinalizer(() => Effect.sync(() => {
        shuttingDown = true
        server.stop(true)
      }))
      const nativeRestart = Context.get(context, SessionRestart.Service).resumeSuspendedSessions
      // Recovery executes saved work outside HTTP admission. Re-read current
      // installation settings when either sweep starts; blocked work stays
      // suspended until explicit recovery after Settings have been repaired.
      const restart = Effect.suspend(() => {
        const error = generationAdmissionError()
        if (!error) return nativeRestart
        console.warn(`Suspended session recovery deferred: ${error} Restart the engine after updating Settings to recover pending work.`)
        return Effect.void
      })
      yield* managedRestart
        ? Effect.forkScoped(Deferred.await(restartRelease).pipe(
            Effect.andThen(Effect.suspend(() => shuttingDown ? Effect.void : restart)),
          ))
        : Effect.forkDetach(restart)
      console.log(`doc.haus engine listening on http://${hostname}:${server.port}`)
      yield* Deferred.await(shutdown)
    }),
  ).pipe(Effect.provide(failureEvidence), Effect.provide(NodeServices.layer), Effect.orDie),
)

async function createChild(request: Request, sessions: Session.Interface): Promise<Response> {
  if (password && request.headers.get("authorization") !== `Basic ${btoa(`opencode:${password}`)}`)
    return Response.json({ error: "Unauthorized" }, { status: 401 })
  const directory = request.headers.get("x-opencode-directory")
  if (!directory) return Response.json({ error: "Matter directory is required" }, { status: 400 })
  try {
    const input = Schema.decodeUnknownSync(ChildInput)(await request.json())
    const parent = await runEngine(sessions.get(input.parentID))
    if (!parent) return Response.json({ error: "Parent session not found" }, { status: 404 })
    if (realpathSync(directory) !== realpathSync(parent.location.directory))
      return Response.json({ error: "Parent session belongs to another matter" }, { status: 403 })
    if (realpathSync(directory) !== realpathSync(workspace) &&
        !realpathSync(directory).startsWith(realpathSync(workspace) + path.sep))
      return Response.json({ error: "Matter is outside the workspace" }, { status: 403 })
    const child = await runEngine(sessions.create(input))
    return Response.json({ data: Schema.encodeSync(SessionSchema.Info)(child) })
  } catch (error) {
    return Response.json({ error: String(error) }, { status: 400 })
  }
}

// A matter activates its plugins in the background after its location boots,
// and the legal agents, skills and commands replace the upstream defaults only
// when that settles. Inventory reads wait for it instead of listing partial state.
async function managedReadinessResponse(request: Request, locations: Context.Service.Shape<typeof LocationServiceMap.Service>, directory?: string) {
  const native = HttpServerRequest.fromWeb(request)
  // Authenticate before resolving a directory, session or Location graph.
  if (password && !(await runEngine(authorizedRequest(native, { username: "opencode", password: Option.some(password) }))))
    return managedJson(request, { error: "Authentication required" }, 401)
  const selected = directory ?? requestRef(native).directory
  const actual = (() => { try { return realpathSync(selected) } catch { return undefined } })()
  const root = realpathSync(workspace)
  if (!actual || actual !== root && !actual.startsWith(root + path.sep))
    return managedJson(request, { ready: false, legalState: "absent", legalPluginCount: 0, consentState: "pending", error: "The requested directory is outside the workspace or unavailable." }, 403)
  try {
    const activated = await runEngine(Effect.gen(function* () {
      yield* Plugin.awaitActivation
      return yield* Plugin.Service.use(plugins => plugins.list())
    }).pipe(Effect.provide(locations.get(Location.Ref.make({ directory: AbsolutePath.make(selected) }))), Effect.timeoutOption("30 seconds")), { signal: request.signal })
    const plugins = Option.isSome(activated) ? activated.value.filter(plugin => plugin.id === "dochaus.legal") : []
    const legalState = Option.isNone(activated) ? "loading" : plugins.some(plugin => plugin.state.status === "failed") ? "failed" : plugins.length === 1 ? "active" : "absent"
    const research = Option.isSome(activated) ? activated.value.filter(plugin => plugin.id === "dochaus.research-mcp") : []
    const consentState = research.some(plugin => plugin.state.status === "failed") ? "failed"
      : Option.isSome(activated) && research.length !== 1 ? "missing" : researchConsentStatus(selected)
    const ready = legalState === "active" && consentState === "active"
    return managedJson(request, { ready, legalState, legalPluginCount: plugins.length, consentState,
      ...(!ready ? { error: legalState !== "active" ? "The legal plugin is missing, failed or still loading. Check the engine log before retrying."
        : "Research consent protection is missing, failed or still loading. Check the engine log before retrying." } : {}),
    }, ready ? 200 : 503)
  } catch {
    return managedJson(request, { ready: false, legalState: "failed", legalPluginCount: 0, consentState: researchConsentStatus(selected), error: "The matter's legal configuration could not activate. Check the engine log." }, 503)
  }
}

async function listProviderCatalog(request: Request, locations: Context.Service.Shape<typeof LocationServiceMap.Service>) {
  const ready = await managedReadinessResponse(request, locations)
  if (!ready.ok) return ready
  try {
    const directory = requestRef(HttpServerRequest.fromWeb(request)).directory
    const data = await runEngine(Provider.Service.use(providers => providers.all()).pipe(
      Effect.map(providers => providerCatalog(providers)),
      Effect.provide(locations.get(Location.Ref.make({ directory }))),
    ), { signal: AbortSignal.any([request.signal, AbortSignal.timeout(10_000)]) })
    return managedJson(request, { data }, 200)
  } catch { return managedJson(request, { error: "Could not discover provider setup. Retry when the engine is ready." }, 503) }
}

async function testProvider(request: Request, locations: Context.Service.Shape<typeof LocationServiceMap.Service>,
  credentials: Credential.Interface, client: Context.Service.Shape<typeof LLMClient.Service>) {
  const ready = await managedReadinessResponse(request, locations)
  if (!ready.ok) return ready
  const input: unknown = await request.json().catch(() => undefined)
  if (!providerTestInput(input))
    return managedJson(request, { ok: false, kind: "model", error: "Choose a catalog provider and model for the connection test." }, 400)
  try {
    const policy = readEffectiveSettings()
    if (Array.isArray(policy.disabled_providers) && policy.disabled_providers.includes(input.providerID) ||
        Array.isArray(policy.enabled_providers) && !policy.enabled_providers.includes(input.providerID))
      return managedJson(request, { ...unavailableModel, error: "This provider is disabled by installation policy." }, 403)
    const ref = Model.Ref.parse(`${input.providerID}/${input.modelID}`)
    const directory = requestRef(HttpServerRequest.fromWeb(request)).directory
    const graph = locations.get(Location.Ref.make({ directory }))
    const fingerprint = await runEngine(Effect.gen(function* () {
      const provider = yield* Provider.Service.use(providers => providers.get(ref.providerID))
      const stored = yield* credentials.all()
      return sha256(JSON.stringify({ directory, ref, policy, credentials: stored.filter(credential =>
        credential.integrationID === (provider?.integrationID ?? ref.providerID)) }))
    }).pipe(Effect.provide(graph)), { signal: AbortSignal.any([request.signal, AbortSignal.timeout(10_000)]) })
    // A disconnected browser must not cancel the shared check for other tabs.
    // The check has its own deadline, and replacing credentials changes the key.
    const result = await cachedProviderTest(fingerprint, input.force === true, () =>
      runEngine(probeProvider(ref).pipe(Effect.provide(graph), Effect.provideService(LLMClient.Service, client)), { signal: AbortSignal.timeout(20_000) }))
    return managedJson(request, result, result.ok ? 200 : result.kind === "model" ? 400 : 503)
  } catch (error) { return managedJson(request, providerTestFailure(error), 503) }
}

function managedJson(request: Request, body: object, status: number) {
  const origin = request.headers.get("origin")
  return Response.json(body, { status, headers: origin && /^https?:\/\/(localhost|127\.0\.0\.1)(:\d+)?$/.test(origin)
    ? { "Access-Control-Allow-Origin": origin, "Access-Control-Allow-Credentials": "true", Vary: "Origin" } : undefined })
}

async function admitGeneration(request: Request, locations: Context.Service.Shape<typeof LocationServiceMap.Service>, sessions: Session.Interface, handler: (request: Request) => Promise<Response>) {
  const native = HttpServerRequest.fromWeb(request)
  if (password && !(await runEngine(authorizedRequest(native, { username: "opencode", password: Option.some(password) }))))
    return managedJson(request, { error: "Authentication required" }, 401)
  const session = /^\/api\/session\/([^/]+)\/(prompt|generate)$/.exec(new URL(request.url).pathname)
  const id = session ? (() => { try { return SessionSchema.ID.make(decodeURIComponent(session[1])) } catch { return undefined } })() : undefined
  const directory = id ? await runEngine(sessions.get(id)).then(info => info.location.directory, () => undefined) : undefined
  // Preserve native decoding and not-found responses without consuming payloads.
  if (session && !directory) return handler(request)
  const ready = await managedReadinessResponse(request, locations, directory)
  if (!ready.ok) return ready
  const error = generationAdmissionError()
  if (error) return managedJson(request, { error, code: "setup_required" }, 503)
  return handler(request)
}

async function generateText(request: Request, locations: Context.Service.Shape<typeof LocationServiceMap.Service>): Promise<Response> {
  if (password && request.headers.get("authorization") !== `Basic ${btoa(`opencode:${password}`)}`)
    return Response.json({ error: "Unauthorized" }, { status: 401 })
  const directory = request.headers.get("x-opencode-directory")
  if (!directory) return Response.json({ error: "Matter directory is required" }, { status: 400 })
  const actual = (() => { try {return realpathSync(directory)} catch {return undefined} })()
  const root = realpathSync(workspace)
  if (!actual || actual === root || !actual.startsWith(root + path.sep))
    return Response.json({ error: "Matter is outside the workspace or unavailable" }, { status: 403 })
  const body:unknown = await request.json().catch(() => undefined)
  if (!body || typeof body !== "object" || Array.isArray(body) || Object.keys(body).join(",") !== "prompt" ||
    !("prompt" in body) || typeof body.prompt !== "string" || !body.prompt.trim() || Buffer.byteLength(body.prompt) > 1_048_576)
    return Response.json({ error: "Provide only a nonempty prompt within the input limit." }, { status: 400 })
  const ready = await managedReadinessResponse(request, locations, actual)
  if (!ready.ok) return ready
  const setupError = generationAdmissionError()
  if (setupError) return managedJson(request, { error: setupError, code: "setup_required" }, 503)
  const trace = { source: "engine_generate_endpoint", directory: actual, prompt: body.prompt, prompt_sha256: sha256(body.prompt), started_at: new Date().toISOString() }
  const began = Date.now()
  try {
    const data = await runEngine(generateMatterText(body.prompt).pipe(
      Effect.provide(locations.get(Location.Ref.make({ directory: AbsolutePath.make(actual) }))),
    ),{signal:request.signal})
    if (tracing) void appendTrace(tracing, "generations", "stateless", { ...trace, model: data.model, text: data.text, elapsed_ms: Date.now() - began })
    return Response.json({ data: { ...data, elapsed_ms: Date.now() - began, usage: null } })
  } catch(error) {
    if (tracing) void appendTrace(tracing, "generations", "stateless", { ...trace, error: describeError(error), elapsed_ms: Date.now() - began })
    console.warn("[matter-generation] configured generation failed",diagnosticErrorMetadata(error))
    return Response.json({ error: "The configured legal generation could not complete. No alternative model was used." }, { status: 503 })
  }
}
