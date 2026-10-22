import { NodeServices } from "@effect/platform-node"
import { Bus } from "@opencode/core/bus"
import { Config } from "@opencode/core/config"
import { Credential } from "@opencode/core/credential"
import { Database } from "@opencode/core/database/database"
import { AppNodeBuilder } from "@opencode/core/effect/app-node-builder"
import { Watcher } from "@opencode/core/filesystem/watcher"
import { Instance } from "@opencode/core/instance"
import { InstructionDiscovery } from "@opencode/core/instruction-discovery"
import { Location } from "@opencode/core/location"
import { ManagedPolicy } from "@opencode/core/managed-policy"
import { ModelsDev } from "@opencode/core/models-dev"
import { Plugin } from "@opencode/core/plugin"
import { AbsolutePath } from "@opencode/core/schema"
import { WellKnown } from "@opencode/core/wellknown"
import { Integration } from "@opencode/schema/integration"
import { makeLocationNode } from "@opencode/util/effect/app-node"
import { LayerNode } from "@opencode/util/effect/layer-node"
import { FSUtil } from "@opencode/util/fs-util"
import { Global } from "@opencode/util/global"
import { Context, Effect, Layer } from "effect"
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { legalConfigLayer, legalManagedPolicy } from "./config-layer"
import type { EngineSettings } from "./effective-settings"
import { modelHttp } from "./model-http"

export type SetupCredential = { integrationID: Integration.ID; label: string; value: Credential.Value }

// Native Integration owns OAuth callback listeners, cancellation and command
// attempts. Its Credential service is replaced before acquisition, so completing
// those attempts cannot activate a connection in the installation database.
export function setupMemoryCredentials(initial?: SetupCredential) {
  const records = new Map<Credential.ID, Credential.Info>()
  let notify = (_info: Credential.Info, _switched: boolean): Effect.Effect<void> => Effect.void
  const snapshot = (integrationID?: string) => {
    const value = Array.from(records.values()).findLast(record => !integrationID || record.integrationID === integrationID)
    return value && { integrationID: value.integrationID, label: value.label, value: structuredClone(value.value) }
  }
  const service = Credential.Service.of({
    all: () => Effect.sync(() => Array.from(records.values())),
    list: integrationID => Effect.sync(() => Array.from(records.values()).filter(record => record.integrationID === integrationID)),
    get: id => Effect.sync(() => records.get(id)),
    create: input => Effect.gen(function* () {
      const info = new Credential.Info({ ...input, id: Credential.ID.create(), label: input.label ?? "default", value: structuredClone(input.value) })
      records.set(info.id, info)
      yield* notify(info, true)
      return info
    }),
    activate: id => Effect.gen(function* () {
      const info = records.get(id)
      if (!info) throw new Error("Setup connection is unavailable.")
      records.delete(id)
      records.set(id, info)
      yield* notify(info, true)
    }),
    update: (id, updates) => Effect.gen(function* () {
      const info = records.get(id)
      if (!info) throw new Error("Setup connection is unavailable.")
      records.set(id, new Credential.Info({ ...info, ...updates, value: structuredClone(updates.value ?? info.value) }))
      if (updates.label !== undefined && updates.label !== info.label) yield* notify(info, false)
    }),
    remove: id => Effect.gen(function* () {
      const info = records.get(id)
      records.delete(id)
      if (info) yield* notify(info, true)
    }),
  })
  if (initial) Effect.runSync(service.create(initial))
  return { service, snapshot, setNotify: (callback: typeof notify) => { notify = callback } }
}

export function createProviderSetupSession(options: {
  legal: string
  workspace?: string
  settings: EngineSettings
  credential?: SetupCredential
}) {
  return Effect.gen(function* () {
    const directory = yield* Effect.acquireRelease(
      Effect.sync(() => mkdtempSync(path.join(tmpdir(), "dochaus-provider-setup-"))),
      directory => Effect.sync(() => rmSync(directory, { recursive: true, force: true })),
    )
    const workspace = path.join(directory, "workspace")
    const configDirectory = path.join(directory, "config")
    mkdirSync(workspace, { recursive: true })
    mkdirSync(configDirectory, { recursive: true, mode: 0o700 })
    const overrides = path.join(configDirectory, "opencode.json")
    writeFileSync(overrides, JSON.stringify(options.settings), { mode: 0o600 })
    const memory = setupMemoryCredentials(options.credential)
    const config = makeLocationNode({
      service: Config.Service,
      layer: legalConfigLayer(options.legal, { shipped: path.join(options.legal, "opencode.json"), overrides }),
      deps: [Watcher.node, Bus.node, FSUtil.node, Global.node, Location.node, Credential.node, WellKnown.node],
    })
    const built = AppNodeBuilder.build(LayerNode.group([Instance.graph, Bus.node, Credential.node]), [
      modelHttp,
      Location.node.replace(Location.boundNode(Location.Ref.make({ directory: AbsolutePath.make(workspace) }))),
      Global.node.replace(Global.layerWith({
        home: directory, tmp: path.join(directory, "tmp"),
        data: path.join(directory, "data"), config: configDirectory, state: path.join(directory, "state"),
        cache: path.join(directory, "cache"), bin: path.join(directory, "bin"), log: path.join(directory, "log"), repos: path.join(directory, "repos"),
      })),
      Database.node.replace(Database.configured({ path: path.join(directory, "setup.db") })),
      Credential.node.replace(Layer.succeed(Credential.Service, memory.service)),
      ModelsDev.node.replace(ModelsDev.configured({ fetch: false })),
      Watcher.node.replace(Watcher.configured({ enabled: false })),
      WellKnown.node.replace(Layer.succeed(WellKnown.Service, {
        entries: () => Effect.succeed([]), snapshot: () => [], refresh: () => Effect.succeed(false),
        add: () => Effect.fail(new Error("Remote configuration discovery is disabled during setup")),
        remove: () => Effect.void, resolve: () => Effect.succeed([]),
      })),
      Config.node.replace(config),
      ManagedPolicy.node.replace(legalManagedPolicy),
      InstructionDiscovery.node.replace(InstructionDiscovery.configured({ global: false, project: false })),
    ]).pipe(Layer.provide(NodeServices.layer))
    const context = yield* Layer.build(built)
    const bus = Context.get(context, Bus.Service)
    memory.setNotify((info, switched) => Effect.gen(function* () {
      yield* bus.publish(Credential.Event.Updated, {}, { global: true })
      if (switched) yield* bus.publish(Credential.Event.Switched, {
        integrationID: info.integrationID,
        credentialID: Array.from(yield* memory.service.list(info.integrationID)).at(-1)?.id ?? null,
      }, { global: true })
    }))
    yield* Plugin.awaitActivation.pipe(Effect.provide(context))
    return {
      context,
      directory,
      credentials: memory.service,
      credential: memory.snapshot,
      provide: <A, E, R>(effect: Effect.Effect<A, E, R>) =>
        effect.pipe(Effect.provide(Context.add(context, Credential.Service, memory.service))),
    }
  })
}

export type ProviderSetupSession = Effect.Success<ReturnType<typeof createProviderSetupSession>>
