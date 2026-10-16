import { expect, test } from "bun:test"
import { mkdtempSync, mkdirSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { Config } from "@opencode/core/config"
import { AppNodeBuilder } from "@opencode/core/effect/app-node-builder"
import { Location } from "@opencode/core/location"
import { Credential } from "@opencode/core/credential"
import { WellKnown } from "@opencode/core/wellknown"
import { Watcher } from "@opencode/core/filesystem/watcher"
import { Bus } from "@opencode/core/bus"
import { Global } from "@opencode/util/global"
import { AbsolutePath } from "@opencode/core/schema"
import { Effect, Layer, Stream } from "effect"
import { NodeServices } from "@effect/platform-node"
import { makeLocationNode, makeGlobalNode } from "@opencode/util/effect/app-node"
import { FSUtil } from "@opencode/util/fs-util"
import { legalConfigLayer } from "./config-layer"
import { Project } from "@opencode/core/project"
const emptyCredentialNode = makeGlobalNode({ service: Credential.Service, layer: Layer.succeed(Credential.Service, Credential.Service.of({
  all: () => Effect.succeed([]), list: () => Effect.succeed([]), get: () => Effect.undefined,
  create: () => Effect.die("unused"), activate: () => Effect.die("unused"), update: () => Effect.die("unused"), remove: () => Effect.die("unused"),
})), deps: [] })
const emptyWellknownNode = makeGlobalNode({ service: WellKnown.Service, layer: Layer.succeed(WellKnown.Service, {
  entries: () => Effect.succeed([]), snapshot: () => [], refresh: () => Effect.succeed(false), add: () => Effect.die("unused"), remove: () => Effect.void, resolve: () => Effect.succeed([]),
}), deps: [] })

test("independent native config services retain policy, overrides and both watch sources", async () => {
  const root = mkdtempSync(path.join(tmpdir(), "legal-config-layer-"))
  try {
    const legal = path.join(root, "legal"), runtime = path.join(root, "runtime"), matter = path.join(root, "matter")
    for (const directory of [legal, runtime, matter]) mkdirSync(directory)
    const paths = { shipped: path.join(legal, "opencode.json"), overrides: path.join(runtime, "opencode.json") }
    writeFileSync(paths.shipped, JSON.stringify({ default_agent: "draft", permission: { bash: "deny" }, agent: { draft: { description: "Draft" } } }))
    writeFileSync(paths.overrides, JSON.stringify({ model: "test/primary", small_model: "test/fast", agent: { narrator: { model: "test/explicit" } } }))
    writeFileSync(path.join(matter, "opencode.json"), JSON.stringify({ model: "untrusted/model" }))
    const node = makeLocationNode({ service: Config.Service, layer: legalConfigLayer(legal, paths), deps: [Watcher.node, Bus.node, FSUtil.node, Global.node, Location.node, Credential.node, WellKnown.node] })
    const built = AppNodeBuilder.build(node, [
      Location.node.replace(Layer.succeed(Location.Service, Location.Service.of({ directory: AbsolutePath.make(matter), project: { id: Project.ID.global, directory: AbsolutePath.make(matter), canonical: AbsolutePath.make(matter) } }))),
      Global.node.replace(Global.layerWith({ config: path.join(root, "global"), home: root, data: path.join(root, "data"), state: path.join(root, "state"), cache: path.join(root, "cache"), bin: path.join(root, "bin"), log: path.join(root, "log"), repos: path.join(root, "repos") })),
      Credential.node.replace(emptyCredentialNode), WellKnown.node.replace(emptyWellknownNode), Watcher.node.replace(Watcher.testLayer),
    ])
    await Effect.runPromise(Effect.scoped(Effect.gen(function* () {
      const config = yield* Config.Service
      const entries = yield* config.entries()
      expect(entries.filter(entry => entry.type === "document").map(entry => String(entry.path)).filter(p => p !== "undefined")).toEqual([paths.shipped, paths.overrides])
      expect(entries.some(entry => entry.type === "directory" && entry.path === legal)).toBe(true)
      expect(String(Config.latest(entries, "model")?.model)).toBe("primary")
      expect(String(Config.latest(entries, "agents")?.draft.model?.model)).toBe("fast")
      expect(entries.some(entry => entry.type === "document" && entry.info.agents?.narrator?.model?.model === "explicit")).toBe(true)
      expect(config.update).toBeUndefined()
      expect(yield* config.compatibility!()).toEqual({ claude: [], agents: [] })
      const watcher = yield* Watcher.Test
      const subscriptions = yield* watcher.subscriptions()
      expect(subscriptions.some(watch => watch.path === legal)).toBe(true)
      expect(subscriptions.some(watch => watch.path === runtime)).toBe(true)
      const updates: string[] = []
      yield* Stream.runForEach(config.changes(), update => Effect.sync(() => { updates.push(update.path) })).pipe(Effect.forkScoped({ startImmediately: true }))
      yield* Effect.sleep("10 millis")
      yield* watcher.emit({ type: "update", path: paths.shipped })
      yield* Effect.sync(() => writeFileSync(paths.overrides, JSON.stringify({ model: "test/revised", small_model: "test/fast" })))
      yield* watcher.emit({ type: "update", path: paths.overrides })
      yield* Effect.sleep("10 millis")
      expect(updates).toContain(paths.shipped); expect(updates).toContain(paths.overrides)
      yield* Effect.sleep("150 millis")
      expect(String(Config.latest(yield* config.entries(), "model")?.model)).toBe("revised")
    }).pipe(Effect.provide(Layer.mergeAll(built, Watcher.testLayer)), Effect.provide(NodeServices.layer))))
  } finally { rmSync(root, { recursive: true, force: true }) }
})
