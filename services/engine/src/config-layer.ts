import { Config } from "@opencode/core/config"
import { Directory, Document } from "@opencode/schema/config"
import { AbsolutePath } from "@opencode/core/schema"
import { Model } from "@opencode/schema/model"
import { Context, Effect, Layer, Stream } from "effect"
import { AGENT_TIER, resolveTierModel } from "../../../dochaus/lib/model-tiers"
import { startupObservation } from "../../../dochaus/lib/startup-observation"
import { readEffectiveSettings, settingsPaths } from "./effective-settings"

export function legalConfigLayer(legal: string, paths = settingsPaths()) {
  return Layer.effect(
    Config.Service,
    Effect.gen(function* () {
      // Capture each service independently; one Config tag must not replace the other.
      const shippedContext = yield* Layer.build(Config.layer({ global: false, project: false, file: paths.shipped }))
      const overrideContext = yield* Layer.build(Config.layer({ global: false, project: false, file: paths.overrides }))
      const shipped = Context.get(shippedContext, Config.Service)
      const overrides = Context.get(overrideContext, Config.Service)
      return Config.Service.of({
        // No update method: global native writes are intentionally unavailable.
        compatibility: () => Effect.succeed({ claude: [], agents: [] }),
        changes: () => Stream.merge(shipped.changes(), overrides.changes()),
        entries: () => Effect.gen(function* () {
          startupObservation("config", "enter")
          const entries = [...(yield* shipped.entries()), ...(yield* overrides.entries())]
          const policy = readEffectiveSettings(paths)
          const agents = Object.fromEntries(Object.entries(AGENT_TIER).flatMap(([id, tier]) => {
            if (policy.agent?.[id]?.model) return []
            const selected = resolveTierModel(tier, policy)
            if (!selected) return []
            const model = Model.Ref.parse(selected)
            return [[id, { model: { providerID: model.providerID, model: model.id, ...(model.variant ? { variant: model.variant } : {}) } }]]
          }))
          return [
            ...entries,
            new Directory({ type: "directory", path: AbsolutePath.make(legal) }),
            ...(Object.keys(agents).length ? [new Document({ type: "document", info: { agents } })] : []),
          ]
        }).pipe(Effect.onExit(exit => Effect.sync(() => startupObservation("config", exit._tag === "Success" ? "exit" : "failed")))),
      })
    }),
  )
}
