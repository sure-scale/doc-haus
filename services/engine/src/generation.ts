import { Generate } from "@opencode/core/generate"
import { Plugin } from "@opencode/core/plugin/service"
import { Model } from "@opencode/schema/model"
import { Effect } from "effect"
import { readEffectiveSettings } from "./effective-settings"

// Resolve after native plugin activation, in the active matter's sealed graph.
// An unavailable configured model must fail rather than choose another catalog item.
export const generateMatterText = Effect.fn("dochaus.generateMatterText")(function* (prompt: string) {
  yield* Plugin.awaitActivation
  const selected = readEffectiveSettings().model
  if (!selected) return yield* Effect.fail(new Error("The legal configuration has no selected generation model."))
  const model = Model.Ref.parse(selected)
  const text = yield* Generate.Service.use((generate) => generate.text({ prompt, model }))
  return { text, model }
})
