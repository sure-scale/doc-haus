import { LLM, LLMClient, type AIError } from "@opencode/ai"
import { Generate } from "@opencode/core/generate"
import { ModelResolver } from "@opencode/core/model-resolver"
import { Plugin } from "@opencode/core/plugin/service"
import { SessionUsage } from "@opencode/core/session/usage"
import { Model } from "@opencode/schema/model"
import { SessionID } from "@opencode/schema/session-id"
import { Effect } from "effect"
import { readEffectiveSettings } from "./effective-settings"

export type GenerationUsage = {
  tokens: { input: number; output: number; reasoning: number; cache: { read: number; write: number } }
  // Catalog-priced dollars; null when the provider reported no usage.
  cost: number | null
}

// Generate.text with the provider's usage kept. Same request shape, model
// resolution and error mapping as core Generate.text (selected model always
// given), but its response is not reduced to text, so stateless jobs can
// report what they spent. No session is stored.
const generateWithUsage = Effect.fn("dochaus.generateWithUsage")(function* (prompt: string, ref: Model.Ref) {
  const resolved = yield* ModelResolver.Service.use((resolver) => resolver.resolve(ref)).pipe(
    Effect.catchTag(
      [
        "SessionRunnerModel.VariantUnavailableError",
        "SessionRunnerModel.UnsupportedPackageError",
        "SessionRunnerModel.ModelConfigurationError",
        "SessionRunnerModel.ModelInitializationError",
        "SessionRunnerModel.UnresolvedProviderVariablesError",
        "SessionRunnerModel.UnsupportedCompactionError",
      ],
      (error) => Effect.fail(new Generate.ModelSelectionError({ message: error.message })),
    ),
  )
  if (!resolved)
    return yield* new Generate.ModelSelectionError({ message: `Model unavailable: ${ref.providerID}/${ref.id}` })
  const response = yield* LLMClient.Service.use((client) => client.generate(LLM.request({
    model: resolved.model,
    prompt,
    // Gateways require session attribution even for a stateless call; no Session is stored.
    http: { headers: { "x-opencode-session": SessionID.create() } },
  }))).pipe(Effect.mapError((error: AIError) => new Generate.UnavailableError({ message: error.message, service: resolved.ref.providerID })))
  const recorded = SessionUsage.record(response.usage, resolved.cost)
  const usage: GenerationUsage = { tokens: recorded.tokens, cost: response.usage ? recorded.cost : null }
  return { text: response.text, usage }
}, (effect) => effect.pipe(Effect.catchTag("Integration.Authorization",
  () => new Generate.UnavailableError({ message: "Generation credentials are unavailable" }))))

// Resolve after native plugin activation, in the active matter's sealed graph.
// An unavailable configured model must fail rather than choose another catalog item.
export const generateMatterText = Effect.fn("dochaus.generateMatterText")(function* (prompt: string) {
  yield* Plugin.awaitActivation
  const selected = readEffectiveSettings().model
  if (!selected) return yield* Effect.fail(new Error("The legal configuration has no selected generation model."))
  const model = Model.Ref.parse(selected)
  const { text, usage } = yield* generateWithUsage(prompt, model)
  return { text, model, usage }
})
