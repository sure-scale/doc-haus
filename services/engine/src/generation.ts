import { ai, LLMClient, type AIError } from "@opencode/ai/effect"
import { Generate } from "@opencode/core/generate"
import { ModelResolver } from "@opencode/core/model-resolver"
import { ModelRetry } from "@opencode/core/model-retry"
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
// resolution and one-shot retry as core Generate.text (selected model always
// given), but its response is not reduced to text, so stateless jobs can
// report what they spent. No session is stored. Failures use core Generate
// error classes; callers report only a generic refusal.
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
      (error) => Effect.fail(
        error._tag === "SessionRunnerModel.VariantUnavailableError"
          ? new Generate.InvalidRequestError({ message: error.message, field: "model.variant" })
          : error._tag === "SessionRunnerModel.UnsupportedPackageError"
            ? new Generate.UnsupportedModelError({
              providerID: ref.providerID,
              modelID: ref.id,
              capability: "chat",
              message: `${ref.providerID}/${ref.id} has no chat API opencode can call (package ${error.package})`,
            })
            : new Generate.InvalidRequestError({ message: error.message, field: "model" })),
    ),
  )
  if (!resolved)
    return yield* new Generate.ModelNotFoundError({
      providerID: ref.providerID,
      modelID: ref.id,
      message: `Model not found: ${ref.providerID}/${ref.id}`,
    })
  const response = yield* LLMClient.Service.use((client) => client.generate(ai.request({
    model: resolved.model,
    prompt,
    // Gateways require session attribution even for a stateless call; no Session is stored.
    http: { headers: { "x-opencode-session": SessionID.create() } },
  }))).pipe(
    ModelRetry.retryOneShot,
    Effect.mapError((error: AIError) => new Generate.UnavailableError({ message: error.message, service: resolved.ref.providerID })),
  )
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
