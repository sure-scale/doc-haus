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
import { DOCUMENT_AMENDMENT_LEGAL_OUTPUT, documentAmendmentLegalJsonSchema, decodeDocumentAmendmentLegal } from "../../../dochaus/lib/document-amendment-legal-contract"

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
const generateWithUsage = Effect.fn("dochaus.generateWithUsage")(function* (prompt: string, ref: Model.Ref, outputContract?: typeof DOCUMENT_AMENDMENT_LEGAL_OUTPUT) {
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
  if (outputContract === DOCUMENT_AMENDMENT_LEGAL_OUTPUT) {
    const result = yield* ai.generateObject({
      model: resolved.model, prompt, jsonSchema: documentAmendmentLegalJsonSchema,
      http: { headers: { "x-opencode-session": SessionID.create() } },
    }).pipe(ModelRetry.retryOneShot, Effect.result)
    if (result._tag === "Failure") {
      // A schema refusal can still have incurred a provider charge. Preserve
      // partial usage before returning a controlled failure; missing usage is
      // explicitly unknown, never a fabricated zero-cost successful result.
      return { error: { code: result.failure.reason._tag === "InvalidProviderOutput" ? "invalid_structured_output" as const : "structured_generation_failed" as const }, usage: generationUsage(result.failure.partial?.usage, resolved.cost) }
    }
    const usage = generationUsage(result.success.usage, resolved.cost)
    try { return { text: JSON.stringify(decodeDocumentAmendmentLegal(result.success.object)), usage } }
    catch { return { error: { code: "invalid_structured_output" as const }, usage } }
  }
  const response = yield* LLMClient.Service.use((client) => client.generate(ai.request({
    model: resolved.model,
    prompt,
    // Gateways require session attribution even for a stateless call; no Session is stored.
    http: { headers: { "x-opencode-session": SessionID.create() } },
  }))).pipe(
    ModelRetry.retryOneShot,
    Effect.mapError((error: AIError) => new Generate.UnavailableError({ message: error.message, service: resolved.ref.providerID })),
  )
  return { text: response.text, usage: generationUsage(response.usage, resolved.cost) }
}, (effect) => effect.pipe(Effect.catchTag("Integration.Authorization",
  () => new Generate.UnavailableError({ message: "Generation credentials are unavailable" }))))

// Resolve after native plugin activation, in the active matter's sealed graph.
// An unavailable configured model must fail rather than choose another catalog item.
export const generateMatterText = Effect.fn("dochaus.generateMatterText")(function* (prompt: string, outputContract?: typeof DOCUMENT_AMENDMENT_LEGAL_OUTPUT, modelRole?: "fast") {
  yield* Plugin.awaitActivation
  const settings = readEffectiveSettings()
  const selected = modelRole === "fast" ? settings.small_model : settings.model
  if (!selected) return yield* Effect.fail(new Error(`The legal configuration has no selected ${modelRole === "fast" ? "fast " : ""}generation model.`))
  const model = Model.Ref.parse(selected)
  const output = yield* generateWithUsage(prompt, model, outputContract)
  return { ...output, model }
})

function generationUsage(usage: Parameters<typeof SessionUsage.record>[0], costs: Parameters<typeof SessionUsage.record>[1]): GenerationUsage {
  const tokens = SessionUsage.tokens(usage)
  const count = (value: unknown): value is number => typeof value === "number" && Number.isFinite(value) && value >= 0
  if (!usage) return { tokens, cost: null }
  // A reported charge (including an authoritative zero) needs no token estimate.
  if (usage.cost !== undefined) return { tokens, cost: count(usage.cost) ? usage.cost : null }
  // Usage is an optional-counter object. Presence alone does not mean that the
  // native normalizer's zero defaults are a complete billable token report.
  const counters = [usage.cacheReadInputTokens, usage.cacheWriteInputTokens, usage.reasoningTokens]
  if (!count(usage.inputTokens) || !count(usage.outputTokens) || !count(usage.nonCachedInputTokens) ||
      counters.some(value => value !== undefined && !count(value)) ||
      usage.nonCachedInputTokens + (usage.cacheReadInputTokens ?? 0) + (usage.cacheWriteInputTokens ?? 0) !== usage.inputTokens ||
      (usage.reasoningTokens ?? 0) > usage.outputTokens)
    return { tokens, cost: null }
  // Match native catalog tier selection before using its price calculator; an
  // absent/unknown rate is not evidence of a zero-price model request either.
  const rate = costs.filter(cost => cost.tier?.type === "context" && usage.inputTokens! > cost.tier.size)
    .toSorted((a, b) => (b.tier?.size ?? 0) - (a.tier?.size ?? 0))[0] ?? costs.find(cost => cost.tier === undefined)
  if (!rate || ![rate.input, rate.output, rate.cache.read, rate.cache.write].every(count)) return { tokens, cost: null }
  return { tokens, cost: SessionUsage.record(usage, costs).cost }
}
