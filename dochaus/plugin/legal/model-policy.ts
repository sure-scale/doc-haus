import { assignedThinkingOf } from "../../lib/model-tiers"

export function capLowThroughputModelOutput(providerID: string, output: { maxOutputTokens: number | undefined }) {
  if (!lowThroughputProviders().has(providerID)) return
  const cap = lowThroughputOutputCap()
  output.maxOutputTokens = output.maxOutputTokens === undefined ? cap : Math.min(output.maxOutputTokens, cap)
}

export function adaptAzureChatCompletionGpt5Params(
  providerID: string,
  model: string,
  output: { maxOutputTokens: number | undefined; options: Record<string, unknown> },
) {
  if (!providerID.toLowerCase().includes("azure")) return
  if (!model.toLowerCase().includes("gpt-5")) return
  output.maxOutputTokens = undefined
  delete output.options.reasoningEffort
  delete output.options.reasoningSummary
  delete output.options.include
  delete output.options.textVerbosity
}

export function adaptOpenAIPlanParams(
  providerID: string,
  output: { maxOutputTokens: number | undefined },
) {
  if (providerID !== "openai") return
  // OpenAI OAuth/Codex-plan models choose their output budget server-side and
  // reject max_output_tokens. Run this after doc.haus role caps so a later legal
  // policy hook cannot accidentally reintroduce the unsupported parameter.
  output.maxOutputTokens = undefined
}

// Apply the agent's tier thinking budget using each model family's supported
// control: Gemini 3 gets a thinkingLevel; models exposing the common
// `enable_thinking` template switch stay on for medium/high judgment roles and
// off for bounded execution roles. Set both request dialects where applicable;
// each provider ignores the option it does not consume.
// Only explicitly tiered roles are touched. The matching model per
// tier is bound by the engine's config projection from the same live configuration.
export function applyAgentThinkingBudget(agent: string, model: string, output: { options: Record<string, any> }) {
  const level = assignedThinkingOf(agent)
  if (!level) return
  if (supportsBooleanThinking(model)) {
    output.options.chat_template_kwargs = {
      ...(output.options.chat_template_kwargs ?? {}),
      enable_thinking: level === "medium" || level === "high",
    }
    if (level === "medium" || level === "high") {
      output.options.thinkingConfig = {
        ...(output.options.thinkingConfig ?? {}),
        includeThoughts: true,
        thinkingLevel: level,
      }
      return
    }
    output.options.thinkingConfig = { includeThoughts: false, thinkingLevel: "minimal" }
    return
  }
  if (!model.toLowerCase().includes("gemini-3")) return
  output.options.thinkingConfig = {
    ...(output.options.thinkingConfig ?? {}),
    thinkingLevel: compatibleGeminiThinkingLevel(model, level),
  }
}

// Auto routing is a classification turn whose entire useful response is one
// short assistant id. Keep the user's selected model, but do not pay that
// model's normal deliberation and output budgets before the substantive legal
// turn begins. Provider-specific no-thinking controls are applied only where
// their request dialect is known. Only non-reasoning models get a visible-output
// cap; reasoning models retain the engine-resolved total response allowance.
export function optimizeRouterParams(
  agent: string,
  model: string,
  output: { temperature: number | undefined; maxOutputTokens: number | undefined; options: Record<string, any> },
  reasoning: boolean,
) {
  if (agent !== "router") return
  output.temperature = 0
  if (!reasoning) output.maxOutputTokens = Math.min(output.maxOutputTokens ?? 16, 16)
  if (supportsBooleanThinking(model)) {
    output.options.chat_template_kwargs = { ...(output.options.chat_template_kwargs ?? {}), enable_thinking: false }
    output.options.thinkingConfig = { includeThoughts: false, thinkingLevel: "minimal" }
    return
  }
  if (!model.toLowerCase().includes("gemini")) return
  output.options.thinkingConfig = model.toLowerCase().includes("gemini-3")
    ? { includeThoughts: false, thinkingLevel: compatibleGeminiThinkingLevel(model, "minimal") }
    : { includeThoughts: false, thinkingBudget: 0 }
}

function lowThroughputProviders() {
  return new Set(
    (process.env.DOCHAUS_LOW_THROUGHPUT_PROVIDERS ?? "azure-foundry")
      .split(",")
      .map((provider) => provider.trim())
      .filter(Boolean),
  )
}

function lowThroughputOutputCap() {
  const value = Number(process.env.DOCHAUS_LOW_TPM_MAX_OUTPUT_TOKENS ?? "4096")
  if (!Number.isFinite(value) || value < 512) return 4096
  return Math.floor(value)
}

function supportsBooleanThinking(model: string) {
  const value = model.toLowerCase()
  return value.includes("gemma") || value.includes("qwen") || value.includes("bonsai")
}

// Gemini 3 Pro rejects the family's `minimal` level while Flash accepts it.
// Raise only the unsupported floor; preserve every substantive low/medium/high
// assignment so the role policy remains model-independent.
function compatibleGeminiThinkingLevel(model: string, level: string) {
  if (level !== "minimal") return level
  return model.toLowerCase().includes("pro") ? "low" : level
}

const lowThroughputQueues = new Map<string, Promise<void>>()
const lowThroughputNextAt = new Map<string, number>()

export async function paceLowThroughputModel(providerID: string, modelID: string) {
  if (!lowThroughputProviders().has(providerID)) return
  const interval = lowThroughputMinIntervalMs()
  if (interval <= 0) return
  const key = `${providerID}/${modelID}`
  const previous = lowThroughputQueues.get(key) ?? Promise.resolve()
  const next = previous.then(async () => {
    const wait = Math.max(0, (lowThroughputNextAt.get(key) ?? 0) - Date.now())
    if (wait) await sleep(wait)
    lowThroughputNextAt.set(key, Date.now() + interval)
  })
  lowThroughputQueues.set(key, next.catch(() => {}))
  await next
}

function lowThroughputMinIntervalMs() {
  const value = Number(process.env.DOCHAUS_LOW_TPM_MIN_INTERVAL_MS ?? "65000")
  if (!Number.isFinite(value) || value < 0) return 65_000
  return Math.floor(value)
}

export function modelID(model: { modelID?: string; id?: string }) {
  return model.modelID ?? model.id ?? "unknown"
}

function sleep(ms: number) {
  return new Promise((resolve) => setTimeout(resolve, ms))
}
