import { applyAgentThinkingBudget, modelID, optimizeRouterParams } from "./model-policy"
import type { SessionRequestKind } from "@opencode/plugin/promise/session"

function record(value: unknown): Record<string, unknown> {
  return value !== null && typeof value === "object" && !Array.isArray(value) ? value as Record<string, unknown> : {}
}

export function declaredModelReasoning(config: unknown, model: { providerID: string; id: string }) {
  const provider = record(record(record(config).provider)[model.providerID])
  const declared = record(record(provider.models)[model.id]).reasoning
  return typeof declared === "boolean" ? declared : true
}

export function nativeAgentSampling(body: unknown) {
  const values = record(body)
  return {
    ...(typeof values.temperature === "number" && Number.isFinite(values.temperature) ? { temperature: values.temperature } : {}),
    ...(typeof values.top_p === "number" && Number.isFinite(values.top_p) ? { topP: values.top_p } : {}),
  }
}

// A model may forbid sampling controls even when an agent declares them.
// Honor the selected model's explicit capability declaration after role policy.
export function applyDeclaredModelSampling(config: unknown, model: { providerID: string; id: string }, options: Record<string, unknown>) {
  const provider = record(record(record(config).provider)[model.providerID])
  if (record(record(provider.models)[model.id]).temperature !== false) return
  delete options.temperature
  delete options.topP
  delete options.topK
}

// Native compatible transport omits provider template options. Recompute the
// same pure legal policy at the HTTP boundary, without request-local caches.
export async function compatibleThinkingRequest(request: Request, input: {
  directory: string; sessionID: string; agent: string; reasoning: boolean
  kind?: SessionRequestKind
  model: { providerID: string; id: string; modelID?: string; package?: string }
}) {
  if (!["@opencode/ai/providers/openai-compatible", "@ai-sdk/openai-compatible"].includes(input.model.package ?? "")) return request
  const body = record(await request.clone().json())
  const output = { options: {} as Record<string, unknown>, temperature: undefined as number | undefined, maxOutputTokens: undefined as number | undefined }
  const activeModel = modelID(input.model)
  applyAgentThinkingBudget(input.agent, activeModel, output)
  optimizeRouterParams(input.agent, activeModel, output, input.reasoning)
  if (!output.options.chat_template_kwargs) return request
  const headers = new Headers(request.headers)
  headers.delete("content-length")
  return new Request(request, { headers, body: JSON.stringify({ ...body, chat_template_kwargs: { ...record(body.chat_template_kwargs), ...record(output.options.chat_template_kwargs) } }) })
}
