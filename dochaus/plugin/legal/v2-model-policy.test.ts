import { expect, test } from "bun:test"
import { mkdtempSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { applyDeclaredModelSampling, compatibleThinkingRequest, declaredModelReasoning, nativeAgentSampling } from "./v2-model-policy"

test("declared reasoning follows selected alias and preserves explicit false", () => {
  const config = { provider: { local: { models: { alias: { reasoning: false } } } } }
  expect(declaredModelReasoning(config, { providerID: "local", id: "alias" })).toBe(false)
  expect(declaredModelReasoning(config, { providerID: "local", id: "wire-name" })).toBe(true)
  expect(nativeAgentSampling({ temperature: 0, top_p: 0.8, unrelated: true })).toEqual({ temperature: 0, topP: 0.8 })
  expect(nativeAgentSampling({ temperature: false, top_p: NaN })).toEqual({})
})

test("explicit unsupported sampling omits agent controls and preserves other options", () => {
  const options = { temperature: 0.1, topP: 0.8, topK: 20, maxTokens: 4096, reasoningEffort: "high" }
  const model = { providerID: "local", id: "selected" }
  applyDeclaredModelSampling({}, model, options)
  expect(options.temperature).toBe(0.1)
  applyDeclaredModelSampling({ provider: { local: { models: { selected: { temperature: false } } } } }, model, options)
  expect(options).toEqual({ maxTokens: 4096, reasoningEffort: "high" })
})

test("compatible request retains body and headers while applying role thinking control", async () => {
  const directory = mkdtempSync(path.join(tmpdir(), "dochaus-native-options-"))
  try {
    const request = new Request("https://provider.invalid/chat/completions", { method: "POST", headers: { "content-type": "application/json", "x-request-test": "preserved" }, body: JSON.stringify({ model: "qwen3", messages: [{ role: "user", content: "hello" }], chat_template_kwargs: { other: true, enable_thinking: true } }) })
    const input = { directory, sessionID: "ses_test", agent: "draft", reasoning: true, model: { providerID: "local", id: "alias", modelID: "qwen3", package: "@opencode/ai/providers/openai-compatible" } }
    const mapped = await compatibleThinkingRequest(request, input)
    expect(await mapped.json()).toEqual({ model: "qwen3", messages: [{ role: "user", content: "hello" }], chat_template_kwargs: { other: true, enable_thinking: false } })
    expect(mapped.headers.get("x-request-test")).toBe("preserved")
    expect(request.bodyUsed).toBe(false)
    expect(await compatibleThinkingRequest(request, { ...input, model: { ...input.model, package: "@opencode/ai/providers/google" } })).toBe(request)
  } finally {
    rmSync(directory, { recursive: true, force: true })
  }
})
