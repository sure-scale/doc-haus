import { INGEST_URL } from "../../config"

export type Config = {
  model?: string
  small_model?: string
  disabled_providers?: string[]
  provider?: Record<string, { npm?: string; name?: string; options?: Record<string, unknown>; models?: Record<string, { name: string }> }>
}

// Ingest writes the explicit legal config; native config watching reloads it.
export async function getConfig() {
  const res = await fetch(`${INGEST_URL}/engine/config`, { signal: AbortSignal.timeout(10000) })
  if (!res.ok) throw new Error(`Failed to load config (${res.status})`)
  return (await res.json()) as Config
}

export async function patchConfig(patch: Config) {
  const res = await fetch(`${INGEST_URL}/engine/config`, {
    method: "PATCH",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(patch),
    signal: AbortSignal.timeout(10000),
  })
  if (!res.ok) {
    // The settings endpoint returns a public, scalar explanation. Never turn a
    // response body, provider payload or credential diagnostic into UI copy.
    const body: unknown = await res.json().catch(() => null)
    const message = body && typeof body === "object" && "error" in body && typeof body.error === "string" ? body.error.trim() : ""
    const safe = message.length > 0 && message.length <= 400 && !/[\r\n\x00-\x1f{}<>]/.test(message)
      && !/(?:bearer\s+\S+|(?:api[ _-]?key|token|authorization|password|secret)\s*[:=]\s*\S+|https?:\/\/|-----BEGIN|\bsk-[A-Za-z0-9_-]{10,})/i.test(message)
    throw new Error(safe ? message : `Failed to save config (${res.status}). Check model availability, provider sign-in and installation policy, then retry.`)
  }
  return (await res.json()) as Config
}

// Save and read back the pair as one update, so routing never sees mixed providers.
export async function setModelPair(model: string, small_model: string) {
  const provider = model.slice(0, model.indexOf("/"))
  if (!provider || small_model.slice(0, small_model.indexOf("/")) !== provider) throw new Error("Choose both models from the same provider.")
  await patchConfig({ model, small_model })
  const saved = await getConfig()
  if (saved.model !== model || saved.small_model !== small_model) throw new Error("Model choices were not saved. Refresh and retry.")
  return saved
}

// Set the engine-wide default model, as a "providerID/modelID" string.
export async function setDefaultModel(model: string) {
  return patchConfig({ model })
}

// Set the engine-wide small/fast model, as a "providerID/modelID" string. The
// engine uses it for cheap tasks like title generation; doc.haus also reads it to
// route the chat's "Auto" assistant (see routeAgent).
export async function setSmallModel(model: string) {
  return patchConfig({ small_model: model })
}

// Hide providers from routing entirely. A disabled provider drops out of the
// catalog, the model picker, and any agent that would route to it.
export async function setDisabledProviders(ids: string[]) {
  return patchConfig({ disabled_providers: ids })
}

// Register a local OpenAI-compatible provider (LM Studio, Ollama, vLLM...). The
// engine loads it through @ai-sdk/openai-compatible at the given baseURL. The
// baseURL must be reachable from where `opencode serve` runs, not the browser.
export async function addLocalProvider(input: { id: string; name: string; baseURL: string; modelID: string }) {
  return patchConfig({
    provider: {
      [input.id]: {
        npm: "@ai-sdk/openai-compatible",
        name: input.name,
        options: { baseURL: input.baseURL },
        models: { [input.modelID]: { name: input.modelID } },
      },
    },
  })
}

// Vertex (project/location) and Bedrock (region/profile) take their settings from
// provider.options in the engine's config, falling back to env only when unset
// (resolved by the native provider plugins). Setting them here means a
// fresh clone connects these host-credential providers entirely from the UI —
// no GOOGLE_VERTEX_PROJECT / AWS_REGION env needed. Merge into any existing entry
// so a key set elsewhere on the same provider is preserved.
export async function setProviderOptions(id: string, options: Record<string, string>) {
  // An explicitly cleared field resets its override; omitted fields stay intact.
  return patchConfig({ provider: { [id]: { options: Object.fromEntries(Object.entries(options).map(([key, value]) => [key, value || null])) } } })
}

// The currently-saved provider.options, so the cloud-setup fields prefill with
// what the engine is actually using rather than starting blank on every open.
export async function getProviderOptions(id: string) {
  const cfg = await getConfig()
  return (cfg.provider?.[id]?.options ?? {}) as Record<string, string | undefined>
}

export function modelSpec(config?: Pick<Config, "model">) {
  if (!config?.model) return undefined
  const slash = config.model.indexOf("/")
  if (slash < 1 || slash === config.model.length - 1) return undefined
  return { providerID: config.model.slice(0, slash), modelID: config.model.slice(slash + 1) }
}
