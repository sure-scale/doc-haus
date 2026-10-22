import { getConfig } from "./config"
import { OPENCODE_URL } from "../../config"
import { isRemovedProvider } from "./provider-readiness"

export type ModelConnection = {
  providerID: string
  modelID: string
  status: "checking" | "ready" | "auth" | "model" | "unavailable"
  expired?: boolean
  restriction?: "free-tier" | "region"
}
type ProbeResult = { ok: true } | { ok: false; kind: "auth" | "model" | "unavailable"; expired?: boolean; restriction?: "free-tier" | "region" }

export async function testModelConnection(providerID: string, modelID: string, force: boolean): Promise<ProbeResult> {
  if (isRemovedProvider(providerID)) return { ok: false, kind: "model" }
  try {
    const response = await fetch(`${OPENCODE_URL}/api/dochaus/provider-test`, {
      method: "POST", credentials: "include", signal: AbortSignal.timeout(25_000),
      headers: { "Content-Type": "application/json" }, body: JSON.stringify({ providerID, modelID, ...(force ? { force: true } : {}) }),
    })
    const result = await response.json()
    if (response.ok && result.ok === true) return { ok: true }
    // HTTP authentication can belong to the local engine, not the model provider.
    // Only the managed endpoint's typed result can invalidate a saved sign-in.
    if (result.kind === "auth" || result.kind === "model") return { ok: false, kind: result.kind, expired: result.expired === true,
      ...(result.kind === "model" && (result.restriction === "free-tier" || result.restriction === "region") ? { restriction: result.restriction } : {}) }
  } catch { /* A timeout or offline endpoint does not prove expired credentials. */ }
  return { ok: false, kind: "unavailable" }
}

// Memory only: an old success must never survive restarting the application.
// The engine coalesces checks across tabs and invalidates its short cache when
// credentials or settings change. UI checks never send matter text or resume work.
export function createProviderHealthMonitor(
  config: () => Promise<{ model?: string; small_model?: string }> = getConfig,
  probe: typeof testModelConnection = testModelConnection,
) {
  let snapshot: readonly ModelConnection[] = []
  let version = 0
  let pending: Promise<void> | undefined
  const listeners = new Set<() => void>()
  const rejectedModels = new Set<string>()
  function publish(next: readonly ModelConnection[]) {
    snapshot = next
    for (const listener of listeners) listener()
  }
  function check(force = false): Promise<void> {
    if (pending && !force) return pending
    const current = ++version
    const task = (async () => {
      try {
        const saved = await config()
        if (version !== current) return
        const selected = [saved.model, saved.small_model].filter((ref): ref is string => Boolean(ref)).map(ref => ref.split("#")[0]!)
        const models = [...new Set([...rejectedModels, ...selected])].map(ref => {
          const slash = ref.indexOf("/")
          return slash > 0 && slash < ref.length - 1 ? { providerID: ref.slice(0, slash), modelID: ref.slice(slash + 1) } : undefined
        }).filter((model): model is { providerID: string; modelID: string } => Boolean(model))
        publish(models.map(model => ({ ...model, status: isRemovedProvider(model.providerID) ? "model" : "checking" })))
        await Promise.all(models.filter(model => !isRemovedProvider(model.providerID)).map(async model => {
          const result = await probe(model.providerID, model.modelID, force)
          if (version !== current) return
          if (result.ok) rejectedModels.delete(`${model.providerID}/${model.modelID}`)
          publish(snapshot.map(item => item.providerID === model.providerID && item.modelID === model.modelID
            ? { ...model, status: result.ok ? "ready" : result.kind, ...(!result.ok && result.expired ? { expired: true } : {}),
              ...(!result.ok && result.kind === "model" && result.restriction ? { restriction: result.restriction } : {}) } : item))
        }))
      } catch {
        if (version === current) publish([{ providerID: "", modelID: "", status: "unavailable" }])
      }
    })()
    pending = task
    void task.finally(() => { if (pending === task) pending = undefined })
    return task
  }
  return {
    check,
    // Setup has already tested these exact models and saved the combined
    // connection/model choices. Reuse that result without another paid probe.
    confirmed(models: readonly { providerID: string; modelID: string }[]) {
      ++version
      pending = undefined
      const unique = new Map(models.map(model => {
        const selected = { ...model, modelID: model.modelID.split("#")[0]! }
        return [`${selected.providerID}/${selected.modelID}`, selected]
      }))
      publish([...unique].map(([ref, model]) => {
        rejectedModels.delete(ref)
        return { ...model, status: "ready" as const }
      }))
    },
    getSnapshot: () => snapshot,
    subscribe: (listener: () => void) => { listeners.add(listener); return () => { listeners.delete(listener) } },
    // Called only for a typed native failure and the failing session's model.
    // A late failure from an old request cannot undo a new connection check.
    version: () => version,
    rejected(model: { providerID: string; id: string }, error: { type: string; message: string; status?: number }, observedVersion = version) {
      if (version !== observedVersion || !["provider.auth", "provider.no-route"].includes(error.type)) return
      if (isRemovedProvider(model.providerID)) {
        ++version
        rejectedModels.delete(`${model.providerID}/${model.id}`)
        publish([...snapshot.filter(item => item.providerID !== model.providerID || item.modelID !== model.id),
          { providerID: model.providerID, modelID: model.id, status: "model" }])
        return
      }
      // Session summaries omit the native error body. Keep a verified typed 403
      // access result instead of replacing it with their lossy auth label. No
      // extra model request is needed; explicit rechecks can refresh the result.
      const known = snapshot.find(item => item.providerID === model.providerID && item.modelID === model.id)
      if (error.type === "provider.auth" && error.status === 403 && known?.status === "model" && known.restriction) return
      ++version
      rejectedModels.add(`${model.providerID}/${model.id}`)
      const failure: ModelConnection = { providerID: model.providerID, modelID: model.id,
        status: error.type === "provider.auth" ? "auth" : "model",
        ...(error.type === "provider.auth" && /\b(?:api[ -]?key|key)\b.{0,40}\bexpired\b|\bexpired\b.{0,40}\b(?:api[ -]?key|key)\b/i.test(error.message) ? { expired: true } : {}) }
      publish([...snapshot.filter(item => item.providerID !== model.providerID || item.modelID !== model.id), failure])
    },
  }
}
export const providerHealth = createProviderHealthMonitor()

export function providerConnection(connections: readonly ModelConnection[], providerID: string) {
  const matches = connections.filter(item => item.providerID === providerID)
  return matches.find(item => item.status === "auth") ?? matches.find(item => item.status === "model")
    ?? matches.find(item => item.status === "unavailable") ?? matches.find(item => item.status === "checking") ?? matches[0]
}
export function connectionMessage(connection: ModelConnection) {
  if (isRemovedProvider(connection.providerID)) return "This provider has been removed. Choose another provider and save your model choices in Settings."
  if (connection.status === "auth") return connection.expired
    ? "The saved API key has expired. Create a new key in your provider’s console and replace it in Settings. Signing in again will not renew this key."
    : "The provider rejected the saved sign-in. Check or reconnect your account in Settings."
  if (connection.status === "model") {
    if (connection.restriction === "free-tier") return "The provider is restricting access to this free model. Check the provider’s free-model availability or choose another model in Settings."
    if (connection.restriction === "region") return "The provider does not offer this model in your region. Choose another available model or provider in Settings."
    return "The selected model is unavailable. Check model access and your model choices in Settings."
  }
  if (connection.status === "unavailable") return "The model connection could not be verified. Check your network and provider availability, then try again."
  return connection.status === "checking" ? "Checking model connection…" : "Connection verified"
}
export function providerName(id: string) {
  return ({ "opencode-go": "OpenCode Go", opencode: "Removed provider", openai: "OpenAI", google: "Google AI", "google-vertex": "Google Vertex", "amazon-bedrock": "Amazon Bedrock" } as Record<string, string>)[id] ?? id
}
