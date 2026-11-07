import type { IntegrationInfo, ProviderInfo } from "@opencode/client"

// These providers use server credential chains. Availability supplies a catalog;
// shared AI setup tests access to the exact selected models before Save.
export const GATED_PROVIDERS = new Set(["google-vertex", "google-vertex-anthropic", "amazon-bedrock"])
export const isRemovedProvider = (id: string | undefined) => id === "opencode"
export const isRemovedModel = (reference: string | undefined) => reference?.startsWith("opencode/") ?? false

export function connectedProviderIDs(providers: ProviderInfo[], integrations: IntegrationInfo[]) {
  const byID = new Map(integrations.map((integration) => [integration.id, integration]))
  return providers.filter((provider) => {
    if (isRemovedProvider(provider.id) || provider.activation === "disabled") return false
    if (GATED_PROVIDERS.has(provider.id)) return true
    const integration = byID.get(provider.integrationID ?? provider.id)
    // The first connection is the active one. An integration that reports
    // needs_auth, such as an external sign-in that is missing, is not ready.
    const active = integration?.connections[0]
    if (active) return active.status?.status !== "needs_auth"
    // Native local-provider discovery removes its integration once no sign-in is needed.
    return provider.integrationID === undefined && !integration
  }).map((provider) => provider.id)
}
