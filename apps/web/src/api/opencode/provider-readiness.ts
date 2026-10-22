import type { IntegrationInfo, ProviderInfo } from "@opencode/client"

// These host sign-ins are verified separately through the existing live probe.
export const GATED_PROVIDERS = new Set(["google-vertex", "google-vertex-anthropic", "amazon-bedrock"])
export const isRemovedProvider = (id: string | undefined) => id === "opencode"
export const isRemovedModel = (reference: string | undefined) => reference?.startsWith("opencode/") ?? false

export function connectedProviderIDs(providers: ProviderInfo[], integrations: IntegrationInfo[]) {
  const byID = new Map(integrations.map((integration) => [integration.id, integration]))
  return providers.filter((provider) => {
    if (isRemovedProvider(provider.id) || provider.activation === "disabled") return false
    if (GATED_PROVIDERS.has(provider.id)) return true
    const integration = byID.get(provider.integrationID ?? provider.id)
    if (integration?.connections.length) return true
    // Native local-provider discovery removes its integration once no sign-in is needed.
    return provider.integrationID === undefined && !integration
  }).map((provider) => provider.id)
}
