import type { IntegrationInfo, ProviderInfo } from "@opencode/client"

// These host sign-ins are verified separately through the existing live probe.
export const GATED_PROVIDERS = new Set(["google-vertex", "google-vertex-anthropic", "amazon-bedrock"])

export function connectedProviderIDs(providers: ProviderInfo[], integrations: IntegrationInfo[]) {
  const byID = new Map(integrations.map((integration) => [integration.id, integration]))
  return providers.filter((provider) => {
    if (provider.activation === "disabled") return false
    if (GATED_PROVIDERS.has(provider.id)) return true
    const integration = byID.get(provider.integrationID ?? provider.id)
    if (integration?.connections.length) return true
    // Native local-provider discovery removes its integration once no sign-in is
    // needed. OpenCode also exposes a restricted free-model catalog publicly.
    if (provider.integrationID === undefined && !integration) return true
    return provider.id === "opencode" && provider.settings?.apiKey === "public"
  }).map((provider) => provider.id)
}
