import type { ModelInfo } from "@opencode/client"
import { location, type Client } from "./client"
import { connectedProviderIDs } from "./provider-readiness"
import { OPENCODE_URL } from "../../config"

export type ProviderMetadata = {
  id: string; name: string; integrationID?: string; activation: "auto" | "enabled" | "disabled"
  host?: { type: "google-adc" | "aws-chain"; options: Record<string, string>; sources: string[] }
}

async function listProviderCatalog(client: Client): Promise<ProviderMetadata[]> {
  const directory = location(client)?.directory
  const response = await fetch(`${OPENCODE_URL}/api/dochaus/provider-catalog`, {
    credentials: "include", signal: AbortSignal.timeout(10_000),
    ...(directory ? { headers: { "x-opencode-directory": directory } } : {}),
  })
  if (!response.ok) throw new Error("Could not discover provider setup. Check the engine and retry.")
  const result = await response.json() as { data: ProviderMetadata[] }
  if (!Array.isArray(result.data)) throw new Error("The engine returned an invalid provider catalog.")
  return result.data
}

// Setup metadata includes unavailable providers. Only native availability and
// enabled model responses can supply model choices or connected status.
export async function listProviders(client: Client) {
  const [providers, models, integrations, catalog] = await Promise.all([
    client.provider.list({ location: location(client) }, { signal: AbortSignal.timeout(10000) }),
    client.model.list({ location: location(client) }, { signal: AbortSignal.timeout(10000) }),
    client.integration.list({ location: location(client) }, { signal: AbortSignal.timeout(10000) }),
    listProviderCatalog(client),
  ])
  const available = new Map(providers.data.map(provider => [provider.id, provider]))
  const accounts = new Map(integrations.data.map(integration => [integration.id, integration]))
  const all = catalog.map(provider => ({
    ...provider,
    nativeAvailable: available.has(provider.id),
    publicModels: provider.id === "opencode" && available.get(provider.id)?.settings?.apiKey === "public",
    // Native integration discovery includes ambient environment connections.
    // Only their names are presentation data; values remain in the engine.
    credentialSources: [...new Set((accounts.get(provider.integrationID ?? provider.id)?.connections ?? [])
      .map(connection => connection.type === "env" ? connection.name : "Saved sign-in for this installation"))],
    models: Object.fromEntries(models.data.filter((model: ModelInfo) => model.providerID === provider.id && model.enabled)
      .map((model: ModelInfo) => [model.id, { id: model.id, name: model.name }])),
  }))
  // Explicit provider config enables its catalog even when it has no credential.
  return { all, connected: connectedProviderIDs(providers.data, integrations.data) }
}

// Per-provider auth methods (api key vs oauth). Empty array => credentials come
// from the environment (Vertex ADC, Bedrock/Azure SDK creds), not a key field.
export async function listIntegrations(client: Client) {
  return (await client.integration.list({ location: location(client) }, { signal: AbortSignal.timeout(10000) })).data
}
export async function listAuthMethods(client: Client) {
  const integrations = await listIntegrations(client)
  return Object.fromEntries(integrations.map(integration => [integration.id, integration.methods.flatMap<{ type: "api" | "oauth"; label: string }>(method =>
    method.type === "key" ? [{ type: "api" as const, label: method.label || "API key" }] : method.type === "oauth" ? [{ type: "oauth" as const, label: method.label }] : [])]))
}

// Resolve aliases through their native integration rather than assuming IDs match.
export async function setProviderKey(client: Client, id: string, key: string) {
  const [providers, integrations] = await Promise.all([listProviderCatalog(client), listIntegrations(client)])
  const integrationID = providers.find(provider => provider.id === id)?.integrationID ?? id
  if (!integrations.some(integration => integration.id === integrationID && integration.methods.some(method => method.type === "key"))) throw new Error("This provider does not offer API key sign-in. Use its available sign-in method.")
  return client.integration.connect.key({ integrationID, key, location: location(client) }, { signal: AbortSignal.timeout(15000) })
}

// Real credential probe for host-credential providers (Vertex ADC, Bedrock AWS).
// The engine reports these "connected" on project/region presence alone — it never
// checks that the sign-in actually resolves — so a project set without a working
// gcloud/AWS login still lists as ready and only fails at the first real call. To
// gate the picker on real auth, the managed endpoint runs a fixed, stateless,
// bounded probe against the selected enabled catalog model. It needs legal and
// consent activation but can run before the user completes model setup.
export async function probeProvider(
  client: Client,
  providerID: string,
  modelID: string,
): Promise<{ ok: true } | { ok: false; error: string }> {
  try {
    const directory = location(client)?.directory
    const response = await fetch(`${OPENCODE_URL}/api/dochaus/provider-test`, {
      method: "POST", credentials: "include", signal: AbortSignal.timeout(25_000),
      headers: { "Content-Type": "application/json", ...(directory ? { "x-opencode-directory": directory } : {}) },
      body: JSON.stringify({ providerID, modelID }),
    })
    const result = await response.json()
    return response.ok && result.ok === true ? { ok: true } : { ok: false, error: typeof result.error === "string" ? result.error : "The provider could not be verified." }
  } catch {
    return { ok: false, error: "The provider test could not complete within 25 seconds. Check the engine and connection, then retry." }
  }
}
