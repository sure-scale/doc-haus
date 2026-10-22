import { expect, test } from "bun:test"
import type { Client } from "../src/api/opencode/client"
import { listAuthMethods, listIntegrations, listProviders, probeProvider, setProviderKey } from "../src/api/opencode/providers"
import { setProviderOptions } from "../src/api/opencode/config"

test("inactive setup metadata neither connects a provider nor creates model choices; sign-in uses its native integration", async () => {
  const previous = globalThis.fetch
  const connected: { integrationID: string; key: string }[] = []
  let catalogRequests = 0
  const client = {
    provider: { list: async () => ({ data: [] }) },
    model: { list: async () => ({ data: [] }) },
    integration: {
      list: async () => ({ data: [{ id: "native-login", name: "Native", methods: [{ type: "key" }], connections: [] }] }),
      connect: { key: async (input: { integrationID: string; key: string }) => { connected.push(input) } },
    },
  } as unknown as Client
  globalThis.fetch = (async (_input, options) => {
    catalogRequests++
    expect(options?.signal).toBeInstanceOf(AbortSignal)
    return Response.json({ data: [
      { id: "amazon-bedrock", name: "Bedrock", activation: "auto", integrationID: "amazon-bedrock", host: { type: "aws-chain", options: { profile: "work" }, sources: ["AWS_PROFILE"] } },
      { id: "alias-models", name: "Alias", activation: "auto", integrationID: "native-login" },
    ] })
  }) as typeof fetch
  try {
    const result = await listProviders(client)
    expect(result.all.map(provider => provider.id)).toEqual(["amazon-bedrock", "alias-models"])
    expect(result.connected).toEqual([])
    expect(result.all.every(provider => !provider.nativeAvailable)).toBe(true)
    expect(result.all.every(provider => Object.keys(provider.models).length === 0)).toBe(true)
    await setProviderKey(client, "alias-models", "synthetic-fixture-key")
    expect(connected.map(input => input.integrationID)).toEqual(["native-login"])
    expect(catalogRequests).toBe(2)
  } finally { globalThis.fetch = previous }
})

test("native environment discovery supplies usable models without importing another installation", async () => {
  const previous = globalThis.fetch
  const hosted = { id: "hosted-models", integrationID: "hosted-account", name: "Hosted", activation: "auto", package: "native-provider" }
  const bedrock = { id: "amazon-bedrock", name: "Bedrock", activation: "enabled", package: "native-provider" }
  const publicProvider = { id: "opencode", name: "OpenCode", activation: "enabled", settings: { apiKey: "public" }, package: "native-provider" }
  const vertex = { id: "google-vertex", name: "Vertex", activation: "auto", host: { type: "google-adc", options: {}, sources: ["Google ADC file"] } }
  const requests: string[] = []
  const client = {
    provider: { list: async () => ({ data: [hosted, bedrock, publicProvider] }) },
    model: { list: async () => ({ data: [{ providerID: hosted.id, id: "only-model", name: "Only model", enabled: true }] }) },
    integration: { list: async () => ({ data: [{
      id: hosted.integrationID, name: "Hosted account", methods: [{ type: "env", names: ["HOSTED_API_KEY"] }],
      connections: [{ type: "env", name: "HOSTED_API_KEY" }],
    }] }) },
  } as unknown as Client
  globalThis.fetch = (async input => {
    requests.push(String(input))
    return Response.json({ data: [hosted, bedrock, publicProvider, vertex] })
  }) as typeof fetch
  try {
    const result = await listProviders(client)
    expect(result.connected).toEqual([hosted.id, bedrock.id])
    expect(result.all.find(provider => provider.id === hosted.id)).toMatchObject({
      nativeAvailable: true, credentialSources: ["HOSTED_API_KEY"],
      models: { "only-model": { id: "only-model", name: "Only model" } },
    })
    expect(result.all.find(provider => provider.id === bedrock.id)?.nativeAvailable).toBe(true)
    expect(result.all.some(provider => provider.id === publicProvider.id)).toBe(false)
    expect(result.all.find(provider => provider.id === vertex.id)).toMatchObject({ nativeAvailable: false, credentialSources: [], models: {} })
    expect(requests).toHaveLength(1)
    expect(requests[0]).toEndWith("/api/dochaus/provider-catalog")
  } finally { globalThis.fetch = previous }
})

test("raw native removed provider, models and sign-in never enter web choices; Go and other aliases remain", async () => {
  const previous = globalThis.fetch
  const ids = ["opencode", "opencode-go", "other-provider"]
  const providers = ids.map(id => ({ id, integrationID: id === "other-provider" ? "other-account" : id, name: id, activation: "enabled", package: "native-provider" }))
  const integrations = ["opencode", "opencode-go", "other-account"].map(id => ({ id, name: id, methods: [{ type: "key" }], connections: [{ type: "credential", id: `${id}-key`, label: "Saved", method: "key" }] }))
  let requests = 0
  let writes = 0
  const client = {
    provider: { list: async () => ({ data: providers }) },
    model: { list: async () => ({ data: ids.map(providerID => ({ providerID, id: "model", name: "Model", enabled: true })) }) },
    integration: { list: async () => ({ data: integrations }), connect: { key: async () => { writes++ } } },
  } as unknown as Client
  globalThis.fetch = (async () => { requests++; return Response.json({ data: providers }) }) as unknown as typeof fetch
  try {
    const catalog = await listProviders(client)
    expect(catalog.all.map(provider => provider.id)).toEqual(["opencode-go", "other-provider"])
    expect(catalog.connected).toEqual(["opencode-go", "other-provider"])
    expect(catalog.all.map(provider => Object.keys(provider.models))).toEqual([["model"], ["model"]])
    expect((await listIntegrations(client)).map(integration => integration.id)).toEqual(["opencode-go", "other-account"])
    expect(Object.keys(await listAuthMethods(client))).toEqual(["opencode-go", "other-account"])
    await expect(setProviderKey(client, "opencode", "unused")).rejects.toThrow("removed")
    expect(await probeProvider(client, "opencode", "model")).toMatchObject({ ok: false, error: expect.stringContaining("another provider") })
    expect(requests).toBe(1)
    expect(writes).toBe(0)
  } finally { globalThis.fetch = previous }
})

test("clearing a cloud profile explicitly resets that override without replacing other provider settings", async () => {
  const previous = globalThis.fetch
  const patches: unknown[] = []
  globalThis.fetch = (async (_input, options) => {
    patches.push(JSON.parse(String(options?.body)))
    return Response.json({})
  }) as typeof fetch
  try {
    await setProviderOptions("amazon-bedrock", { region: "us-east-1", profile: "" })
    expect(patches).toEqual([{ provider: { "amazon-bedrock": { options: { region: "us-east-1", profile: null } } } }])
  } finally { globalThis.fetch = previous }
})
