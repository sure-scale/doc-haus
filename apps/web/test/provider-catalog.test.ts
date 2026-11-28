import { afterEach, expect, test } from "bun:test"
import type { IntegrationInfo, ProviderInfo } from "@opencode/client"
import type { Client } from "../src/api/opencode/client"
import { listReusableSignIns, reuseSignIn } from "../src/api/ingest/credential-reuse"
import { saveCandidate, SaveUnconfirmedError, testCandidate } from "../src/api/opencode/provider-candidate"
import { connectedProviderIDs, GATED_PROVIDERS } from "../src/api/opencode/provider-readiness"
import { listAuthMethods, listIntegrations, listProviders, probeProvider, setProviderKey } from "../src/api/opencode/providers"
import { setProviderOptions } from "../src/api/opencode/config"

const originalFetch = globalThis.fetch
afterEach(() => { globalThis.fetch = originalFetch })

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

test("reuse discovery and explicit selection use the cookie-free cross-origin ingest contract", async () => {
  const requests: Request[] = []
  const connection = { id: "selected-connection", integrationID: "opencode-go", label: "Go account", authType: "key" as const }
  const source = { id: "native-source", label: "OpenCode", status: "ready" as const, unavailableCount: 0, connections: [connection] }
  globalThis.fetch = (async (input, init) => {
    const request = new Request(input, init)
    requests.push(request)
    expect(init?.credentials).toBe("omit")
    expect(request.headers.has("cookie")).toBe(false)
    expect(request.headers.has("authorization")).toBe(false)
    return Response.json(request.method === "POST" ? { connection } : { sources: [source] }, {
      headers: { "Access-Control-Allow-Origin": "http://localhost:5174" },
    })
  }) as typeof fetch
  const sources = await listReusableSignIns()
  expect(sources).toEqual([source])
  expect(requests.map(request => request.method)).toEqual(["GET"])
  const selected = sources[0]!.connections[0]!
  expect(await reuseSignIn(sources[0]!.id, selected.id)).toEqual(connection)
  expect(await requests[1]!.json()).toEqual({ sourceId: "native-source", connectionId: "selected-connection" })
  expect(requests.map(request => new URL(request.url).pathname)).toEqual(["/auth/reuse", "/auth/reuse"])
})

test("a lost Save response cannot invite another credential activation", async () => {
  globalThis.fetch = (async () => { throw new TypeError("Network disconnected") }) as unknown as typeof fetch
  await expect(saveCandidate("draft", "primary", "fast")).rejects.toBeInstanceOf(SaveUnconfirmedError)
})

test("a confirmed rollback permits retry; partial state requires inspection", async () => {
  globalThis.fetch = (async () => Response.json({ error: "Previous connection restored.", code: "connection_save_failed", partialState: false }, { status: 503 })) as unknown as typeof fetch
  let failure: unknown
  try { await saveCandidate("draft", "primary", "fast") } catch (error) { failure = error }
  expect(failure).toBeInstanceOf(Error)
  expect(failure).not.toBeInstanceOf(SaveUnconfirmedError)
  globalThis.fetch = (async () => Response.json({ error: "The previous connection could not be restored.", code: "setup_partial_state", partialState: true }, { status: 503 })) as unknown as typeof fetch
  await expect(saveCandidate("draft", "primary", "fast")).rejects.toBeInstanceOf(SaveUnconfirmedError)
})

test("successful Save remains successful when candidate cleanup is pending", async () => {
  globalThis.fetch = (async () => Response.json({ ok: true, sessionClosePending: true })) as unknown as typeof fetch
  expect(await saveCandidate("draft", "primary", "fast")).toEqual({ ok: true, sessionClosePending: true })
  globalThis.fetch = (async () => Response.json({})) as unknown as typeof fetch
  await expect(saveCandidate("draft", "primary", "fast")).rejects.toBeInstanceOf(SaveUnconfirmedError)
})

test("expected model rejection and interrupted transport remain different recovery outcomes", async () => {
  const rejected = { ok: false, kind: "model", error: "Choose an available model." } as const
  globalThis.fetch = (async () => Response.json(rejected, { status: 400 })) as unknown as typeof fetch
  expect(await testCandidate("draft", "fast")).toEqual(rejected)
  globalThis.fetch = (async () => { throw new TypeError("Network disconnected") }) as unknown as typeof fetch
  expect(await testCandidate("draft", "fast")).toMatchObject({ ok: false, kind: "unavailable" })
})

const readinessProvider: ProviderInfo = {
  id: "opencode-go", integrationID: "opencode-go", name: "OpenCode Go",
  activation: "enabled", package: "@opencode/ai/providers/openai-compatible",
}
const readinessIntegration: IntegrationInfo = {
  id: "opencode-go", name: "OpenCode Go", methods: [{ type: "key" }],
  connections: [{ type: "credential", id: "test-account", label: "Account", method: "key" }],
}

test("removing a credential makes an explicitly enabled hosted provider connectable again", () => {
  expect(connectedProviderIDs([readinessProvider], [readinessIntegration])).toEqual(["opencode-go"])
  expect(connectedProviderIDs([readinessProvider], [{ ...readinessIntegration, connections: [] }])).toEqual([])
})

test("shared integrations follow integrationID and environment connections need no stored key", () => {
  const alias = { ...readinessProvider, id: "console-models", integrationID: "console-account" }
  const account = { ...readinessIntegration, id: "console-account" }
  expect(connectedProviderIDs([alias], [account])).toEqual(["console-models"])
  expect(connectedProviderIDs([alias], [{ ...account, connections: [] }])).toEqual([])
  expect(connectedProviderIDs([readinessProvider], [{ ...readinessIntegration, connections: [{ type: "env", name: "MODEL_API_KEY" }] }])).toEqual(["opencode-go"])
})

test("disabled providers remain unavailable even with credentials", () => {
  expect(connectedProviderIDs([{ ...readinessProvider, activation: "disabled" }], [readinessIntegration])).toEqual([])
})

test("native credential-free local discovery remains usable while the removed public provider stays unavailable", () => {
  expect(connectedProviderIDs([{ ...readinessProvider, id: "ollama", integrationID: undefined }], [])).toEqual(["ollama"])
  const publicProvider = { ...readinessProvider, id: "opencode", integrationID: "opencode", settings: { apiKey: "public" } }
  expect(connectedProviderIDs([publicProvider], [{ ...readinessIntegration, id: "opencode", connections: [] }])).toEqual([])
  expect(connectedProviderIDs([publicProvider, readinessProvider], [{ ...readinessIntegration, id: "opencode" }, readinessIntegration])).toEqual(["opencode-go"])
})

test("host credentials still reach the existing live-probe gate", () => {
  const hosts = [...GATED_PROVIDERS].map((id) => ({ ...readinessProvider, id, integrationID: id }))
  expect(connectedProviderIDs(hosts, [])).toEqual([...GATED_PROVIDERS])
  expect(connectedProviderIDs(hosts.map((item) => ({ ...item, activation: "disabled" })), [])).toEqual([])
})

test("an external sign-in counts as connected only while its active connection reports no needs_auth", () => {
  const claude = { ...readinessProvider, id: "claude-agent", integrationID: "claude-agent", name: "Claude Agent", activation: "auto" as const }
  const signedIn: IntegrationInfo = {
    id: "claude-agent", name: "Claude Agent", methods: [{ id: "claude-login", type: "external", label: "Claude Code login" }],
    connections: [{ type: "credential", id: "local-sign-in", label: "Claude Agent", method: "external" }],
  }
  expect(connectedProviderIDs([claude], [signedIn])).toEqual(["claude-agent"])
  const needsAuth = { ...signedIn, connections: [{ ...signedIn.connections[0]!, status: { status: "needs_auth" as const, message: "Claude Code is not logged in." } }] }
  expect(connectedProviderIDs([claude], [needsAuth])).toEqual([])
  expect(connectedProviderIDs([claude], [{ ...signedIn, connections: [] }])).toEqual([])
})
