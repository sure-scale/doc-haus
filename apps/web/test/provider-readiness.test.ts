import { expect, test } from "bun:test"
import type { IntegrationInfo, ProviderInfo } from "@opencode/client"
import { connectedProviderIDs, GATED_PROVIDERS } from "../src/api/opencode/provider-readiness"

const provider: ProviderInfo = {
  id: "opencode-go", integrationID: "opencode-go", name: "OpenCode Go",
  activation: "enabled", package: "@opencode/ai/providers/openai-compatible",
}
const integration: IntegrationInfo = {
  id: "opencode-go", name: "OpenCode Go", methods: [{ type: "key" }],
  connections: [{ type: "credential", id: "test-account", label: "Account", method: "key" }],
}

test("removing a credential makes an explicitly enabled hosted provider connectable again", () => {
  expect(connectedProviderIDs([provider], [integration])).toEqual(["opencode-go"])
  expect(connectedProviderIDs([provider], [{ ...integration, connections: [] }])).toEqual([])
})

test("shared integrations follow integrationID and environment connections need no stored key", () => {
  const alias = { ...provider, id: "console-models", integrationID: "console-account" }
  const account = { ...integration, id: "console-account" }
  expect(connectedProviderIDs([alias], [account])).toEqual(["console-models"])
  expect(connectedProviderIDs([alias], [{ ...account, connections: [] }])).toEqual([])
  expect(connectedProviderIDs([provider], [{ ...integration, connections: [{ type: "env", name: "MODEL_API_KEY" }] }])).toEqual(["opencode-go"])
})

test("disabled providers remain unavailable even with credentials", () => {
  expect(connectedProviderIDs([{ ...provider, activation: "disabled" }], [integration])).toEqual([])
})

test("native credential-free local discovery and the restricted public catalog remain usable", () => {
  expect(connectedProviderIDs([{ ...provider, id: "ollama", integrationID: undefined }], [])).toEqual(["ollama"])
  const publicProvider = { ...provider, id: "opencode", integrationID: "opencode", settings: { apiKey: "public" } }
  expect(connectedProviderIDs([publicProvider], [{ ...integration, id: "opencode", connections: [] }])).toEqual(["opencode"])
})

test("host credentials still reach the existing live-probe gate", () => {
  const hosts = [...GATED_PROVIDERS].map((id) => ({ ...provider, id, integrationID: id }))
  expect(connectedProviderIDs(hosts, [])).toEqual([...GATED_PROVIDERS])
  expect(connectedProviderIDs(hosts.map((item) => ({ ...item, activation: "disabled" })), [])).toEqual([])
})
