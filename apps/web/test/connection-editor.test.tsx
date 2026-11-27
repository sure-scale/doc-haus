import { expect, test } from "bun:test"
import { renderToStaticMarkup } from "react-dom/server"
import type { IntegrationInfo } from "@opencode/client"
import ConnectionEditor, { connectionChoices, type ConnectionDraft } from "../src/components/ai-setup/ConnectionEditor"
import type { Provider } from "../src/components/settings/types"

const provider: Provider = {
  id: "claude-agent", name: "Claude Agent", integrationID: "claude-agent", activation: "auto",
  nativeAvailable: false, credentialSources: [], models: {},
}
const integration: IntegrationInfo = {
  id: "claude-agent", name: "Claude Agent", connections: [],
  methods: [{ id: "claude-login", type: "external", label: "Claude Code login (claude auth login)" }],
}
const draft: ConnectionDraft = { choice: "method:0", key: "", answer: {}, options: {}, name: "" }
const render = (props: Partial<Parameters<typeof ConnectionEditor>[0]> = {}) => renderToStaticMarkup(
  <ConnectionEditor provider={provider} integration={integration} reusable={[]} draft={draft} onChange={() => {}} authCode="" onAuthCode={() => {}} onAuthAction={() => {}} {...props} />,
)

test("the Claude Code sign-in is offered once, as a new or a saved connection", () => {
  expect(connectionChoices(provider, integration, [])).toEqual([
    { value: "method:0", label: "Use your Claude Code sign-in", method: integration.methods[0], external: true },
  ])
  const saved = { ...integration, connections: [{ type: "credential" as const, id: "local-sign-in", label: "Claude Agent", method: "external" as const }] }
  expect(connectionChoices(provider, saved, [])).toEqual([
    { value: "credential:local-sign-in", label: "Use your Claude Code sign-in", external: true },
  ])
})

test("a cloud setup's external sign-in stays with its server-credential setup", () => {
  const vertex = { ...provider, id: "google-vertex", integrationID: "google-vertex", name: "Vertex" }
  const methods: IntegrationInfo["methods"] = [
    { type: "env", names: ["GOOGLE_VERTEX_API_KEY"] },
    { id: "adc", type: "external", label: "Google Cloud credentials", form: [{ key: "project", type: "string", title: "Project", required: true }] },
  ]
  expect(connectionChoices(vertex, { ...integration, id: "google-vertex", methods }, [])).toEqual([])
})

test("external sign-in links allow https URLs and reject javascript: URLs", () => {
  const safe = render({
    external: { status: "needs_auth", message: "Install CLI", url: "https://code.claude.com/docs/en/setup" },
    onCheckAgain: () => {},
  })
  expect(safe).toContain('href="https://code.claude.com/docs/en/setup"')
  const unsafe = render({ external: { status: "needs_auth", message: "Check it", url: "javascript:alert(1)" }, onCheckAgain: () => {} })
  expect(unsafe).not.toContain("javascript:")
})

test("an external sign-in with fields is offered with its fields and stays offered beside a saved one", () => {
  const azure: Provider = { ...provider, id: "azure", integrationID: "azure", name: "Azure" }
  const method: IntegrationInfo["methods"][number] = {
    id: "azure-cli", type: "external", label: "Microsoft Entra ID (Azure CLI)",
    form: [{ key: "resourceName", type: "string", title: "Azure resource", required: true, pattern: "^[a-zA-Z0-9][a-zA-Z0-9-]*$", description: "Enter the name of your Azure OpenAI or AI Services resource." }],
  }
  const info: IntegrationInfo = { id: "azure", name: "Azure", connections: [], methods: [{ type: "env", names: ["AZURE_API_KEY"] }, { type: "key", label: "API key" }, method] }
  expect(connectionChoices(azure, info, [])).toEqual([
    { value: "method:1", label: "API key", method: info.methods[1] },
    { value: "method:2", label: "Microsoft Entra ID (Azure CLI)", method, external: true },
  ])
  const saved = { ...info, connections: [{ type: "credential" as const, id: "entra", label: "Azure", method: "external" as const }] }
  expect(connectionChoices(azure, saved, []).map(choice => choice.value)).toEqual(["credential:entra", "method:1", "method:2"])
})
