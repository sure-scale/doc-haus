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

test("external methods with fields stay with the server-credential setup", () => {
  const vertex = { ...provider, id: "google-vertex", integrationID: "google-vertex", name: "Vertex" }
  const methods: IntegrationInfo["methods"] = [
    { type: "env", names: ["GOOGLE_VERTEX_API_KEY"] },
    { id: "adc", type: "external", label: "Google Cloud credentials", form: [{ key: "project", type: "string", title: "Project", required: true }] },
  ]
  expect(connectionChoices(vertex, { ...integration, id: "google-vertex", methods }, [])).toEqual([])
})

test("the editor explains the sign-in before the check and shows the native problem unchanged after it", () => {
  const before = render()
  expect(before).toContain("Use your Claude Code sign-in")
  expect(before).toContain("Uses the Claude Code sign-in on the computer running doc.haus.")
  expect(before).not.toContain("Check again")

  const missing = render({
    external: { status: "needs_auth", message: "Claude Code is not installed. Install it, then run `claude auth login`.", url: "https://code.claude.com/docs/en/setup" },
    onCheckAgain: () => {},
  })
  expect(missing).toContain("Claude Code is not installed. Install it, then run <code>claude auth login</code>.")
  expect(missing).toContain('href="https://code.claude.com/docs/en/setup"')
  expect(missing).toMatch(/<button type="button" class="primary">Check again<\/button>/)
  expect(render({ external: { status: "needs_auth", message: "Check it", url: "javascript:alert(1)" }, onCheckAgain: () => {} })).not.toContain("javascript:")

  const ready = render({ external: { status: "ready" }, onCheckAgain: () => {} })
  expect(ready).toContain("Claude Code is signed in.")
  expect(ready).toMatch(/<button type="button" class="ghost">Check again<\/button>/)
  expect(render({ external: { status: "ready" }, onCheckAgain: () => {}, checking: true, disabled: true })).toContain("Checking…")
  expect(render({ external: { status: "checking" }, onCheckAgain: () => {} })).toContain("Still checking the sign-in.")
})
