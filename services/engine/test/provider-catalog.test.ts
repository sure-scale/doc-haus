import { expect, test } from "bun:test"
import { providerCatalog } from "../src/provider-catalog"

test("setup catalog retains inactive cloud providers and emits only nonsecret metadata", () => {
  const providers = [{ id: "google-vertex", name: "Vertex", activation: "auto" as const, integrationID: "google-vertex",
    settings: { project: "synthetic-project", location: "global", apiKey: "synthetic-private-key", headers: { Authorization: "synthetic-private-header" } } },
  { id: "amazon-bedrock", name: "Bedrock", activation: "auto" as const, integrationID: "amazon-bedrock", settings: { region: "us-east-1" } }]
  const result = providerCatalog(providers, { AWS_PROFILE: "work", AWS_SECRET_ACCESS_KEY: "synthetic-private-secret" }, "/fixture-home", file => file.endsWith("application_default_credentials.json") || /\.aws[\\/]config$/.test(file))
  expect(result[0]).toEqual({ id: "google-vertex", name: "Vertex", activation: "auto", integrationID: "google-vertex", host: { type: "google-adc", options: { project: "synthetic-project", location: "global" }, sources: ["Google ADC file"] } })
  expect(result[1]?.host).toEqual({ type: "aws-chain", options: { region: "us-east-1", profile: "work" }, sources: ["AWS_PROFILE", "AWS shared profiles"] })
  expect(JSON.stringify(result)).not.toContain("synthetic-private")
  expect(providers[0]?.activation).toBe("auto")
})

test("catalog removes exact opencode while preserving Go and Console aliases without assuming access", () => {
  const result = providerCatalog([
    { id: "opencode", name: "OpenCode Zen", activation: "enabled", integrationID: "opencode", settings: { apiKey: "synthetic-private-key" } },
    { id: "opencode-go", name: "OpenCode Go", activation: "auto", integrationID: "opencode-go" },
    { id: "organization-models", name: "Organization", activation: "auto", integrationID: "console-account" },
    { id: "company", name: "Company", activation: "enabled", integrationID: "opencode", settings: { apiKey: "synthetic-private-key" } },
  ], {}, "/fixture-home", () => false)
  expect(result.map(provider => provider.id)).toEqual(["opencode-go", "organization-models", "company"])
  expect(result.map(provider => provider.integrationID)).toEqual(["opencode-go", "console-account", "opencode"])
  expect(result.every(provider => !("models" in provider) && !("settings" in provider) && !("connected" in provider))).toBe(true)
  expect(JSON.stringify(result)).not.toContain("synthetic-private")
})

test("catalog lists an external sign-in provider without host detection or its settings", () => {
  const result = providerCatalog([
    { id: "claude-agent", name: "Claude Agent", activation: "auto", integrationID: "claude-agent",
      settings: { apiKey: "synthetic-private-key", baseURL: "https://synthetic.example/v1" } },
  ], { OPENCODE_CLAUDE_PATH: "/synthetic/claude" }, "/fixture-home", () => true)
  expect(result).toEqual([{ id: "claude-agent", name: "Claude Agent", activation: "auto", integrationID: "claude-agent" }])
  expect(JSON.stringify(result)).not.toContain("synthetic")
})

test("catalog lists a provider whose sign-in has fields without its configured values", () => {
  const result = providerCatalog([
    { id: "azure", name: "Azure", activation: "auto", integrationID: "azure",
      settings: { resourceName: "synthetic-resource", apiKey: "synthetic-private-key" } },
  ], { AZURE_API_KEY: "synthetic-private-key" }, "/fixture-home", () => true)
  expect(result).toEqual([{ id: "azure", name: "Azure", activation: "auto", integrationID: "azure" }])
  expect(JSON.stringify(result)).not.toContain("synthetic")
})
