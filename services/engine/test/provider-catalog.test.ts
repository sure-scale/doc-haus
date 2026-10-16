import { expect, test } from "bun:test"
import { providerCatalog } from "../src/provider-catalog"

test("setup catalog retains inactive cloud providers and emits only nonsecret metadata", () => {
  const providers = [{ id: "google-vertex", name: "Vertex", activation: "auto" as const, integrationID: "google-vertex",
    settings: { location: "global", apiKey: "synthetic-private-key", headers: { Authorization: "synthetic-private-header" } } },
  { id: "amazon-bedrock", name: "Bedrock", activation: "auto" as const, integrationID: "amazon-bedrock", settings: { region: "us-east-1" } }]
  const result = providerCatalog(providers, { AWS_PROFILE: "work", AWS_SECRET_ACCESS_KEY: "synthetic-private-secret" }, "/fixture-home", file => file.endsWith("application_default_credentials.json") || file.endsWith(".aws/config"))
  expect(result[0]).toEqual({ id: "google-vertex", name: "Vertex", activation: "auto", integrationID: "google-vertex", host: { type: "google-adc", options: { location: "global" }, sources: ["Google ADC file"] } })
  expect(result[1]?.host).toEqual({ type: "aws-chain", options: { region: "us-east-1", profile: "work" }, sources: ["AWS_PROFILE", "AWS shared profiles"] })
  expect(JSON.stringify(result)).not.toContain("synthetic-private")
  expect(providers[0]?.activation).toBe("auto")
})

test("catalog preserves distinct native integration IDs without assuming access", () => {
  const result = providerCatalog([
    { id: "opencode", name: "OpenCode Zen", activation: "enabled", integrationID: "opencode" },
    { id: "opencode-go", name: "OpenCode Go", activation: "auto", integrationID: "opencode-go" },
    { id: "organization-models", name: "Organization", activation: "auto", integrationID: "console-account" },
  ], {}, "/fixture-home", () => false)
  expect(result.map(provider => provider.integrationID)).toEqual(["opencode", "opencode-go", "console-account"])
  expect(result.every(provider => !("models" in provider) && !("settings" in provider) && !("connected" in provider))).toBe(true)
})
