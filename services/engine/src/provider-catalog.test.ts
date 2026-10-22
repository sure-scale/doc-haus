import { expect, test } from "bun:test"
import { providerCatalog } from "./provider-catalog"

test("provider setup catalog excludes only removed opencode and preserves Go and nonsecret cloud setup", () => {
  const catalog = providerCatalog([
    { id: "opencode", name: "Removed", activation: "enabled", integrationID: "opencode", settings: { apiKey: "synthetic-secret" } },
    { id: "opencode-go", name: "OpenCode Go", activation: "auto", integrationID: "opencode-go" },
    { id: "company", name: "Company", activation: "enabled", integrationID: "opencode" },
    { id: "google-vertex", name: "Vertex", activation: "auto", settings: { project: "synthetic-project", location: "europe-west1", apiKey: "synthetic-secret" } },
  ], {}, "/synthetic-home", () => false)
  expect(catalog.map(provider => provider.id)).toEqual(["opencode-go", "company", "google-vertex"])
  expect(catalog[0].integrationID).toBe("opencode-go")
  expect(catalog[2].host?.options).toEqual({ project: "synthetic-project", location: "europe-west1" })
  expect(JSON.stringify(catalog)).not.toContain("synthetic-secret")
})
