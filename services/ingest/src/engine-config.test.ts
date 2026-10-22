import { afterEach, expect, test } from "bun:test"
import { mkdtempSync, readFileSync, rmSync, statSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { engineConfigFile, patchEngineConfig, readEngineConfig, readEngineSettings, setAgentDisabled, setSkillDisabled, writeEngineOverrides } from "./engine-config"
const previous = { ...process.env }
const directories: string[] = []
afterEach(() => {
  for (const key of ["DOCHAUS_DIR", "DOCHAUS_RUNTIME_DIR"]) if (previous[key] === undefined) delete process.env[key]; else process.env[key] = previous[key]
  directories.splice(0).forEach(directory => rmSync(directory, { recursive: true, force: true }))
})
function fixture() {
  const directory = mkdtempSync(path.join(tmpdir(), "dochaus-config-")); directories.push(directory)
  process.env.DOCHAUS_DIR = directory
  process.env.DOCHAUS_RUNTIME_DIR = path.join(directory, "runtime")
  const file = path.join(directory, "opencode.json")
  writeFileSync(file, JSON.stringify({ plugins: ["./plugin/legal-v2.ts"], agent: { draft: { prompt: "matter policy" } }, permission: { skill: "allow", external_directory: "deny" } }))
  return file
}
test("incremental edits preserve immutable policy and sibling settings; restricted atomic file", () => {
  const file = fixture(); const before = readFileSync(file, "utf8")
  patchEngineConfig({ model: "test/primary", small_model: "test/fast" })
  patchEngineConfig({ provider: { "google-vertex": { options: { project: "my-project" } } } })
  patchEngineConfig({ provider: { "google-vertex": { options: { location: "us-central1" } }, local: { npm: "@ai-sdk/openai-compatible", options: { baseURL: "http://localhost:1234/v1" }, models: { custom: { name: "Custom" } } } } })
  setAgentDisabled("custom", true); setSkillDisabled("my-skill", true)
  expect(readEngineConfig().provider?.["google-vertex"].options).toEqual({ project: "my-project", location: "us-central1" })
  expect(readEngineConfig().agent?.draft.prompt).toBe("matter policy")
  expect(readEngineConfig().permission?.external_directory).toBe("deny")
  expect(readFileSync(file, "utf8")).toBe(before)
  expect(statSync(engineConfigFile()).mode & 0o777).toBe(0o600)
  expect(statSync(path.dirname(engineConfigFile())).mode & 0o777).toBe(0o700)
  patchEngineConfig({ model: null, provider: { "google-vertex": { options: { project: null } } } })
  expect(readEngineConfig().model).toBeUndefined()
  expect(readEngineConfig().small_model).toBe("test/fast")
  expect(readEngineConfig().provider?.["google-vertex"].options).toEqual({ location: "us-central1" })
})
test("credentials, code, discovery and malformed provider input never reach disk", () => {
  fixture(); patchEngineConfig({ model: "test/selected" }); const before = readFileSync(engineConfigFile(), "utf8")
  for (const patch of [{ plugins: ["untrusted"] }, { instructions: ["/tmp/ambient"] }, { mcp: {} }, { shell: "sh" }, { model: "missing-provider" }, { disabled_providers: [1] }, { provider: null }, { provider: { local: { npm: "untrusted" } } }, { provider: { local: { options: { apiKey: "test-only-placeholder" } } } }, { provider: { local: { options: { baseURL: "https://user:pass@example.com" } } } }, { provider: { local: { models: { a: { request: {} } } } } }, { agent: { a: { prompt: "untrusted" } } }, JSON.parse('{"provider":{"__proto__":{}}}')]) expect(() => patchEngineConfig(patch)).toThrow()
  expect(readFileSync(engineConfigFile(), "utf8")).toBe(before)
})
test("fresh install has no implicit generation model", () => { fixture(); expect(readEngineConfig().model).toBeUndefined(); expect(readEngineConfig().small_model).toBeUndefined() })

test("removed provider cannot be newly selected or configured; Go and unrelated settings remain editable", () => {
  fixture(); patchEngineConfig({ model: "test/primary", small_model: "test/fast" })
  const before = readFileSync(engineConfigFile(), "utf8")
  for (const patch of [
    { model: "opencode/primary" }, { small_model: "opencode/fast" },
    { agent: { draft: { model: "opencode/agent" } } },
    { provider: { opencode: {} } }, { provider: { opencode: { models: { custom: { name: "Custom" } } } } },
  ]) {
    expect(() => patchEngineConfig(patch)).toThrow("no longer available")
    expect(readFileSync(engineConfigFile(), "utf8")).toBe(before)
  }
  patchEngineConfig({ model: "opencode-go/primary", small_model: "opencode-go/fast", agent: { draft: { model: "opencode-go/agent" } }, provider: { "opencode-go": { name: "OpenCode Go" } } })
  expect(readEngineSettings().model).toBe("opencode-go/primary")
  expect(readEngineSettings().provider?.["opencode-go"].name).toBe("OpenCode Go")
})

test("effective removal survives runtime overrides without rewriting existing choices or credentials", () => {
  const shipped = fixture()
  const saved = { model: "opencode/primary", small_model: "opencode/fast", disabled_providers: [], provider: { opencode: { options: { apiKey: "synthetic-existing-secret" } } } }
  writeEngineOverrides(saved)
  const before = readFileSync(engineConfigFile(), "utf8"), shippedBefore = readFileSync(shipped, "utf8")
  expect(readEngineConfig().disabled_providers).toEqual(["opencode"])
  expect(readEngineSettings().model).toBe("opencode/primary")
  expect(readFileSync(engineConfigFile(), "utf8")).toBe(before)
  expect(readFileSync(shipped, "utf8")).toBe(shippedBefore)
  expect(JSON.stringify(readEngineSettings())).not.toContain("synthetic-existing-secret")
  patchEngineConfig({ disabled_providers: ["other"], agent: { custom: { disable: true } } })
  expect(readEngineConfig().disabled_providers).toEqual(["other", "opencode"])
  expect(readEngineConfig().model).toBe("opencode/primary")
  expect(readEngineConfig().provider?.opencode.options).toEqual({ apiKey: "synthetic-existing-secret" })
  patchEngineConfig({ disabled_providers: null })
  expect(readEngineConfig().disabled_providers).toEqual(["opencode"])
})

test("browser projection omits shipped policy and credentials", () => {
  const file = fixture()
  writeFileSync(file, JSON.stringify({ plugins: ["private-path"], agent: { draft: { prompt: "private-prompt" } }, provider: { local: { options: { apiKey: "test-secret-placeholder", baseURL: "https://example.com" } } } }))
  const publicSettings = JSON.stringify(readEngineSettings())
  expect(publicSettings).not.toContain("private-")
  expect(publicSettings).not.toContain("test-secret-placeholder")
})
test("malformed saved overrides fail with a value-free repair message", () => {
  fixture(); patchEngineConfig({ model: "test/primary" }); writeFileSync(engineConfigFile(), '{"model":"test-secret-placeholder"')
  expect(() => readEngineConfig()).toThrow("Settings file is unreadable or contains invalid JSON")
})
