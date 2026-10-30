import { expect, test } from "bun:test"
import { mkdtempSync, readFileSync, rmSync, statSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { migrateSettings } from "../../../scripts/migrate-settings"
import { writeEngineOverrides } from "./engine-config"
test("explicit migration dry-run, secure backup, existing wins and idempotent apply", () => {
  const root = mkdtempSync(path.join(tmpdir(), "settings-migration-"))
  try {
    const source = path.join(root, "legacy.json"), runtime = path.join(root, "runtime")
    const original = JSON.stringify({ model: "legacy/primary", small_model: "legacy/fast", model_concurrency: 50, plugins: ["private"], provider: { "google-vertex": { options: { project: "project-value", apiKey: "test-secret-placeholder" } } } })
    writeFileSync(source, original)
    const dry = migrateSettings({ source, runtime })
    expect(dry.status).toBe("dry-run"); expect(JSON.stringify(dry)).not.toContain("test-secret-placeholder"); expect(JSON.stringify(dry)).not.toContain("project-value")
    expect(dry.fields).toContain("model_concurrency")
    const target = path.join(runtime, "config", "opencode", "opencode.json")
    writeEngineOverrides({ model: "existing/primary" }, target)
    expect(migrateSettings({ source, runtime, apply: true }).status).toBe("migrated")
    expect(JSON.parse(readFileSync(target, "utf8"))).toEqual({ model: "existing/primary", small_model: "legacy/fast", model_concurrency: 50, provider: { "google-vertex": { options: { project: "project-value" } } } })
    const secondRuntime = path.join(root, "existing-runtime")
    const secondTarget = path.join(secondRuntime, "config", "opencode", "opencode.json")
    writeEngineOverrides({ model_concurrency: 10 }, secondTarget)
    migrateSettings({ source, runtime: secondRuntime, apply: true })
    expect(JSON.parse(readFileSync(secondTarget, "utf8")).model_concurrency).toBe(10)
    const backup = path.join(path.dirname(target), "legacy-settings.backup.json")
    expect(readFileSync(backup, "utf8")).toBe(original); expect(statSync(backup).mode & 0o777).toBe(0o600)
    expect(migrateSettings({ source, runtime, apply: true }).status).toBe("already-migrated")
    expect(readFileSync(source, "utf8")).toBe(original)
    expect(() => migrateSettings({ source: "relative", runtime })).toThrow()
  } finally { rmSync(root, { recursive: true, force: true }) }
})
