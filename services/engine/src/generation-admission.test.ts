import { expect, test } from "bun:test"
import { mkdtempSync, mkdirSync, rmSync, writeFileSync } from "node:fs"
import os from "node:os"
import path from "node:path"
import { generationAdmissionError } from "./generation-admission"

test("generation admission reads the current effective pair and optional setup state without fallback", () => {
  const root = mkdtempSync(path.join(os.tmpdir(), "dochaus-admission-"))
  const paths = { shipped: path.join(root, "shipped.json"), overrides: path.join(root, "runtime", "opencode.json") }
  mkdirSync(path.dirname(paths.overrides))
  const save = (value: object) => writeFileSync(paths.overrides, JSON.stringify(value))
  try {
    writeFileSync(paths.shipped, "{}")
    expect(generationAdmissionError(paths)).toContain("primary and fast")
    save({ model: "neutral/primary" })
    expect(generationAdmissionError(paths)).toContain("primary and fast")
    save({ model: "neutral/primary", small_model: "other/fast" })
    expect(generationAdmissionError(paths)).toContain("same provider")
    save({ model: "neutral/primary", small_model: "neutral/fast" })
    expect(generationAdmissionError(paths)).toBeUndefined()
    writeFileSync(path.join(path.dirname(paths.overrides), "setup.json"), JSON.stringify({ version: 1, completed: ["models"], skipped: [] }))
    expect(generationAdmissionError(paths)).toContain("Finish installation setup")
    writeFileSync(path.join(path.dirname(paths.overrides), "setup.json"), JSON.stringify({ version: 1, completed: ["models", "details"], skipped: ["research"] }))
    expect(generationAdmissionError(paths)).toBeUndefined()
    save({ model: "neutral/primary", small_model: "neutral/fast", disabled_providers: ["neutral"] })
    expect(generationAdmissionError(paths)).toContain("disabled")
    writeFileSync(paths.overrides, '{"secret":"synthetic-private-key",broken')
    expect(generationAdmissionError(paths)).toContain("could not be read")
    expect(generationAdmissionError(paths)).not.toContain("synthetic-private-key")
  } finally {
    rmSync(root, { recursive: true, force: true })
  }
})
