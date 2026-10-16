import { Database } from "bun:sqlite"
import { expect, test } from "bun:test"
import { mkdtempSync, mkdirSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { recordContextBudget, recordTelemetrySession } from "../plugin/legal/telemetry"

test("the legal telemetry hook applies saved custom rate overrides on the next turn", () => {
  const workspace = mkdtempSync(path.join(tmpdir(), "telemetry-rate-hook-"))
  const previous = process.env.WORKSPACE_ROOT
  process.env.WORKSPACE_ROOT = workspace
  try {
    const directory = path.join(workspace, "matter")
    const sessionID = "synthetic-rates-session"
    const model = "synthetic/custom"
    recordTelemetrySession(sessionID, { directory, model, agent: "draft" })
    const messages = [{ info: { sessionID }, parts: [{ text: "x".repeat(400) }] }]
    recordContextBudget(directory, messages)
    mkdirSync(path.join(workspace, ".preferences"))
    writeFileSync(path.join(workspace, ".preferences", "preferences.json"), JSON.stringify({
      telemetry: { rate_overrides: { [model]: { input_per_mtok: 2, output_per_mtok: 6 } } },
    }))
    recordContextBudget(directory, messages)
    const db = new Database(path.join(directory, ".dochaus", "telemetry.db"), { readonly: true })
    expect(db.query("SELECT estimated_input_tokens, cost_usd FROM turns ORDER BY id").all()).toEqual([
      { estimated_input_tokens: 100, cost_usd: null },
      { estimated_input_tokens: 100, cost_usd: 0.0002 },
    ])
    db.close()
  } finally {
    if (previous === undefined) delete process.env.WORKSPACE_ROOT
    else process.env.WORKSPACE_ROOT = previous
    rmSync(workspace, { recursive: true, force: true })
  }
})
