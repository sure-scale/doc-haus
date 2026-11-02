import { afterEach, expect, test } from "bun:test"
import { mkdtempSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import { join } from "node:path"
import { openTelemetryDb, recordGeneration } from "../../../dochaus/lib/telemetry"
import { generationUsage } from "./matter-usage"

let matter: string | undefined
afterEach(() => { if (matter) rmSync(matter, { recursive: true, force: true }); matter = undefined })

test("stored usage of a still-live session is skipped; deleted sessions and stateless jobs count", () => {
  matter = mkdtempSync(join(tmpdir(), "matter-usage-test-"))
  const tokens = { input: 3, output: 1, reasoning: 0, cache: { read: 0, write: 0 } }
  const db = openTelemetryDb(matter)
  try {
    recordGeneration(db, { agent: "prepare-changes", model: "fixture/note", tokens, cost: 0.01, ts: 1 })
    recordGeneration(db, { agent: "qa", model: "fixture/note", tokens, cost: 0.02, ts: 2, session: "ses_deleted" })
    recordGeneration(db, { agent: "review-compare", tokens, cost: 0.03, ts: 3, session: "ses_live" })
  } finally { db.close() }
  expect(generationUsage(matter, [{ id: "ses_live" }])).toEqual([
    { time: 1, cost: 0.01, agent: "prepare-changes", model: "fixture/note", tokens },
    { time: 2, cost: 0.02, session: "ses_deleted", agent: "qa", model: "fixture/note", tokens },
  ])
  // Without a live session list (the engine could not be read) every stored row counts.
  expect(generationUsage(matter).map(entry => entry.cost)).toEqual([0.01, 0.02, 0.03])
})
