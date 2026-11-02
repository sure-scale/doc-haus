import { afterEach, expect, test } from "bun:test"
import { mkdtempSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import { join } from "node:path"
import {
  applyRetentionPolicy,
  clearTelemetry,
  markSessionStatus,
  openTelemetryDb,
  readGenerations,
  recordGeneration,
  recordToolCall,
  recordTurn,
  summarizeMessages,
  touchSession,
} from "./telemetry"

let cleanup: (() => void) | undefined

afterEach(() => {
  cleanup?.()
  cleanup = undefined
})

function tempMatter() {
  const dir = mkdtempSync(join(tmpdir(), "telemetry-test-"))
  cleanup = () => rmSync(dir, { recursive: true, force: true })
  return dir
}

test("openTelemetryDb creates the schema idempotently", () => {
  const matter = tempMatter()
  const db = openTelemetryDb(matter)
  db.query("INSERT INTO meta (key, value) VALUES (?, ?)").run("pricing_as_of", "2026-01-01")
  db.close()

  const reopened = openTelemetryDb(matter)
  expect(
    reopened
      .query("SELECT name FROM sqlite_master WHERE type = 'table' AND name NOT LIKE 'sqlite_%' ORDER BY name")
      .all()
      .map((row) => row.name),
  ).toEqual(["generations", "meta", "sessions", "tool_calls", "turns"])
  expect(reopened.query("SELECT value FROM meta WHERE key = ?").get("pricing_as_of")).toBeNull()
  reopened.close()
})

test("touchSession upserts and updates last_activity_at", () => {
  const db = openTelemetryDb(tempMatter())
  touchSession(db, "session-1", { agent: "draft" })
  const first = db.query("SELECT last_activity_at FROM sessions WHERE id = ?").get("session-1").last_activity_at

  touchSession(db, "session-1", { status: "idle" })
  touchSession(db, "session-1", { model: "google-vertex/gemini-3.5-flash" })
  const second = db.query("SELECT last_activity_at, agent, model, status FROM sessions WHERE id = ?").get("session-1")

  expect(second.last_activity_at).toBeGreaterThanOrEqual(first)
  expect(second.agent).toBe("draft")
  expect(second.model).toBe("google-vertex/gemini-3.5-flash")
  expect(second.status).toBe("idle")
  db.close()
})

test("recordTurn keeps the context size check and leaves cost to native sessions", () => {
  const db = openTelemetryDb(tempMatter())
  recordTurn(db, {
    sessionID: "session-1",
    turnIndex: 3,
    model: "synthetic/model",
    agent: "draft",
    contextChars: 40_000,
    estimatedInputTokens: 10_000,
  })
  const row = db.query("SELECT turn_index, estimated_input_tokens, cost_usd FROM turns").get()

  expect(row).toEqual({ turn_index: 3, estimated_input_tokens: 10_000, cost_usd: null })
  db.close()
})

test("recordToolCall inserts with duration and output size", () => {
  const db = openTelemetryDb(tempMatter())
  recordToolCall(db, {
    sessionID: "session-1",
    turnIndex: 1,
    tool: "search-document",
    agent: "draft",
    durationMs: 320,
    outputChars: 5000,
    phase: "ok",
  })
  const row = db.query("SELECT tool, duration_ms, output_chars, phase FROM tool_calls").get()

  expect(row.tool).toBe("search-document")
  expect(row.duration_ms).toBe(320)
  expect(row.output_chars).toBe(5000)
  expect(row.phase).toBe("ok")
  db.close()
})

test("markSessionStatus flips status", () => {
  const db = openTelemetryDb(tempMatter())
  touchSession(db, "session-1", { status: "active" })
  markSessionStatus(db, "session-1", "idle")

  expect(db.query("SELECT status FROM sessions WHERE id = ?").get("session-1").status).toBe("idle")
  db.close()
})

test("applyRetentionPolicy deletes old rows but keeps recent ones", () => {
  const db = openTelemetryDb(tempMatter())
  const now = Date.now()
  db.query(
    "INSERT INTO turns (session_id, turn_index, ts, model, agent, context_chars, estimated_input_tokens) VALUES (?, ?, ?, ?, ?, ?, ?)",
  ).run("old", 1, now - 100 * 86_400_000, "model", "agent", 10, 3)
  db.query(
    "INSERT INTO turns (session_id, turn_index, ts, model, agent, context_chars, estimated_input_tokens) VALUES (?, ?, ?, ?, ?, ?, ?)",
  ).run("recent", 1, now - 1 * 86_400_000, "model", "agent", 10, 3)
  db.query(
    "INSERT INTO tool_calls (session_id, turn_index, ts, tool, agent, duration_ms, output_chars, phase) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
  ).run("old", 1, now - 100 * 86_400_000, "search-document", "agent", 10, 20, "ok")
  db.query(
    "INSERT INTO tool_calls (session_id, turn_index, ts, tool, agent, duration_ms, output_chars, phase) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
  ).run("recent", 1, now - 1 * 86_400_000, "search-document", "agent", 10, 20, "ok")
  db.query(
    "INSERT INTO sessions (id, created_at, last_activity_at, status) VALUES (?, ?, ?, ?)",
  ).run("old-finished", now - 100 * 86_400_000, now - 100 * 86_400_000, "finished")
  db.query(
    "INSERT INTO sessions (id, created_at, last_activity_at, status) VALUES (?, ?, ?, ?)",
  ).run("old-active", now - 100 * 86_400_000, now - 100 * 86_400_000, "active")
  db.query(
    "INSERT INTO sessions (id, created_at, last_activity_at, status) VALUES (?, ?, ?, ?)",
  ).run("old-idle", now - 100 * 86_400_000, now - 100 * 86_400_000, "idle")
  db.query(
    "INSERT INTO sessions (id, created_at, last_activity_at, status) VALUES (?, ?, ?, ?)",
  ).run("recent-finished", now - 1 * 86_400_000, now - 1 * 86_400_000, "finished")

  const counts = applyRetentionPolicy(db, 90)

  expect(counts.turns).toBe(1)
  expect(counts.toolCalls).toBe(1)
  expect(counts.sessions).toBe(1)
  expect(db.query("SELECT COUNT(*) AS count FROM turns").get().count).toBe(1)
  expect(db.query("SELECT session_id FROM turns").get().session_id).toBe("recent")
  expect(db.query("SELECT COUNT(*) AS count FROM tool_calls").get().count).toBe(1)
  expect(db.query("SELECT session_id FROM tool_calls").get().session_id).toBe("recent")
  expect(
    db
      .query("SELECT id FROM sessions ORDER BY id")
      .all()
      .map((row) => row.id),
  ).toEqual(["old-active", "old-idle", "recent-finished"])
  db.close()
})

test("clearTelemetry wipes rows but keeps the file open", () => {
  const db = openTelemetryDb(tempMatter())
  touchSession(db, "session-1", {})
  recordTurn(db, {
    sessionID: "session-1",
    turnIndex: 1,
    model: "google-vertex/gemini-3.5-flash",
    agent: "draft",
    contextChars: 100,
    estimatedInputTokens: 25,
  })

  clearTelemetry(db)

  expect(db.query("SELECT COUNT(*) AS count FROM turns").get().count).toBe(0)
  expect(db.query("SELECT COUNT(*) AS count FROM sessions").get().count).toBe(0)
  db.close()
})

test("summarizeMessages estimates tokens as chars/4", () => {
  expect(
    summarizeMessages([
      {
        info: {
          type: "text",
          agent: "draft",
          model: { providerID: "google-vertex", modelID: "gemini-3.5-flash" },
        },
        parts: [{ text: "a".repeat(400) }],
      },
    ]),
  ).toEqual({ chars: 400, estimatedTokens: 100 })
})

test("stateless generation usage round-trips and follows retention and clear", () => {
  const db = openTelemetryDb(tempMatter())
  try {
    const tokens = { input: 1200, output: 300, reasoning: 50, cache: { read: 400, write: 0 } }
    recordGeneration(db, { agent: "prepare-changes", model: "google/gemini", tokens, cost: 0.0123, ts: Date.now() - 10 * 86_400_000 })
    recordGeneration(db, { agent: "comment-replies", tokens, cost: null })
    expect(readGenerations(db).map((row) => [row.agent, row.model, row.cost, row.tokens])).toEqual([
      ["prepare-changes", "google/gemini", 0.0123, tokens],
      ["comment-replies", "", null, tokens],
    ])
    applyRetentionPolicy(db, 5)
    expect(readGenerations(db).map((row) => row.agent)).toEqual(["comment-replies"])
    clearTelemetry(db)
    expect(readGenerations(db)).toEqual([])
  } finally {
    db.close()
  }
})
