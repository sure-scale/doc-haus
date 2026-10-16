import { Database } from "bun:sqlite"
import { afterEach, expect, test } from "bun:test"
import { mkdtempSync, mkdirSync, rmSync, statSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { migrateLegacyContextBudget, readDetailPayload, readLeanPayload, readStoragePayload } from "./telemetry"
import { PRICING_SOURCE } from "../../../dochaus/lib/pricing"

const tempDirs: string[] = []
const rateOverrides = {
  "synthetic/priced": { input_per_mtok: 2, output_per_mtok: 6 },
  "synthetic/legacy": { input_per_mtok: 1.2, output_per_mtok: 4 },
}

afterEach(() => {
  for (const dir of tempDirs.splice(0)) rmSync(dir, { recursive: true, force: true })
})

test("readLeanPayload returns empty payload when DB is missing", () => {
  const dir = tempMatter()

  expect(readLeanPayload(dir)).toEqual({
    current_activity: null,
    today: { cost_usd: 0, turns: 0 },
    matter_total: { cost_usd: 0, turns: 0, sessions: 0, since_ts: 0 },
    current_session: null,
    rate_card_as_of: null,
    unknown_models: [],
  })
})

test("readLeanPayload sums current activity and matter total", () => {
  const dir = tempMatter()
  const now = Date.now()
  const olderTs = now - 3_600_000
  const db = seedDb(dir)
  insertSession(db, { id: "s1", lastActivityAt: now - 1_000, model: "gemini-2.5-pro", agent: "drafter" })
  insertTurn(db, { sessionId: "s1", ts: now - 1_000, costUsd: 0.25, tokens: 120_000, agent: "drafter" })
  insertTurn(db, { sessionId: "s1", ts: olderTs, costUsd: 0.75, tokens: 80_000, agent: "drafter" })
  db.close()
  const expectedToday =
    olderTs >= new Date().setHours(0, 0, 0, 0) ? { cost_usd: 0.4, turns: 2 } : { cost_usd: 0.24, turns: 1 }

  expect(readLeanPayload(dir, { session: "s1", rateOverrides })).toMatchObject({
    current_activity: { cost_usd: 0.24, turns: 1 },
    today: expectedToday,
    matter_total: { cost_usd: 0.4, turns: 2, sessions: 1 },
    current_session: { tokens: 120_000, status: "warn", agent: "drafter", model: "gemini-2.5-pro" },
    rate_card_as_of: null,
  })
})

test("readLeanPayload respects activity window", () => {
  const dir = tempMatter()
  const now = Date.now()
  const db = seedDb(dir)
  insertSession(db, { id: "s1", lastActivityAt: now - 5_000 })
  insertTurn(db, { sessionId: "s1", ts: now - 20_000, costUsd: 0.2 })
  db.close()

  expect(readLeanPayload(dir, { activityWindowMs: 10_000 }).current_activity).toBeNull()
})

test("readLeanPayload surfaces unknown models", () => {
  const dir = tempMatter()
  const db = seedDb(dir)
  db.query("INSERT INTO meta (key, value) VALUES (?, ?)").run("unknown_models", JSON.stringify(["model-b", "model-a"]))
  db.close()

  expect(readLeanPayload(dir).unknown_models).toEqual(["model-b", "model-a"])
})

test("read payloads omit formerly unknown models that now have rates", () => {
  const first = tempMatter()
  const second = tempMatter()
  const firstDb = seedDb(first)
  firstDb
    .query("INSERT INTO meta (key, value) VALUES (?, ?)")
    .run("unknown_models", JSON.stringify(["synthetic/priced", "unknown/model"]))
  firstDb.close()
  const secondDb = seedDb(second)
  secondDb
    .query("INSERT INTO meta (key, value) VALUES (?, ?)")
    .run("unknown_models", JSON.stringify(["synthetic/priced"]))
  secondDb.close()

  expect(readLeanPayload(first, { rateOverrides }).unknown_models).toEqual(["unknown/model"])
  expect(readStoragePayload([first, second], rateOverrides).unknown_models).toEqual(["unknown/model"])
})

test("read payloads estimate null costs for turns whose models now have rates", () => {
  const dir = tempMatter()
  const now = Date.now()
  const db = seedDb(dir)
  insertSession(db, { id: "s1", lastActivityAt: now - 1_000, model: "synthetic/priced", agent: "drafter" })
  insertTurn(db, {
    sessionId: "s1",
    ts: now - 1_000,
    costUsd: null,
    tokens: 1_000_000,
    agent: "drafter",
    model: "synthetic/priced",
  })
  db.close()

  expect(readLeanPayload(dir, { session: "s1", rateOverrides })).toMatchObject({
    current_activity: { cost_usd: 2, turns: 1 },
    today: { cost_usd: 2, turns: 1 },
    matter_total: { cost_usd: 2, turns: 1, sessions: 1 },
  })
  expect(readDetailPayload(dir, rateOverrides)).toMatchObject({
    by_model: [{ model: "synthetic/priced", cost_usd: 2, turns: 1 }],
    by_agent: [{ agent: "drafter", cost_usd: 2, turns: 1, tool_calls: 0 }],
    recent_sessions: [{ agent: "drafter", cost_usd: 2 }],
    sparkline: [{ cost_usd: 2 }],
  })
})

test("current overrides reprice historical estimates without changing original rows", () => {
  const dir = tempMatter()
  const db = seedDb(dir)
  insertTurn(db, { sessionId: "s1", ts: Date.now(), model: "synthetic/priced", tokens: 1_000_000, costUsd: 0.01 })
  db.close()

  expect(readLeanPayload(dir).unknown_models).toEqual(["synthetic/priced"])
  expect(readStoragePayload([dir]).unknown_models).toEqual(["synthetic/priced"])
  expect(readDetailPayload(dir).recent_turns[0]?.cost_usd).toBeNull()
  expect(readLeanPayload(dir, { rateOverrides }).matter_total.cost_usd).toBe(2)
  expect(readLeanPayload(dir, { rateOverrides }).unknown_models).toEqual([])
  expect(readDetailPayload(dir, { "synthetic/priced": { input_per_mtok: 3, output_per_mtok: 9 } }).recent_turns[0]?.cost_usd).toBe(3)
  expect(readDetailPayload(dir, { "synthetic/priced": { input_per_mtok: 3 } }).recent_turns[0]?.cost_usd).toBeNull()

  const original = new Database(path.join(dir, ".dochaus", "telemetry.db"), { readonly: true })
  expect(original.query("SELECT cost_usd FROM turns").get()).toEqual({ cost_usd: 0.01 })
  original.close()
})

test("migrateLegacyContextBudget imports pre-SQLite context snapshots once", () => {
  const dir = tempMatter()
  const db = seedDb(dir)
  insertTurn(db, {
    sessionId: "s-existing",
    ts: Date.parse("2026-01-01T00:10:00Z"),
    costUsd: 0.5,
    tokens: 500_000,
    agent: "lean-draft",
    model: "synthetic/priced",
  })
  db.close()
  writeFileSync(
    path.join(dir, ".dochaus", "context-budget.jsonl"),
    [
      JSON.stringify({
        created_at: "2026-01-01T00:00:00Z",
        chars: 4_000_000,
        estimated_tokens: 1_000_000,
        status: "ok",
        model: { providerID: "synthetic", modelID: "legacy" },
        largest_messages: [{ agent: "draft", chars: 4_000_000, estimated_tokens: 1_000_000 }],
      }),
      JSON.stringify({
        created_at: "2026-01-01T00:20:00Z",
        chars: 4_000_000,
        estimated_tokens: 1_000_000,
        status: "ok",
        model: { providerID: "synthetic", modelID: "legacy" },
        largest_messages: [{ agent: "draft", chars: 4_000_000, estimated_tokens: 1_000_000 }],
      }),
    ].join("\n") + "\n",
  )

  expect(migrateLegacyContextBudget(dir, rateOverrides)).toEqual({ inserted: 1, skipped: 1, backfilled: 0 })
  expect(migrateLegacyContextBudget(dir, rateOverrides)).toEqual({ inserted: 0, skipped: 0, backfilled: 0 })
  expect(readLeanPayload(dir, { rateOverrides }).matter_total).toMatchObject({ cost_usd: 2.2, turns: 2 })
  expect(readDetailPayload(dir, rateOverrides).by_agent).toContainEqual({ agent: "draft", cost_usd: 1.2, turns: 1, tool_calls: 0 })
})

test("migrateLegacyContextBudget backfills null costs for known models", () => {
  const dir = tempMatter()
  const db = seedDb(dir)
  insertTurn(db, {
    sessionId: "s-existing",
    ts: Date.parse("2026-01-01T00:10:00Z"),
    costUsd: null,
    tokens: 1_000_000,
    agent: "lean-draft",
    model: "synthetic/priced",
  })
  writeFileSync(path.join(dir, ".dochaus", "context-budget.jsonl"), "")
  db.close()

  expect(migrateLegacyContextBudget(dir, rateOverrides)).toEqual({ inserted: 0, skipped: 0, backfilled: 1 })
  const migrated = new Database(path.join(dir, ".dochaus", "telemetry.db"), { readonly: true })
  const row = migrated.query("SELECT cost_usd FROM turns").get() as { cost_usd: number }
  expect(row.cost_usd).toBe(2)
  migrated.close()
})

test("readDetailPayload groups by agent and tool", () => {
  const dir = tempMatter()
  const now = Date.now()
  const db = seedDb(dir)
  insertSession(db, { id: "s1", lastActivityAt: now - 1_000, agent: "drafter" })
  insertTurn(db, { sessionId: "s1", ts: now - 1_000, costUsd: 0.25, tokens: 125_000, agent: "drafter" })
  insertTurn(db, { sessionId: "s1", ts: now - 2_000, costUsd: null, agent: "drafter", model: "synthetic/unpriced" })
  insertToolCall(db, { sessionId: "s1", ts: now - 1_000, tool: "search-document", agent: "drafter", durationMs: 20 })
  insertToolCall(db, { sessionId: "s1", ts: now - 2_000, tool: "search-document", agent: "drafter", durationMs: 40 })
  db.close()

  expect(readDetailPayload(dir, rateOverrides).by_agent).toEqual([{ agent: "drafter", cost_usd: 0.25, turns: 2, tool_calls: 2 }])
  expect(readDetailPayload(dir, rateOverrides).by_tool).toEqual([
    { tool: "search-document", calls: 2, total_duration_ms: 60, avg_duration_ms: 30 },
  ])
})

test("readStoragePayload sums DB sizes across matters", () => {
  const first = tempMatter()
  const second = tempMatter()
  const db = seedDb(first)
  insertTurn(db, { sessionId: "s1", ts: Date.now() - 1_000, costUsd: 0.1 })
  db.query("INSERT INTO meta (key, value) VALUES (?, ?)").run("unknown_models", JSON.stringify(["unknown-model"]))
  db.close()

  const storage = readStoragePayload([first, second], rateOverrides)

  expect(storage.size_bytes).toBeGreaterThan(0)
  expect(storage.matter_count).toBe(2)
  expect(storage.unknown_models).toEqual(["unknown-model"])
  expect(storage.rate_card_as_of).toBeNull()
  expect(storage.rate_source).toBe(PRICING_SOURCE)
})

test("readStoragePayload includes WAL sibling file sizes", () => {
  const dir = tempMatter()
  const db = seedDb(dir)
  db.close()
  const base = path.join(dir, ".dochaus", "telemetry.db")
  writeFileSync(`${base}-wal`, "wal")
  writeFileSync(`${base}-shm`, "shm-data")
  const expectedSize = statSync(base).size + statSync(`${base}-wal`).size + statSync(`${base}-shm`).size

  expect(readStoragePayload([dir]).size_bytes).toBe(expectedSize)
})

test("read payloads return safe empty values for invalid telemetry DBs", () => {
  const dir = tempMatter()
  mkdirSync(path.join(dir, ".dochaus"), { recursive: true })
  new Database(path.join(dir, ".dochaus", "telemetry.db")).close()

  expect(readLeanPayload(dir)).toEqual({
    current_activity: null,
    today: { cost_usd: 0, turns: 0 },
    matter_total: { cost_usd: 0, turns: 0, sessions: 0, since_ts: 0 },
    current_session: null,
    rate_card_as_of: null,
    unknown_models: [],
  })
  expect(readDetailPayload(dir)).toEqual({ by_agent: [], by_model: [], by_tool: [], recent_turns: [], recent_sessions: [], sparkline: [] })
  expect(readStoragePayload([dir])).toMatchObject({
    matter_count: 1,
    oldest_ts: null,
    unknown_models: [],
    rate_card_as_of: null,
  })
})

test("read payloads return safe values for garbage telemetry DB files", () => {
  const dir = tempMatter()
  mkdirSync(path.join(dir, ".dochaus"), { recursive: true })
  const file = path.join(dir, ".dochaus", "telemetry.db")
  writeFileSync(file, "not sqlite")

  expect(readLeanPayload(dir)).toEqual({
    current_activity: null,
    today: { cost_usd: 0, turns: 0 },
    matter_total: { cost_usd: 0, turns: 0, sessions: 0, since_ts: 0 },
    current_session: null,
    rate_card_as_of: null,
    unknown_models: [],
  })
  expect(readDetailPayload(dir)).toEqual({ by_agent: [], by_model: [], by_tool: [], recent_turns: [], recent_sessions: [], sparkline: [] })
  expect(readStoragePayload([dir])).toMatchObject({
    size_bytes: statSync(file).size,
    matter_count: 1,
    oldest_ts: null,
    unknown_models: [],
    rate_card_as_of: null,
  })
})

function tempMatter() {
  const dir = mkdtempSync(path.join(tmpdir(), "dochaus-telemetry-"))
  tempDirs.push(dir)
  return dir
}

function seedDb(matterDir: string) {
  mkdirSync(path.join(matterDir, ".dochaus"), { recursive: true })
  const db = new Database(path.join(matterDir, ".dochaus", "telemetry.db"))
  db.exec(`
    CREATE TABLE sessions (
      id TEXT PRIMARY KEY,
      parent_id TEXT,
      created_at INTEGER NOT NULL,
      last_activity_at INTEGER NOT NULL,
      status TEXT NOT NULL DEFAULT 'active',
      model TEXT,
      agent TEXT
    );
    CREATE TABLE turns (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      session_id TEXT NOT NULL,
      turn_index INTEGER NOT NULL,
      ts INTEGER NOT NULL,
      model TEXT,
      agent TEXT,
      context_chars INTEGER NOT NULL,
      estimated_input_tokens INTEGER NOT NULL,
      input_tokens INTEGER,
      output_tokens INTEGER,
      cache_read_tokens INTEGER,
      reasoning_tokens INTEGER,
      cost_usd REAL,
      status TEXT NOT NULL DEFAULT 'ok'
    );
    CREATE TABLE tool_calls (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      session_id TEXT NOT NULL,
      turn_index INTEGER,
      ts INTEGER NOT NULL,
      tool TEXT NOT NULL,
      agent TEXT,
      duration_ms INTEGER,
      output_chars INTEGER,
      phase TEXT NOT NULL DEFAULT 'ok'
    );
    CREATE TABLE meta (
      key TEXT PRIMARY KEY,
      value TEXT NOT NULL
    );
  `)
  db.query("INSERT INTO meta (key, value) VALUES (?, ?)").run("pricing_as_of", "2026-01-01")
  return db
}

function insertSession(
  db: Database,
  input: { id: string; lastActivityAt: number; model?: string; agent?: string; status?: string },
) {
  db.query(`
    INSERT INTO sessions (id, parent_id, created_at, last_activity_at, status, model, agent)
    VALUES (?, NULL, ?, ?, ?, ?, ?)
  `).run(input.id, input.lastActivityAt, input.lastActivityAt, input.status ?? "active", input.model ?? null, input.agent ?? null)
}

function insertTurn(
  db: Database,
  input: {
    sessionId: string
    ts: number
    costUsd: number | null
    tokens?: number
    agent?: string
    model?: string
    status?: string
  },
) {
  db.query(`
    INSERT INTO turns (
      session_id,
      turn_index,
      ts,
      model,
      agent,
      context_chars,
      estimated_input_tokens,
      cost_usd,
      status
    )
    VALUES (?, 0, ?, ?, ?, 0, ?, ?, ?)
  `).run(
    input.sessionId,
    input.ts,
    input.model ?? "synthetic/priced",
    input.agent ?? "drafter",
    input.tokens ?? 1_000,
    input.costUsd,
    input.status ?? "ok",
  )
}

function insertToolCall(
  db: Database,
  input: { sessionId: string; ts: number; tool: string; agent: string; durationMs: number | null },
) {
  db.query(`
    INSERT INTO tool_calls (session_id, turn_index, ts, tool, agent, duration_ms, output_chars, phase)
    VALUES (?, 0, ?, ?, ?, ?, 0, 'ok')
  `).run(input.sessionId, input.ts, input.tool, input.agent, input.durationMs)
}
