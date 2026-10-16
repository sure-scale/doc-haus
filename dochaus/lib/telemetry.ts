import { Database } from "bun:sqlite"
import { mkdirSync } from "node:fs"
import { join } from "node:path"
import { computeCost, getRate, PRICING_SOURCE, type RateOverrides } from "./pricing"

export type TurnInput = {
  sessionID: string
  parentID?: string
  turnIndex: number
  model: string
  agent: string
  contextChars: number
  estimatedInputTokens: number
  rateOverrides?: RateOverrides
}

export type ToolCallInput = {
  sessionID: string
  turnIndex: number | null
  tool: string
  agent: string
  durationMs: number | null
  outputChars: number | null
  phase: "ok" | "error"
}

export type SessionTouch = {
  parentID?: string
  model?: string
  agent?: string
  status?: "active" | "idle" | "finished"
}

export type MessageLike = {
  parts?: unknown
  [key: string]: unknown
}

export function openTelemetryDb(matterDir: string) {
  mkdirSync(join(matterDir, ".dochaus"), { recursive: true })
  const db = new Database(join(matterDir, ".dochaus", "telemetry.db"))
  db.exec("PRAGMA journal_mode = WAL")
  db.exec(`
    CREATE TABLE IF NOT EXISTS sessions (
      id TEXT PRIMARY KEY,
      parent_id TEXT,
      created_at INTEGER NOT NULL,
      last_activity_at INTEGER NOT NULL,
      status TEXT NOT NULL DEFAULT 'active',
      model TEXT,
      agent TEXT
    );
    CREATE TABLE IF NOT EXISTS turns (
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
    CREATE TABLE IF NOT EXISTS tool_calls (
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
    CREATE TABLE IF NOT EXISTS meta (
      key TEXT PRIMARY KEY,
      value TEXT NOT NULL
    );
    CREATE INDEX IF NOT EXISTS turns_session_ts ON turns (session_id, ts);
    CREATE INDEX IF NOT EXISTS turns_ts ON turns (ts);
    CREATE INDEX IF NOT EXISTS tool_calls_session_ts ON tool_calls (session_id, ts);
    CREATE INDEX IF NOT EXISTS sessions_last_activity ON sessions (last_activity_at);
  `)
  db.query("INSERT OR IGNORE INTO meta (key, value) VALUES (?, ?)").run("schema_version", "1")
  db.query("INSERT OR IGNORE INTO meta (key, value) VALUES (?, ?)").run("created_at", String(Date.now()))
  db.query("INSERT INTO meta (key, value) VALUES (?, ?) ON CONFLICT(key) DO UPDATE SET value = excluded.value").run(
    "pricing_source",
    PRICING_SOURCE,
  )
  db.query("DELETE FROM meta WHERE key = ?").run("pricing_as_of")
  return db
}

export function touchSession(db: Database, sessionID: string, patch: SessionTouch) {
  const now = Date.now()
  db.query(`
    INSERT OR IGNORE INTO sessions (id, parent_id, created_at, last_activity_at, status, model, agent)
    VALUES (?, ?, ?, ?, ?, ?, ?)
  `).run(sessionID, patch.parentID ?? null, now, now, patch.status ?? "active", patch.model ?? null, patch.agent ?? null)
  db.query(`
    UPDATE sessions SET
      parent_id = COALESCE(?, parent_id),
      last_activity_at = ?,
      status = COALESCE(?, status),
      model = COALESCE(?, model),
      agent = COALESCE(?, agent)
    WHERE id = ?
  `).run(patch.parentID ?? null, now, patch.status ?? null, patch.model ?? null, patch.agent ?? null, sessionID)
}

export function recordTurn(db: Database, input: TurnInput) {
  const rate = getRate(input.model, input.rateOverrides, input.estimatedInputTokens)
  const costUsd = rate ? computeCost(input.estimatedInputTokens, rate) : null
  const unknownModel = !rate
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
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
  `).run(
    input.sessionID,
    input.turnIndex,
    Date.now(),
    input.model,
    input.agent,
    input.contextChars,
    input.estimatedInputTokens,
    costUsd,
    turnStatus(input.estimatedInputTokens),
  )
  touchSession(db, input.sessionID, {
    parentID: input.parentID,
    model: input.model,
    agent: input.agent,
    status: "active",
  })
  if (unknownModel) recordUnknownModel(db, input.model)
  return { costUsd, unknownModel }
}

export function recordToolCall(db: Database, input: ToolCallInput) {
  db.query(`
    INSERT INTO tool_calls (session_id, turn_index, ts, tool, agent, duration_ms, output_chars, phase)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
  `).run(
    input.sessionID,
    input.turnIndex,
    Date.now(),
    input.tool,
    input.agent,
    input.durationMs,
    input.outputChars,
    input.phase,
  )
  touchSession(db, input.sessionID, { agent: input.agent, status: "active" })
}

export function markSessionStatus(db: Database, sessionID: string, status: "active" | "idle" | "finished") {
  touchSession(db, sessionID, { status })
}

export function applyRetentionPolicy(db: Database, retentionDays: number) {
  const cutoff = Date.now() - retentionDays * 86_400_000
  const turns = db.query("DELETE FROM turns WHERE ts < ?").run(cutoff).changes
  const toolCalls = db.query("DELETE FROM tool_calls WHERE ts < ?").run(cutoff).changes
  const sessions = db
    .query("DELETE FROM sessions WHERE last_activity_at < ? AND status = 'finished'")
    .run(cutoff).changes
  return { turns, toolCalls, sessions }
}

export function clearTelemetry(db: Database) {
  db.query("DELETE FROM turns").run()
  db.query("DELETE FROM tool_calls").run()
  db.query("DELETE FROM sessions").run()
  db.query("DELETE FROM meta WHERE key = ?").run("unknown_models")
}

export function summarizeMessages(messages: MessageLike[]) {
  const chars = messages.reduce((sum, message) => sum + textLength(message.parts), 0)
  return { chars, estimatedTokens: Math.ceil(chars / 4) }
}

export function getUnknownModels(db: Database, rateOverrides?: RateOverrides) {
  const row = db.query("SELECT value FROM meta WHERE key = ?").get("unknown_models") as { value: string } | null
  if (!row) return []
  const parsed = parseJson(row.value)
  if (!Array.isArray(parsed)) return []
  return parsed.filter((value): value is string => typeof value === "string" && !getRate(value, rateOverrides))
}

export function getPricingSource(db: Database) {
  const row = db.query("SELECT value FROM meta WHERE key = ?").get("pricing_source") as { value: string } | null
  return row?.value ?? null
}

function turnStatus(estimatedInputTokens: number) {
  if (estimatedInputTokens >= 180_000) return "fail"
  if (estimatedInputTokens >= 90_000) return "warn"
  return "ok"
}

function recordUnknownModel(db: Database, model: string) {
  const models = Array.from(new Set([...getUnknownModels(db), model])).sort()
  db.query("INSERT INTO meta (key, value) VALUES (?, ?) ON CONFLICT(key) DO UPDATE SET value = excluded.value").run(
    "unknown_models",
    JSON.stringify(models),
  )
}

function parseJson(value: string) {
  try {
    return JSON.parse(value) as unknown
  } catch {
    return undefined
  }
}

function textLength(value: unknown): number {
  if (typeof value === "string") return value.length
  if (Array.isArray(value)) return value.reduce((sum, item) => sum + textLength(item), 0)
  if (!isRecord(value)) return 0
  return (
    textLength(value.text) +
    textLength(value.content) +
    textLength(value.state) +
    textLength(value.input) +
    textLength(value.output) +
    textLength(value.result)
  )
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null
}
