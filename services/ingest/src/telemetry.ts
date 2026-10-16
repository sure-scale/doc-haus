import { Database, type SQLQueryBindings } from "bun:sqlite"
import { existsSync, readFileSync, statSync } from "node:fs"
import path from "node:path"
import { computeCost, getRate, PRICING_SOURCE, PRICING_REFERENCE_URL, type RateOverrides } from "../../../dochaus/lib/pricing"

export type LeanPayload = {
  current_activity: { cost_usd: number; turns: number; since_ts: number } | null
  today: { cost_usd: number; turns: number }
  matter_total: { cost_usd: number; turns: number; sessions: number; since_ts: number }
  current_session: { tokens: number; status: "ok" | "warn" | "fail"; agent: string; model: string } | null
  rate_card_as_of: string | null
  unknown_models: string[]
}

export type DetailPayload = {
  by_agent: Array<{ agent: string; cost_usd: number; turns: number; tool_calls: number }>
  by_model: Array<{ model: string; cost_usd: number; turns: number }>
  by_tool: Array<{ tool: string; calls: number; total_duration_ms: number; avg_duration_ms: number }>
  recent_turns: Array<{ ts: number; agent: string; model: string; tokens: number; cost_usd: number | null; status: string }>
  recent_sessions: Array<{ id: string; agent: string; status: string; last_activity_at: number; cost_usd: number }>
  sparkline: Array<{ ts: number; cost_usd: number }>
}

export type StoragePayload = {
  size_bytes: number
  matter_count: number
  oldest_ts: number | null
  unknown_models: string[]
  rate_card_as_of: string | null
  rate_source: string
  rate_reference_url: string
}

type SumRow = { turns: number; cost_usd: number | null; since_ts?: number | null }
type CountRow = { sessions: number }
type MetaRow = { value: string } | null
type SessionRow = { id: string; model: string | null; agent: string | null } | null
type CostRow = {
  ts?: number
  session_id?: string
  agent?: string | null
  model: string | null
  estimated_input_tokens: number
  cost_usd: number | null
}
type LegacyContextBudgetRow = {
  ts: number
  model: string
  agent: string
  context_chars: number
  estimated_input_tokens: number
  status: string
}
type TurnRow = {
  ts: number
  agent: string | null
  model: string | null
  tokens: number
  cost_usd: number | null
  status: string
} | null

export function readLeanPayload(matterDir: string, opts?: { activityWindowMs?: number; session?: string; rateOverrides?: RateOverrides }): LeanPayload {
  const db = openReadOnly(matterDir)
  if (!db) return emptyLeanPayload()
  try {
    if (!hasTelemetrySchema(db)) return emptyLeanPayload()
    const now = Date.now()
    const activityWindowMs = opts?.activityWindowMs ?? 30 * 60_000
    const current = summarizeCostRows(readCostRows(db, "WHERE ts >= ?", [now - activityWindowMs]), opts?.rateOverrides)
    const today = summarizeCostRows(readCostRows(db, "WHERE ts >= ?", [new Date().setHours(0, 0, 0, 0)]), opts?.rateOverrides)
    const total = summarizeCostRows(readCostRows(db), opts?.rateOverrides)
    const sessions = db.query("SELECT COUNT(*) AS sessions FROM sessions").get() as CountRow
    const currentSession = readCurrentSession(db, opts?.session)
    return {
      current_activity:
        current.turns > 0 ? { cost_usd: current.cost_usd ?? 0, turns: current.turns, since_ts: now - activityWindowMs } : null,
      today: { cost_usd: today.cost_usd ?? 0, turns: today.turns },
      matter_total: {
        cost_usd: total.cost_usd ?? 0,
        turns: total.turns,
        sessions: sessions.sessions,
        since_ts: total.since_ts ?? 0,
      },
      current_session: currentSession,
      rate_card_as_of: null,
      unknown_models: readUnknownModels(db, opts?.rateOverrides),
    }
  } catch {
    return emptyLeanPayload()
  } finally {
    db.close()
  }
}

export function readDetailPayload(matterDir: string, rateOverrides?: RateOverrides): DetailPayload {
  const db = openReadOnly(matterDir)
  if (!db) return emptyDetailPayload()
  try {
    if (!hasTelemetrySchema(db)) return emptyDetailPayload()
    const turnRows = readCostRows(db, "", [], "agent")
    const toolCallsByAgent = new Map(
      (
        db
          .query("SELECT COALESCE(agent, 'unknown') AS agent, COUNT(*) AS tool_calls FROM tool_calls GROUP BY COALESCE(agent, 'unknown')")
          .all() as { agent: string; tool_calls: number }[]
      ).map((row) => [row.agent, row.tool_calls]),
    )
    const byAgent = Array.from(
      turnRows
        .reduce((map, row) => {
          const agent = row.agent ?? "unknown"
          const current = map.get(agent) ?? { agent, cost_usd: 0, turns: 0, tool_calls: 0 }
          map.set(agent, { ...current, cost_usd: current.cost_usd + displayCost(row, rateOverrides), turns: current.turns + 1 })
          return map
        }, new Map<string, { agent: string; cost_usd: number; turns: number; tool_calls: number }>())
        .values(),
    )
      .map((row) => ({ ...row, tool_calls: toolCallsByAgent.get(row.agent) ?? 0 }))
      .sort((a, b) => b.cost_usd - a.cost_usd)
    const byModel = Array.from(
      turnRows
        .reduce((map, row) => {
          const model = row.model ?? "unknown"
          const current = map.get(model) ?? { model, cost_usd: 0, turns: 0 }
          map.set(model, { ...current, cost_usd: current.cost_usd + displayCost(row, rateOverrides), turns: current.turns + 1 })
          return map
        }, new Map<string, { model: string; cost_usd: number; turns: number }>())
        .values(),
    ).sort((a, b) => b.cost_usd - a.cost_usd)
    const recentSessions = (
      db
        .query(`
          SELECT
            id,
            COALESCE(agent, 'unknown') AS agent,
            status,
            last_activity_at
          FROM sessions
          ORDER BY last_activity_at DESC
          LIMIT 20
        `)
        .all() as Array<{ id: string; agent: string; status: string; last_activity_at: number }>
    ).map((session) => ({
      ...session,
      cost_usd: sumDisplayCost(readCostRows(db, "WHERE session_id = ?", [session.id]), rateOverrides),
    }))
    const sparkline = Array.from(
      readCostRows(db, "WHERE ts >= ?", [Date.now() - 24 * 60 * 60_000])
        .reduce((map, row) => {
          const ts = Math.trunc((row.ts ?? 0) / 3_600_000) * 3_600_000
          map.set(ts, (map.get(ts) ?? 0) + displayCost(row, rateOverrides))
          return map
        }, new Map<number, number>())
        .entries(),
    )
      .map(([ts, cost_usd]) => ({ ts, cost_usd }))
      .sort((a, b) => a.ts - b.ts)
    const recentTurns = (
      db
        .query(`
          SELECT
            ts,
            COALESCE(agent, 'unknown') AS agent,
            COALESCE(model, 'unknown') AS model,
            estimated_input_tokens AS tokens,
            estimated_input_tokens,
            cost_usd,
            status
          FROM turns
          ORDER BY ts DESC
          LIMIT 50
        `)
        .all() as Array<{
          ts: number
          agent: string
          model: string
          tokens: number
          estimated_input_tokens: number
          cost_usd: number | null
          status: string
        }>
    ).map((row) => ({
      ts: row.ts,
      agent: row.agent,
      model: row.model,
      tokens: row.tokens,
      cost_usd: displayCostOrNull(row, rateOverrides),
      status: row.status,
    }))
    return {
      by_agent: byAgent,
      by_model: byModel,
      by_tool: db
        .query(`
          SELECT
            tool,
            COUNT(*) AS calls,
            COALESCE(SUM(duration_ms), 0) AS total_duration_ms,
            COALESCE(AVG(duration_ms), 0) AS avg_duration_ms
          FROM tool_calls
          GROUP BY tool
          ORDER BY calls DESC
        `)
        .all() as { tool: string; calls: number; total_duration_ms: number; avg_duration_ms: number }[],
      recent_turns: recentTurns,
      recent_sessions: recentSessions,
      sparkline,
    }
  } catch {
    return emptyDetailPayload()
  } finally {
    db.close()
  }
}

export function readStoragePayload(matterDirs: string[], rateOverrides?: RateOverrides): StoragePayload {
  const models = new Set<string>()
  const payload = matterDirs.reduce(
    (acc, dir) => {
      const file = telemetryDbPath(dir)
      if (!existsSync(file)) return acc
      const sizeBytes = telemetryDbFiles(dir).reduce((sum, sibling) => sum + statSync(sibling).size, 0)
      const db = openReadOnly(dir)
      if (!db)
        return {
          ...acc,
          size_bytes: acc.size_bytes + sizeBytes,
        }
      try {
        if (!hasTelemetrySchema(db))
          return {
            ...acc,
            size_bytes: acc.size_bytes + sizeBytes,
          }
        const oldest = db.query("SELECT MIN(ts) AS ts FROM turns").get() as { ts: number | null }
        for (const model of readUnknownModels(db, rateOverrides)) models.add(model)
        return {
          size_bytes: acc.size_bytes + sizeBytes,
          oldest_ts: acc.oldest_ts === null ? oldest.ts : oldest.ts === null ? acc.oldest_ts : Math.min(acc.oldest_ts, oldest.ts),
        }
      } catch {
        return {
          ...acc,
          size_bytes: acc.size_bytes + sizeBytes,
        }
      } finally {
        db.close()
      }
    },
    { size_bytes: 0, oldest_ts: null as number | null },
  )
  return {
    ...payload,
    matter_count: matterDirs.length,
    unknown_models: Array.from(models).sort(),
    rate_card_as_of: null,
    rate_source: PRICING_SOURCE,
    rate_reference_url: PRICING_REFERENCE_URL,
  }
}

export function migrateLegacyContextBudget(matterDir: string, rateOverrides?: RateOverrides) {
  const legacyFile = path.join(matterDir, ".dochaus", "context-budget.jsonl")
  const dbFile = telemetryDbPath(matterDir)
  if (!existsSync(dbFile)) return { inserted: 0, skipped: 0, backfilled: 0 }
  const db = new Database(dbFile)
  try {
    if (!hasTelemetrySchema(db)) return { inserted: 0, skipped: 0, backfilled: 0 }
    const result = readMeta(db, "legacy_context_budget_migrated") || !existsSync(legacyFile)
      ? { inserted: 0, skipped: 0 }
      : importLegacyContextBudgetRows(db, legacyFile, rateOverrides)
    const backfilled = backfillKnownTurnCosts(db, rateOverrides)
    pruneKnownUnknownModels(db, rateOverrides)
    db.query("INSERT INTO meta (key, value) VALUES (?, ?) ON CONFLICT(key) DO UPDATE SET value = excluded.value").run(
      "pricing_source",
      PRICING_SOURCE,
    )
    db.query("DELETE FROM meta WHERE key = ?").run("pricing_as_of")
    return { ...result, backfilled }
  } finally {
    db.close()
  }
}

function importLegacyContextBudgetRows(db: Database, legacyFile: string, rateOverrides?: RateOverrides) {
  const firstTurn = db.query("SELECT MIN(ts) AS ts FROM turns").get() as { ts: number | null }
  const cutoff = firstTurn.ts ?? Number.POSITIVE_INFINITY
  const turnIndexes = new Map<string, number>()
  const result = readFileSync(legacyFile, "utf8")
    .split("\n")
    .map(parseLegacyContextBudgetRow)
    .reduce(
      (acc, row) => {
        if (!row) return acc
        if (row.ts >= cutoff) return { ...acc, skipped: acc.skipped + 1 }
        const sessionID = legacySessionID(row.agent, row.model)
        const turnIndex = turnIndexes.get(sessionID) ?? readSessionTurnCount(db, sessionID)
        turnIndexes.set(sessionID, turnIndex + 1)
        const rate = getRate(row.model, rateOverrides, row.estimated_input_tokens)
        db.query(`
            INSERT INTO sessions (id, parent_id, created_at, last_activity_at, status, model, agent)
            VALUES (?, NULL, ?, ?, 'finished', ?, ?)
            ON CONFLICT(id) DO UPDATE SET
              last_activity_at = MAX(last_activity_at, excluded.last_activity_at),
              model = excluded.model,
              agent = excluded.agent
          `).run(sessionID, row.ts, row.ts, row.model, row.agent)
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
            sessionID,
            turnIndex,
            row.ts,
            row.model,
            row.agent,
            row.context_chars,
            row.estimated_input_tokens,
            rate ? computeCost(row.estimated_input_tokens, rate) : null,
            row.status,
          )
        return { ...acc, inserted: acc.inserted + 1 }
      },
      { inserted: 0, skipped: 0 },
    )
  db.query("INSERT INTO meta (key, value) VALUES (?, ?) ON CONFLICT(key) DO UPDATE SET value = excluded.value").run(
    "legacy_context_budget_migrated",
    JSON.stringify({ ...result, migrated_at: new Date().toISOString() }),
  )
  return result
}

function readCurrentSession(db: Database, sessionID?: string) {
  const session = (
    sessionID
      ? db.query("SELECT id, model, agent FROM sessions WHERE id = ?").get(sessionID)
      : db.query("SELECT id, model, agent FROM sessions ORDER BY last_activity_at DESC LIMIT 1").get()
  ) as SessionRow
  if (!session) return null
  const turn = db
    .query(`
      SELECT
        ts,
        agent,
        model,
        estimated_input_tokens AS tokens,
        cost_usd,
        status
      FROM turns
      WHERE session_id = ?
      ORDER BY ts DESC
      LIMIT 1
    `)
    .get(session.id) as TurnRow
  const tokens = turn?.tokens ?? 0
  return {
    tokens,
    status: tokenStatus(tokens),
    agent: session.agent ?? turn?.agent ?? "unknown",
    model: session.model ?? turn?.model ?? "unknown",
  }
}

function tokenStatus(tokens: number): "ok" | "warn" | "fail" {
  if (tokens >= 180_000) return "fail"
  if (tokens >= 90_000) return "warn"
  return "ok"
}

function readMeta(db: Database, key: string) {
  return (db.query("SELECT value FROM meta WHERE key = ?").get(key) as MetaRow)?.value ?? null
}

function readUnknownModels(db: Database, rateOverrides?: RateOverrides) {
  const row = readMeta(db, "unknown_models")
  const used = db.query("SELECT DISTINCT model, estimated_input_tokens FROM turns WHERE model IS NOT NULL").all() as { model: string; estimated_input_tokens: number }[]
  const missing = Array.from(new Set(used
    .filter((turn) => !getRate(turn.model, rateOverrides, turn.estimated_input_tokens))
    .map((turn) => turn.model)))
  if (!row) return missing
  try {
    const parsed = JSON.parse(row) as unknown
    if (!Array.isArray(parsed)) return missing
    return Array.from(new Set([...parsed.filter((value): value is string => typeof value === "string" && !getRate(value, rateOverrides)), ...missing]))
  } catch {
    return missing
  }
}

function readCostRows(db: Database, clause = "", params: SQLQueryBindings[] = [], extra = "") {
  const fields = ["ts", "session_id", "model", "estimated_input_tokens", "cost_usd", extra].filter(Boolean).join(", ")
  return db.query(`SELECT ${fields} FROM turns ${clause}`).all(...params) as CostRow[]
}

function summarizeCostRows(rows: CostRow[], rateOverrides?: RateOverrides): SumRow {
  return {
    turns: rows.length,
    cost_usd: sumDisplayCost(rows, rateOverrides),
    since_ts: rows.reduce<number | null>((min, row) => (row.ts === undefined ? min : min === null ? row.ts : Math.min(min, row.ts)), null),
  }
}

function sumDisplayCost(rows: CostRow[], rateOverrides?: RateOverrides) {
  return rows.reduce((sum, row) => sum + displayCost(row, rateOverrides), 0)
}

function displayCost(row: CostRow, rateOverrides?: RateOverrides) {
  return displayCostOrNull(row, rateOverrides) ?? 0
}

function displayCostOrNull(row: CostRow, rateOverrides?: RateOverrides) {
  // Stored cost_usd values are historical input estimates, including estimates
  // from the retired rate card. Reprice the view without rewriting those rows.
  if (!row.model) return null
  const rate = getRate(row.model, rateOverrides, row.estimated_input_tokens)
  return rate ? computeCost(row.estimated_input_tokens, rate) : null
}

function readSessionTurnCount(db: Database, sessionID: string) {
  const row = db.query("SELECT COUNT(*) AS n FROM turns WHERE session_id = ?").get(sessionID) as { n: number } | null
  return row?.n ?? 0
}

function backfillKnownTurnCosts(db: Database, rateOverrides?: RateOverrides) {
  return (
    db
      .query("SELECT id, model, estimated_input_tokens FROM turns WHERE cost_usd IS NULL")
      .all() as { id: number; model: string | null; estimated_input_tokens: number }[]
  ).reduce((count, row) => {
    if (!row.model) return count
    const rate = getRate(row.model, rateOverrides, row.estimated_input_tokens)
    if (!rate) return count
    db.query("UPDATE turns SET cost_usd = ? WHERE id = ?").run(computeCost(row.estimated_input_tokens, rate), row.id)
    return count + 1
  }, 0)
}

function pruneKnownUnknownModels(db: Database, rateOverrides?: RateOverrides) {
  const row = readMeta(db, "unknown_models")
  if (!row) return
  try {
    const parsed = JSON.parse(row) as unknown
    if (!Array.isArray(parsed)) return
    const unknown = parsed.filter((value): value is string => typeof value === "string" && !getRate(value, rateOverrides))
    if (unknown.length === 0) {
      db.query("DELETE FROM meta WHERE key = ?").run("unknown_models")
      return
    }
    db.query("INSERT INTO meta (key, value) VALUES (?, ?) ON CONFLICT(key) DO UPDATE SET value = excluded.value").run(
      "unknown_models",
      JSON.stringify(unknown),
    )
  } catch {
    return
  }
}

function parseLegacyContextBudgetRow(line: string): LegacyContextBudgetRow | undefined {
  if (!line.trim()) return undefined
  try {
    const row = JSON.parse(line) as unknown
    if (!isRecord(row)) return undefined
    const ts = Date.parse(stringValue(row.created_at))
    const model = legacyModel(row.model)
    const estimatedInputTokens = numberValue(row.estimated_tokens)
    if (!Number.isFinite(ts) || !model || !Number.isFinite(estimatedInputTokens)) return undefined
    const contextChars = numberValue(row.chars)
    return {
      ts,
      model,
      agent: legacyAgent(row.largest_messages),
      context_chars: Number.isFinite(contextChars) ? contextChars : estimatedInputTokens * 4,
      estimated_input_tokens: estimatedInputTokens,
      status: legacyStatus(row.status, estimatedInputTokens),
    }
  } catch {
    return undefined
  }
}

function legacyModel(value: unknown) {
  if (typeof value === "string") return value
  if (!isRecord(value)) return undefined
  const providerID = stringValue(value.providerID)
  const modelID = stringValue(value.modelID) || stringValue(value.id)
  if (!providerID || !modelID) return undefined
  return `${providerID}/${modelID}`
}

function legacyAgent(value: unknown) {
  if (!Array.isArray(value)) return "unknown"
  const counts = value
    .map((item) => (isRecord(item) ? stringValue(item.agent) : ""))
    .filter(Boolean)
    .reduce((map, agent) => map.set(agent, (map.get(agent) ?? 0) + 1), new Map<string, number>())
  return Array.from(counts.entries()).sort((a, b) => b[1] - a[1])[0]?.[0] ?? "unknown"
}

function legacyStatus(value: unknown, estimatedInputTokens: number) {
  if (value === "ok" || value === "warn" || value === "fail") return value
  return tokenStatus(estimatedInputTokens)
}

function legacySessionID(agent: string, model: string) {
  return `legacy:${agent}:${model}`.replace(/[^A-Za-z0-9:._-]/g, "_")
}

function stringValue(value: unknown) {
  return typeof value === "string" ? value : ""
}

function numberValue(value: unknown) {
  return typeof value === "number" && Number.isFinite(value) ? value : Number.NaN
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null
}

function openReadOnly(matterDir: string) {
  const file = telemetryDbPath(matterDir)
  if (!existsSync(file)) return null
  try {
    return new Database(file, { readonly: true })
  } catch {
    return null
  }
}

function hasTelemetrySchema(db: Database) {
  const tables = db
    .query("SELECT name FROM sqlite_master WHERE type = 'table' AND name IN ('sessions', 'turns', 'tool_calls', 'meta')")
    .all() as { name: string }[]
  return tables.length === 4
}

function telemetryDbPath(matterDir: string) {
  return path.join(matterDir, ".dochaus", "telemetry.db")
}

function telemetryDbFiles(matterDir: string) {
  const file = telemetryDbPath(matterDir)
  return [file, `${file}-wal`, `${file}-shm`].filter((sibling) => existsSync(sibling))
}

function emptyLeanPayload(): LeanPayload {
  return {
    current_activity: null,
    today: { cost_usd: 0, turns: 0 },
    matter_total: { cost_usd: 0, turns: 0, sessions: 0, since_ts: 0 },
    current_session: null,
    rate_card_as_of: null,
    unknown_models: [],
  }
}

function emptyDetailPayload(): DetailPayload {
  return { by_agent: [], by_model: [], by_tool: [], recent_turns: [], recent_sessions: [], sparkline: [] }
}
