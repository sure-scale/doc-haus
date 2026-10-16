import { openTelemetryDb, recordToolCall, recordTurn, summarizeMessages, touchSession } from "../../lib/telemetry"
import { sessionAgent } from "./session"
import { diagnosticErrorMetadata } from "../../lib/diagnostic-metadata"
import { normalizePreferences } from "../../lib/drafting-preferences"
import { readFirmPreferences } from "../../lib/preferences"

const SESSION_TELEMETRY = new Map<string, { directory: string; agent: string; model: string }>()
const TURN_COUNTERS = new Map<string, number>()
const TOOL_START_TS = new Map<string, number>()

export function startToolCall(callID: string | undefined) {
  if (callID) TOOL_START_TS.set(callID, Date.now())
}

export function clearToolCall(callID: string | undefined) {
  if (callID) TOOL_START_TS.delete(callID)
}

export function recordContextBudget(
  directory: string,
  messages: {
    info: { type?: string; id?: string; sessionID?: string; agent?: string; model?: { providerID?: string; modelID?: string; id?: string } }
    parts: unknown[]
  }[],
) {
  let db: import("bun:sqlite").Database | undefined
  try {
    const summary = summarizeMessages(messages)
    const session = telemetrySessionContext(directory, messageSessionID(messages))
    if (!session) return
    db = openTelemetryDb(directory)
    recordTurn(db, {
      sessionID: session.sessionID,
      turnIndex: nextTurnIndex(db, session.sessionID),
      model: session.model,
      agent: session.agent,
      contextChars: summary.chars,
      estimatedInputTokens: summary.estimatedTokens,
      rateOverrides: normalizePreferences(readFirmPreferences()).telemetry.rate_overrides,
    })
  } catch (e) {
    console.error("telemetry recordTurn failed", diagnosticErrorMetadata(e))
  } finally {
    closeTelemetryDb(db)
  }
}

export function recordTelemetrySession(sessionID: string, telemetry: { directory: string; agent: string; model: string }): void {
  if (!SESSION_TELEMETRY.has(sessionID) && SESSION_TELEMETRY.size >= 512) {
    const oldest = SESSION_TELEMETRY.keys().next().value
    if (oldest !== undefined) SESSION_TELEMETRY.delete(oldest)
  }
  SESSION_TELEMETRY.set(sessionID, telemetry)
}

function telemetrySessionContext(
  directory: string,
  sessionID: string | undefined,
): { sessionID: string; agent: string; model: string } | undefined {
  if (!sessionID) return
  const telemetry = SESSION_TELEMETRY.get(sessionID)
  if (!telemetry || telemetry.directory !== directory) return
  return { sessionID, agent: telemetry.agent, model: telemetry.model }
}

function messageSessionID(messages: { info: { sessionID?: string } }[]): string | undefined {
  const ids = new Set(messages.map((message) => message.info.sessionID).filter((id): id is string => !!id))
  if (ids.size !== 1) return
  return ids.values().next().value
}

function nextTurnIndex(db: import("bun:sqlite").Database, sessionID: string): number {
  const current = TURN_COUNTERS.get(sessionID)
  if (current !== undefined) {
    const next = current + 1
    TURN_COUNTERS.set(sessionID, next)
    return next
  }
  const row = db.query("SELECT COUNT(*) AS n FROM turns WHERE session_id = ?").get(sessionID) as { n: number } | null
  const seed = row?.n ?? 0
  TURN_COUNTERS.set(sessionID, seed)
  return seed
}

export function touchTelemetrySession(
  directory: string,
  sessionID: string,
  agent: string,
  model: string,
): void {
  let db: import("bun:sqlite").Database | undefined
  try {
    db = openTelemetryDb(directory)
    touchSession(db, sessionID, { agent, model, status: "active" })
  } catch (e) {
    console.error("telemetry touchSession failed", diagnosticErrorMetadata(e))
  } finally {
    closeTelemetryDb(db)
  }
}

export function recordToolCallTelemetry(
  directory: string,
  input: { tool: string; sessionID?: string; callID?: string; args?: Record<string, unknown> },
  output: { output?: unknown; metadata?: Record<string, unknown> },
): void {
  if (!input.sessionID) {
    if (input.callID) TOOL_START_TS.delete(input.callID)
    return
  }
  let db: import("bun:sqlite").Database | undefined
  try {
    const start = input.callID ? TOOL_START_TS.get(input.callID) : undefined
    if (input.callID) TOOL_START_TS.delete(input.callID)
    const durationMs = start ? Date.now() - start : null
    const outputStr = typeof output.output === "string" ? output.output : ""
    const agent = sessionAgent(input.sessionID) ?? "unknown"
    db = openTelemetryDb(directory)
    recordToolCall(db, {
      sessionID: input.sessionID,
      turnIndex: null,
      tool: input.tool,
      agent,
      durationMs,
      outputChars: outputStr.length,
      phase: output.metadata?.error ? "error" : "ok",
    })
  } catch (e) {
    console.error("telemetry recordToolCall failed", diagnosticErrorMetadata(e))
  } finally {
    closeTelemetryDb(db)
  }
}

function closeTelemetryDb(db: import("bun:sqlite").Database | undefined): void {
  if (!db) return
  try {
    db.close()
  } catch (e) {
    console.error("telemetry db close failed", diagnosticErrorMetadata(e))
  }
}
