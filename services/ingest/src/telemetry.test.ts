import { Database } from "bun:sqlite"
import { afterEach, expect, test } from "bun:test"
import { mkdtempSync, mkdirSync, rmSync, statSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { clearUsageCache, readDetailPayload, readLeanPayload, readStoragePayload } from "./telemetry"

const tempDirs: string[] = []
const originalFetch = globalThis.fetch
const signal = () => AbortSignal.timeout(5_000)

afterEach(() => {
  globalThis.fetch = originalFetch
  clearUsageCache()
  for (const dir of tempDirs.splice(0)) rmSync(dir, { recursive: true, force: true })
})

type FakeMessage = { id: string; time: number; cost: number; agent?: string; model?: string; tokens?: { input: number; output: number } }
type FakeSession = { id: string; agent?: string; parentID?: string; title?: string; metadata?: Record<string, unknown>; cost: number; tokens?: { input: number; output: number }; updated: number; messages: FakeMessage[] }

// A synthetic native engine: v2 session pages and transcripts, with message
// reads counted so cache reuse is observable.
function fakeEngine(dir: string, sessions: FakeSession[], options: { down?: boolean } = {}) {
  const reads: string[] = []
  const native = (session: FakeSession) => ({
    id: session.id,
    location: { directory: dir },
    title: session.title ?? session.id,
    ...(session.agent ? { agent: session.agent } : {}),
    ...(session.parentID ? { parentID: session.parentID } : {}),
    ...(session.metadata ? { metadata: session.metadata } : {}),
    cost: session.cost,
    tokens: { input: 0, output: 0, reasoning: 0, cache: { read: 0, write: 0 }, ...session.tokens },
    time: { created: session.updated - 1_000, updated: session.updated },
  })
  globalThis.fetch = Object.assign(async (input: Parameters<typeof fetch>[0]) => {
    const url = new URL(input instanceof Request ? input.url : String(input))
    if (options.down) return new Response("down", { status: 503 })
    if (url.pathname === "/api/session") return Response.json({ data: sessions.map(native), cursor: {} })
    const match = /^\/api\/session\/([^/]+)(\/message)?$/.exec(url.pathname)
    const session = sessions.find((item) => item.id === decodeURIComponent(match?.[1] ?? ""))
    if (!match || !session) return new Response("missing", { status: 404 })
    if (!match[2]) return Response.json({ data: native(session) })
    reads.push(session.id)
    return Response.json({
      data: session.messages.map((message) => ({
        id: message.id,
        type: "assistant",
        time: { created: message.time },
        cost: message.cost,
        ...(message.agent ? { agent: message.agent } : {}),
        ...(message.model ? { model: { providerID: message.model.split("/")[0], id: message.model.split("/")[1] } } : {}),
        tokens: { input: 0, output: 0, reasoning: 0, cache: { read: 0, write: 0 }, ...message.tokens },
        content: [],
      })).toReversed(),
      cursor: {},
    })
  }, { preconnect: originalFetch.preconnect }) as typeof fetch
  return reads
}

test("totals are the provider-reported cost of every model step, upkeep and generation", async () => {
  const dir = tempMatter()
  const now = Date.now()
  const db = seedDb(dir)
  insertGeneration(db, { ts: now - 2_000, agent: "prepare-changes", model: "synthetic/fast", cost: 0.05 })
  db.close()
  fakeEngine(dir, [{
    id: "s1", agent: "draft", cost: 0.5, updated: now - 1_000,
    messages: [
      { id: "m1", time: now - 3_000, cost: 0.3, agent: "draft", model: "synthetic/strong", tokens: { input: 1_000, output: 200 } },
      { id: "m2", time: now - 2_500, cost: 0.1, agent: "draft", model: "synthetic/strong", tokens: { input: 500, output: 100 } },
    ],
  }])

  const lean = await readLeanPayload(dir, { signal: signal(), session: "s1" })
  expect(lean.matter_total.cost_usd).toBeCloseTo(0.55)
  expect(lean.matter_total).toMatchObject({ turns: 3, sessions: 1, since_ts: now - 3_000 })
  expect(lean.current_activity?.cost_usd).toBeCloseTo(0.55)
  expect(lean.today.cost_usd).toBeCloseTo(0.55)
  expect(lean).toMatchObject({ unpriced_turns: 0, partial: false, current_session: null })

  const detail = await readDetailPayload(dir, signal())
  expect(detail.by_model).toEqual([
    { model: "synthetic/strong", cost_usd: expect.closeTo(0.4), turns: 2, unpriced_turns: 0 },
    { model: null, label: "Conversation titles and compaction", cost_usd: expect.closeTo(0.1), turns: 0, unpriced_turns: 0 },
    { model: "synthetic/fast", cost_usd: 0.05, turns: 1, unpriced_turns: 0 },
  ])
  expect(detail.by_model.reduce((sum, row) => sum + row.cost_usd, 0)).toBeCloseTo(lean.matter_total.cost_usd)
  expect(detail.by_agent.map(({ agent, label, group, turns }) => ({ agent, label, group, turns }))).toEqual([
    { agent: "draft", label: "Draft", group: "work", turns: 2 },
    { agent: "conversation-upkeep", label: "Conversation titles and compaction", group: "work", turns: 0 },
    { agent: "prepare-changes", label: "Prepare changes", group: "work", turns: 1 },
  ])
  expect(detail.recent_turns.map((row) => row.cost_usd)).toEqual([0.05, 0.1, 0.3])
  expect(detail.recent_turns[0]).toMatchObject({ label: "Prepare changes", model: "synthetic/fast", tokens: 0 })
  expect(detail.recent_turns[1]).toMatchObject({ label: "Draft", tokens: 600 })
  expect(detail.sparkline.reduce((sum, point) => sum + point.cost_usd, 0)).toBeCloseTo(0.55)
})

test("background agents and Word suggestions are split from the lawyer's own work", async () => {
  const dir = tempMatter()
  const now = Date.now()
  const db = seedDb(dir)
  insertToolCall(db, { sessionId: "words", ts: now, tool: "read-document", agent: "qa", durationMs: 10 })
  insertToolCall(db, { sessionId: "chat", ts: now, tool: "read-document", agent: "qa", durationMs: 30 })
  insertToolCall(db, { sessionId: "chat", ts: now, tool: "record-key-date", agent: "qa", durationMs: 5 })
  db.close()
  fakeEngine(dir, [
    { id: "chat", agent: "qa", cost: 0.4, updated: now - 1_000, messages: [{ id: "m1", time: now - 1_000, cost: 0.4, agent: "qa", model: "synthetic/strong" }] },
    { id: "words", agent: "qa", metadata: { kind: "suggestions" }, cost: 0.2, updated: now - 2_000, messages: [{ id: "m2", time: now - 2_000, cost: 0.2, agent: "qa", model: "synthetic/fast" }] },
    { id: "story", agent: "narrator", cost: 0.1, updated: now - 3_000, messages: [{ id: "m3", time: now - 3_000, cost: 0.1, agent: "narrator", model: "synthetic/fast" }] },
  ])

  const detail = await readDetailPayload(dir, signal())
  expect(detail.by_group).toEqual([
    { group: "background", label: "Background agents", cost_usd: expect.closeTo(0.3), turns: 2 },
    { group: "work", label: "Your work", cost_usd: 0.4, turns: 1 },
  ])
  expect(detail.by_agent).toContainEqual({ agent: "word-suggestions", label: "Suggest ideas in Word", group: "background", cost_usd: 0.2, turns: 1, tool_calls: 1 })
  expect(detail.by_agent).toContainEqual({ agent: "qa", label: "Questions", group: "work", cost_usd: 0.4, turns: 1, tool_calls: 2 })
  expect(detail.by_tool).toEqual([
    { tool: "read-document", label: "Read document", calls: 2, total_duration_ms: 40, avg_duration_ms: 20 },
    { tool: "record-key-date", label: "Record key date", calls: 1, total_duration_ms: 5, avg_duration_ms: 5 },
  ])
  expect(detail.recent_sessions.map(({ id, label, group }) => ({ id, label, group }))).toEqual([
    { id: "chat", label: "Questions", group: "work" },
    { id: "words", label: "Suggest ideas in Word", group: "background" },
    { id: "story", label: "Narrate drafting progress", group: "background" },
  ])
})

test("a conversation's cost includes the sub-agents it started", async () => {
  const dir = tempMatter()
  const now = Date.now()
  fakeEngine(dir, [
    { id: "root", agent: "draft", cost: 0.2, updated: now - 1_000, messages: [{ id: "m1", time: now - 2_000, cost: 0.2, agent: "draft" }] },
    { id: "child", parentID: "root", agent: "lean-advise", cost: 0.3, updated: now - 1_500, messages: [{ id: "m2", time: now - 1_500, cost: 0.3, agent: "lean-advise" }] },
  ])

  const detail = await readDetailPayload(dir, signal())
  expect(detail.recent_sessions).toEqual([{ id: "root", title: "root", agent: "draft", label: "Draft", group: "work", last_activity_at: now - 1_000, cost_usd: 0.5 }])
})

test("polling re-reads only conversations that changed", async () => {
  const dir = tempMatter()
  const now = Date.now()
  const sessions: FakeSession[] = [
    { id: "a", agent: "draft", cost: 0.2, updated: now - 2_000, messages: [{ id: "m1", time: now - 2_000, cost: 0.2 }] },
    { id: "b", agent: "qa", cost: 0.1, updated: now - 1_000, messages: [{ id: "m2", time: now - 1_000, cost: 0.1 }] },
  ]
  const reads = fakeEngine(dir, sessions)

  expect((await readLeanPayload(dir, { signal: signal() })).matter_total.cost_usd).toBeCloseTo(0.3)
  expect(reads.toSorted()).toEqual(["a", "b"])
  await readLeanPayload(dir, { signal: signal() })
  await readDetailPayload(dir, signal())
  expect(reads).toHaveLength(2)

  sessions[1] = { ...sessions[1]!, cost: 0.15, updated: now, messages: [...sessions[1]!.messages, { id: "m3", time: now, cost: 0.05 }] }
  expect((await readLeanPayload(dir, { signal: signal() })).matter_total.cost_usd).toBeCloseTo(0.35)
  expect(reads.slice(2)).toEqual(["b"])
})

test("usage with tokens but no reported cost is counted as unpriced", async () => {
  const dir = tempMatter()
  const now = Date.now()
  const db = seedDb(dir)
  insertGeneration(db, { ts: now - 1_000, agent: "comment-replies", model: "synthetic/local", cost: null, input: 400 })
  db.close()
  fakeEngine(dir, [
    { id: "local", agent: "qa", cost: 0, tokens: { input: 2_000, output: 300 }, updated: now - 2_000, messages: [] },
    { id: "empty", agent: "qa", cost: 0, updated: now - 3_000, messages: [] },
  ])

  const lean = await readLeanPayload(dir, { signal: signal() })
  expect(lean.matter_total).toMatchObject({ cost_usd: 0, turns: 2 })
  expect(lean.unpriced_turns).toBe(2)
  const detail = await readDetailPayload(dir, signal())
  expect(detail.by_model).toContainEqual({ model: "synthetic/local", cost_usd: 0, turns: 1, unpriced_turns: 1 })
  expect(detail.recent_turns.every((row) => row.cost_usd === null)).toBe(true)
})

test("an unreachable engine keeps generation usage and marks the totals partial", async () => {
  const dir = tempMatter()
  const db = seedDb(dir)
  insertGeneration(db, { ts: Date.now(), agent: "prepare-changes", model: "synthetic/fast", cost: 0.02 })
  db.close()
  fakeEngine(dir, [], { down: true })

  const lean = await readLeanPayload(dir, { signal: signal() })
  expect(lean).toMatchObject({ partial: true, matter_total: { cost_usd: 0.02, turns: 1, sessions: 0 } })
})

test("the current conversation reports its context size with a plain assistant name", async () => {
  const dir = tempMatter()
  const now = Date.now()
  const db = seedDb(dir)
  insertSession(db, { id: "words", lastActivityAt: now, model: "synthetic/fast", agent: "qa" })
  insertTurn(db, { sessionId: "words", ts: now, tokens: 120_000 })
  db.close()
  fakeEngine(dir, [{ id: "words", agent: "qa", metadata: { kind: "suggestions" }, cost: 0, updated: now, messages: [] }])

  expect((await readLeanPayload(dir, { signal: signal(), session: "words" })).current_session).toEqual({
    tokens: 120_000, status: "warn", agent: "word-suggestions", label: "Suggest ideas in Word", model: "synthetic/fast",
  })
})

test("the latest conversation passes over background agents, sub-agents and runtime jobs", async () => {
  const dir = tempMatter()
  const now = Date.now()
  const db = seedDb(dir)
  insertSession(db, { id: "chat", lastActivityAt: now - 5_000, model: "synthetic/strong", agent: "qa" })
  insertTurn(db, { sessionId: "chat", ts: now - 5_000, tokens: 4_000 })
  insertSession(db, { id: "child", parentId: "chat", lastActivityAt: now - 4_000, model: "synthetic/fast", agent: "lean-advise" })
  insertSession(db, { id: "job", lastActivityAt: now - 3_000, model: "synthetic/fast", agent: "dashboard" })
  insertSession(db, { id: "sub", lastActivityAt: now - 2_000, model: "synthetic/fast", agent: "explore" })
  insertSession(db, { id: "compare", lastActivityAt: now - 1_000, model: "synthetic/fast", agent: "review-compare" })
  insertTurn(db, { sessionId: "compare", ts: now - 1_000, tokens: 296 })
  insertSession(db, { id: "cell", lastActivityAt: now - 500, model: "synthetic/fast", agent: "extract" })
  insertTurn(db, { sessionId: "cell", ts: now - 500, tokens: 900 })
  db.close()
  fakeEngine(dir, [
    { id: "chat", agent: "qa", cost: 0, updated: now - 5_000, messages: [] },
    { id: "child", agent: "lean-advise", parentID: "chat", cost: 0, updated: now - 4_000, messages: [] },
    { id: "job", agent: "dashboard", metadata: { dochaus_role: "dashboard" }, cost: 0, updated: now - 3_000, messages: [] },
    { id: "sub", agent: "explore", parentID: "chat", cost: 0, updated: now - 2_000, messages: [] },
    { id: "compare", agent: "review-compare", metadata: { dochaus_role: "review-compare" }, cost: 0, updated: now - 1_000, messages: [] },
    { id: "cell", agent: "extract", title: "doc.haus grid: Lease.docx", cost: 0, updated: now - 500, messages: [] },
  ])

  expect((await readLeanPayload(dir, { signal: signal() })).current_session).toEqual({
    tokens: 4_000, status: "ok", agent: "qa", label: "Questions", model: "synthetic/strong",
  })
  expect((await readLeanPayload(dir, { signal: signal(), session: "compare" })).current_session).toMatchObject({
    tokens: 296, agent: "review-compare", label: "Compare review answers",
  })
})

test("storage sums database files and finds the oldest clearable record", () => {
  const first = tempMatter()
  const second = tempMatter()
  const db = seedDb(first)
  insertToolCall(db, { sessionId: "s1", ts: 5_000, tool: "read-document", agent: "qa", durationMs: null })
  insertGeneration(db, { ts: 2_000, agent: "prepare-changes", cost: 0.01 })
  insertTurn(db, { sessionId: "s1", ts: 9_000 })
  db.close()

  const storage = readStoragePayload([first, second])
  expect(storage.size_bytes).toBeGreaterThan(0)
  expect(storage).toMatchObject({ matter_count: 2, oldest_ts: 2_000 })
})

test("storage includes WAL sibling file sizes", () => {
  const dir = tempMatter()
  seedDb(dir).close()
  const base = path.join(dir, ".dochaus", "telemetry.db")
  writeFileSync(`${base}-wal`, "wal")
  writeFileSync(`${base}-shm`, "shm-data")
  const expectedSize = statSync(base).size + statSync(`${base}-wal`).size + statSync(`${base}-shm`).size

  expect(readStoragePayload([dir]).size_bytes).toBe(expectedSize)
})

test("invalid or garbage telemetry databases leave native usage readable", async () => {
  const empty = tempMatter()
  mkdirSync(path.join(empty, ".dochaus"), { recursive: true })
  new Database(path.join(empty, ".dochaus", "telemetry.db")).close()
  const garbage = tempMatter()
  mkdirSync(path.join(garbage, ".dochaus"), { recursive: true })
  writeFileSync(path.join(garbage, ".dochaus", "telemetry.db"), "not sqlite")

  for (const dir of [empty, garbage]) {
    fakeEngine(dir, [{ id: "s1", agent: "qa", cost: 0.1, updated: Date.now(), messages: [{ id: "m1", time: Date.now(), cost: 0.1 }] }])
    expect(await readLeanPayload(dir, { signal: signal(), session: "s1" })).toMatchObject({ matter_total: { cost_usd: 0.1, turns: 1 }, current_session: null })
    expect((await readDetailPayload(dir, signal())).by_tool).toEqual([])
    expect(readStoragePayload([dir])).toMatchObject({ matter_count: 1, oldest_ts: null })
  }
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
    CREATE TABLE generations (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      ts INTEGER NOT NULL,
      agent TEXT NOT NULL,
      model TEXT,
      input_tokens INTEGER NOT NULL DEFAULT 0,
      output_tokens INTEGER NOT NULL DEFAULT 0,
      reasoning_tokens INTEGER NOT NULL DEFAULT 0,
      cache_read_tokens INTEGER NOT NULL DEFAULT 0,
      cache_write_tokens INTEGER NOT NULL DEFAULT 0,
      cost_usd REAL,
      session_id TEXT
    );
    CREATE TABLE meta (
      key TEXT PRIMARY KEY,
      value TEXT NOT NULL
    );
  `)
  return db
}

function insertSession(db: Database, input: { id: string; parentId?: string; lastActivityAt: number; model?: string; agent?: string }) {
  db.query("INSERT INTO sessions (id, parent_id, created_at, last_activity_at, status, model, agent) VALUES (?, ?, ?, ?, 'active', ?, ?)")
    .run(input.id, input.parentId ?? null, input.lastActivityAt, input.lastActivityAt, input.model ?? null, input.agent ?? null)
}

function insertTurn(db: Database, input: { sessionId: string; ts: number; tokens?: number }) {
  db.query("INSERT INTO turns (session_id, turn_index, ts, model, agent, context_chars, estimated_input_tokens, cost_usd, status) VALUES (?, 0, ?, 'synthetic/fast', 'qa', 0, ?, NULL, 'ok')")
    .run(input.sessionId, input.ts, input.tokens ?? 1_000)
}

function insertToolCall(db: Database, input: { sessionId: string; ts: number; tool: string; agent: string; durationMs: number | null }) {
  db.query("INSERT INTO tool_calls (session_id, turn_index, ts, tool, agent, duration_ms, output_chars, phase) VALUES (?, 0, ?, ?, ?, ?, 0, 'ok')")
    .run(input.sessionId, input.ts, input.tool, input.agent, input.durationMs)
}

function insertGeneration(db: Database, input: { ts: number; agent: string; model?: string; cost: number | null; input?: number }) {
  db.query("INSERT INTO generations (ts, agent, model, input_tokens, cost_usd) VALUES (?, ?, ?, ?, ?)")
    .run(input.ts, input.agent, input.model ?? null, input.input ?? 0, input.cost)
}
