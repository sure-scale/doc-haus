import { afterEach, describe, expect, test } from "bun:test"
import { mkdtempSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import {
  ChatContinuationConflictError,
  dispatchChatContinuation,
  malformedToolContinuationInput,
  readSessionExecutionState,
  type ChatContinuationInput,
} from "./chat-continuation"
import { openDb } from "./db"

const cleanups: Array<() => void> = []
afterEach(() => cleanups.splice(0).forEach(cleanup => cleanup()))

function matter() {
  const directory = mkdtempSync(path.join(tmpdir(), "chat-continuation-"))
  cleanups.push(() => rmSync(directory, { recursive: true, force: true }))
  return directory
}

function input(overrides: Partial<ChatContinuationInput> = {}): ChatContinuationInput {
  return {
    ...malformedToolContinuationInput("ses_owner", "msg_failed"),
    agent: "qa",
    model: { providerID: "provider", modelID: "model" },
    ...overrides,
  }
}

// A native v2 engine stand-in: the owner session, its history (newest first, as
// v2 pages it), its durable inbox and the prompt admission route.
function engine() {
  const directory = matter()
  const continuation = input()
  const session: Record<string, unknown> = { id: "ses_owner", agent: "qa", model: { providerID: "provider", id: "model" }, location: { directory } }
  const history: Array<Record<string, unknown>> = [
    { id: "msg_user", type: "user", text: "Continue the work.", time: { created: 1 } },
    { id: "msg_failed", type: "assistant", agent: "qa", model: session.model, time: { created: 2, completed: 3 }, content: [{ type: "reasoning", text: '<|tool_call|>call:task{"description":"continue"}' }] },
  ]
  const state = {
    posts: 0, available: true, queued: false, changed: false, inboxAvailable: true,
    active: {} as Record<string, unknown>, body: undefined as Record<string, unknown> | undefined,
  }
  const admission = () => ({ id: continuation.messageID, sessionID: continuation.sessionID, type: "user", delivery: "queue", payload: { text: continuation.text + (state.changed ? " changed" : "") } })
  const server = Bun.serve({ port: 0, hostname: "127.0.0.1", async fetch(request) {
    const url = new URL(request.url)
    expect(request.headers.get("x-opencode-directory")).toBe(directory)
    if (url.pathname === "/api/session/ses_owner") return Response.json({ data: session })
    if (url.pathname === "/api/session/ses_child") return Response.json({ data: { ...session, id: "ses_child", parentID: "ses_owner" } })
    if (url.pathname === "/api/session/active") return Response.json({ data: state.active })
    if (url.pathname === "/api/permission/request") return Response.json({ data: [] })
    if (url.pathname === "/api/session/ses_owner/inbox") return state.inboxAvailable ? Response.json({ data: state.queued ? [admission()] : [] }) : new Response(null, { status: 503 })
    if (url.pathname === "/api/session/ses_owner/message") return Response.json({ data: history.toReversed(), cursor: {} })
    if (url.pathname === "/api/session/ses_owner/prompt") {
      state.posts++
      state.body = await request.json() as Record<string, unknown>
      if (!state.available) return new Response(null, { status: 503 })
      state.queued = true
      return Response.json({ data: admission() })
    }
    return new Response(null, { status: 404 })
  } })
  cleanups.push(() => server.stop(true))
  return { directory, input: continuation, session, history, state, options: { engineURL: server.url.origin } }
}

function claims(directory: string) {
  const db = openDb(directory)
  const rows = db.query("SELECT attempts, status FROM continuation_dispatches").all()
  db.close()
  return rows
}

describe("chat continuation admission", () => {
  test("concurrent continuation calls share one native durable admission", async () => {
    const f = engine()
    const results = await Promise.all([dispatchChatContinuation(f.directory, f.input, f.options), dispatchChatContinuation(f.directory, f.input, f.options)])
    expect(f.state.posts).toBe(1)
    expect(f.state.body).toMatchObject({ id: f.input.messageID, text: f.input.text, delivery: "queue" })
    expect(results[0]).toMatchObject({ accepted: true, persisted: true, uncertain: false })
    expect(results[1]).toMatchObject({ accepted: true, existing: true, persisted: true, pending: true })
  })

  test("reconciles admission after a lost response without sending another prompt", async () => {
    const f = engine()
    f.state.available = false
    expect(await dispatchChatContinuation(f.directory, f.input, f.options)).toMatchObject({ uncertain: true, persisted: false })
    f.state.queued = true
    expect(await dispatchChatContinuation(f.directory, f.input, f.options)).toMatchObject({ accepted: true, persisted: true, pending: true })
    expect(f.state.posts).toBe(1)
    expect(claims(f.directory)).toEqual([{ attempts: 1, status: "observed" }])
  })

  test("a conflicting durable inbox payload cannot acknowledge the claim", async () => {
    const f = engine()
    f.state.available = false
    await dispatchChatContinuation(f.directory, f.input, f.options)
    f.state.queued = true
    f.state.changed = true
    await expect(dispatchChatContinuation(f.directory, f.input, f.options)).rejects.toThrow("admission conflicts")
    expect(f.state.posts).toBe(1)
  })

  test("an exact retry keeps its ID and stays bounded when admission is unavailable", async () => {
    const f = engine()
    f.state.available = false
    for (let attempt = 0; attempt < 3; attempt++) await dispatchChatContinuation(f.directory, f.input, f.options)
    expect(f.state.posts).toBe(2)
    expect(f.state.body?.id).toBe(f.input.messageID)
    expect(await dispatchChatContinuation(f.directory, f.input, f.options)).toMatchObject({ attempted: false, exhausted: true, uncertain: true })
  })

  test("rejects a different payload for an already claimed message ID", async () => {
    const f = engine()
    await dispatchChatContinuation(f.directory, f.input, f.options)
    await expect(dispatchChatContinuation(f.directory, { ...f.input, text: `${f.input.text}\nDifferent payload.` }, f.options))
      .rejects.toThrow("already claimed by another payload")
  })
})

describe("chat continuation currentness", () => {
  test("rejects client-minted identities and a turn that already answered", async () => {
    const f = engine()
    await expect(dispatchChatContinuation(f.directory, input({ messageID: "msg_recovery_client_minted" }), f.options))
      .rejects.toBeInstanceOf(ChatContinuationConflictError)
    f.history[1] = { ...f.history[1], content: [{ type: "text", text: "A complete answer." }] }
    await expect(dispatchChatContinuation(f.directory, f.input, f.options))
      .rejects.toThrow("no longer the current recoverable failure")
    expect(f.state.posts).toBe(0)
    expect(claims(f.directory)).toEqual([])
  })

  test("enforces the malformed-output recovery bound from session history", async () => {
    const f = engine()
    const failed = f.history[1]!
    f.history.splice(1, 1,
      { id: "msg_recovery_one", type: "user", text: f.input.text, time: { created: 2 } },
      { ...failed, id: "msg_failed_one", content: [{ type: "reasoning", text: "<|tool_call|>" }] },
      { id: "msg_recovery_two", type: "user", text: f.input.text, time: { created: 3 } },
      failed,
    )
    await expect(dispatchChatContinuation(f.directory, f.input, f.options))
      .rejects.toThrow(/exhausted its bounded malformed-output recovery attempts/)
    expect(f.state.posts).toBe(0)
  })

  test("refuses child and runtime-owned sessions before claiming", async () => {
    for (const owned of [{ parentID: "ses_parent" }, { metadata: { dochaus_role: "lean-drafting" } }]) {
      const f = engine()
      Object.assign(f.session, owned)
      await expect(dispatchChatContinuation(f.directory, f.input, { ...f.options, requireIdleOwner: true }))
        .rejects.toThrow("applies only to top-level chat sessions")
      expect(f.state.posts).toBe(0)
      expect(claims(f.directory)).toEqual([])
    }
  })
})

describe("idle-only recovery", () => {
  test("queued inputs, an active subagent task and unreadable inboxes block a new claim", async () => {
    const f = engine()
    const idleOnly = { ...f.options, requireIdleOwner: true }
    f.state.queued = true
    expect(await readSessionExecutionState(f.directory, f.input.sessionID, f.options)).toBe("busy")
    expect(await dispatchChatContinuation(f.directory, f.input, idleOnly)).toMatchObject({ attempted: false, executing: true, ownerState: "busy" })
    f.state.queued = false

    f.history.push({ ...f.history[1], id: "msg_task", time: { created: 4 }, content: [
      { type: "tool", id: "call_task", name: "task", time: { created: 4 }, state: { status: "running", input: {}, metadata: { sessionId: "ses_child" } } },
    ] })
    f.state.active = { ses_child: { type: "running" } }
    expect(await readSessionExecutionState(f.directory, f.input.sessionID, f.options)).toBe("busy")
    f.state.active = {}
    expect(await readSessionExecutionState(f.directory, f.input.sessionID, f.options)).toBe("idle")
    f.history.pop()

    f.state.inboxAvailable = false
    expect(await readSessionExecutionState(f.directory, f.input.sessionID, f.options)).toBe("unavailable")
    expect(await dispatchChatContinuation(f.directory, f.input, idleOnly)).toMatchObject({ attempted: false, uncertain: true, ownerState: "unavailable" })
    expect(claims(f.directory)).toEqual([])
    expect(f.state.posts).toBe(0)

    f.state.inboxAvailable = true
    expect(await dispatchChatContinuation(f.directory, f.input, idleOnly)).toMatchObject({ accepted: true, persisted: true })
    expect(f.state.posts).toBe(1)
  })
})
