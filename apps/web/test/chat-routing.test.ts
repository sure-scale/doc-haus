import { describe, expect, test } from "bun:test"
import type { SessionInfo, SessionMessageInfo } from "@opencode/client"
import { abortSession, createSession, deleteSession } from "../src/api/opencode/sessions"
import { modelSpec } from "../src/api/opencode/config"
import { routeAgent } from "../src/api/opencode/routing"
import { recoveryHistoryNeedsCompaction, sendMalformedToolRecovery } from "../src/api/opencode/recovery"
import type { Client } from "../src/api/opencode/client"
import type { Message } from "../src/api/session-view"
import { CHAT_ASSISTANTS } from "../src/agents"
import { childSessionIDs, loadConversation } from "../src/components/chat/conversation"
import { isDraftingRecoveryMessage } from "../src/components/chat/drafting-recovery"
import {
  errorOf,
  toTurns,
  toolStepStatus,
  visibleReasoning,
} from "../src/components/chat/turn-model"
import {
  malformedToolRecoveryCount,
  malformedToolRecoveryDue,
  malformedToolRecoveryToken,
  shouldAbortStalledTurn,
  stalledAbortFailureMessage,
} from "../src/components/chat/turn-recovery"

function continuationPromptRequest(request: Request) {
  const path = new URL(request.url).pathname
  return path.startsWith("/matters/") && path.endsWith("/continuations/dispatch")
}

function continuationAccepted(existing = false, persisted = true) {
  return Response.json({
    accepted: true,
    attempted: true,
    existing,
    executing: true,
    persisted,
    uncertain: !persisted,
  })
}

function nativeSessionResponse(id: string, directory: string, model?: SessionInfo["model"]) {
  return Response.json({
    data: {
      id,
      projectID: "fixture-project",
      agent: "draft",
      model,
      location: { directory },
      cost: 0,
      tokens: { input: 0, output: 0, reasoning: 0, cache: { read: 0, write: 0 } },
      time: { created: 1, updated: 1 },
    } satisfies SessionInfo,
  })
}

describe("chat dispatch", () => {
  test("compacts recovery history after output exhaustion or a large persisted activity", () => {
    expect(
      recoveryHistoryNeedsCompaction([
        { info: { id: "length", role: "assistant", finish: "length" }, parts: [] },
      ] as never),
    ).toBe(true)
    expect(
      recoveryHistoryNeedsCompaction([
        { info: { id: "large", role: "user" }, parts: [{ type: "text", text: "x".repeat(256_000) }] },
        { info: { id: "stopped", role: "assistant", finish: "stop" }, parts: [] },
      ] as never),
    ).toBe(true)
    expect(
      recoveryHistoryNeedsCompaction([
        { info: { id: "small", role: "assistant", finish: "stop" }, parts: [{ type: "text", text: "done" }] },
      ] as never),
    ).toBe(false)
  })

  test("counts only the active projection after a completed session compaction", () => {
    const tail = { info: { id: "tail", role: "user" }, parts: [{ type: "text", text: "retained" }] }
    const history = [
      { info: { id: "old", role: "user" }, parts: [{ type: "text", text: "x".repeat(256_000) }] },
      tail,
      {
        info: { id: "compact", role: "user" },
        parts: [{ type: "compaction", tail_start_id: "tail" }],
      },
      {
        info: {
          id: "summary",
          role: "assistant",
          parentID: "compact",
          summary: true,
          finish: "stop",
        },
        parts: [{ type: "text", text: "summary" }],
      },
      { info: { id: "after", role: "assistant", finish: "stop" }, parts: [{ type: "text", text: "done" }] },
    ]
    expect(recoveryHistoryNeedsCompaction(history as never)).toBe(false)

    tail.parts[0]!.text = "y".repeat(256_000)
    expect(recoveryHistoryNeedsCompaction(history as never)).toBe(true)
  })

  test("ignores incomplete compactions and fails open on an invalid retained-tail pointer", () => {
    const large = { info: { id: "old", role: "user" }, parts: [{ type: "text", text: "x".repeat(256_000) }] }
    expect(
      recoveryHistoryNeedsCompaction([
        large,
        { info: { id: "compact", role: "user" }, parts: [{ type: "compaction" }] },
        {
          info: { id: "summary", role: "assistant", parentID: "compact", summary: true },
          parts: [{ type: "text", text: "unfinished" }],
        },
      ] as never),
    ).toBe(true)
    expect(
      recoveryHistoryNeedsCompaction([
        large,
        {
          info: { id: "compact", role: "user" },
          parts: [{ type: "compaction", tail_start_id: "missing" }],
        },
        {
          info: { id: "summary", role: "assistant", parentID: "compact", summary: true, finish: "stop" },
          parts: [{ type: "text", text: "summary" }],
        },
      ] as never),
    ).toBe(true)
  })

  test("uses the latest completed compaction boundary", () => {
    expect(
      recoveryHistoryNeedsCompaction([
        { info: { id: "old", role: "user" }, parts: [{ type: "text", text: "x".repeat(256_000) }] },
        { info: { id: "compact-one", role: "user" }, parts: [{ type: "compaction" }] },
        {
          info: { id: "summary-one", role: "assistant", parentID: "compact-one", summary: true, finish: "stop" },
          parts: [{ type: "text", text: "first" }],
        },
        { info: { id: "large-tail", role: "user" }, parts: [{ type: "text", text: "y".repeat(256_000) }] },
        { info: { id: "compact-two", role: "user" }, parts: [{ type: "compaction" }] },
        {
          info: { id: "summary-two", role: "assistant", parentID: "compact-two", summary: true, finish: "stop" },
          parts: [{ type: "text", text: "second" }],
        },
      ] as never),
    ).toBe(false)
  })

  test("bounds session creation when the engine never resolves", async () => {
    const client = {
      session: { create: () => new Promise<never>(() => {}) },
    } as unknown as Client
    await expect(createSession(client, "Matter conversation", 5)).rejects.toThrow("Session creation timed out.")
  })

  test("bounds session deletion when cleanup never resolves", async () => {
    const client = {
      session: { remove: () => new Promise<never>(() => {}) },
    } as unknown as Client
    await expect(deleteSession(client, "ses_delete", 5)).rejects.toThrow("Session deletion timed out.")
  })

  test("bounds a stalled-turn abort when the engine never resolves", async () => {
    const client = {
      session: { interrupt: () => new Promise<never>(() => {}) },
    } as unknown as Client
    await expect(abortSession(client, "ses_abort", 5)).rejects.toThrow("Session abort timed out.")
  })

  test("surfaces a terminal connection error after a stalled-turn abort fails", () => {
    expect(stalledAbortFailureMessage(new Error("Session abort timed out."))).toBe(
      "The engine did not respond while stopping the stalled turn. Session abort timed out. Check the connection, then retry.",
    )
  })

  test("pins the global Settings model on a prompt", () => {
    expect(modelSpec({ model: "google/gemma-4-31b-it" })).toEqual({
      providerID: "google",
      modelID: "gemma-4-31b-it",
    })
  })

  test("recognizes a durable recovery user message even without the legacy synthetic flag", () => {
    expect(
      isDraftingRecoveryMessage({
        parts: [
          {
            type: "text",
            text: '<dochaus-recovery kind="malformed-tool" run-id="run_current" progress-token="abc123" />\nContinue.',
          },
        ],
      } as never),
    ).toBe(true)
  })

  test("a saved turn keeps each step's text apart, as the live view does", () => {
    const text = (id: string, messageID: string, value: string) => ({
      type: "text" as const,
      id,
      sessionID: "ses_current",
      messageID,
      text: value,
      time: { start: 1, end: 2 },
    })
    const [, answer] = toTurns(
      [
        { info: { id: "msg_user", role: "user", agent: "compare" }, parts: [text("part-user", "msg_user", "Compare the drafts.")] },
        { info: { id: "msg_step1", role: "assistant" }, parts: [text("part-1", "msg_step1", "Comparing the two drafts.")] },
        { info: { id: "msg_step2", role: "assistant" }, parts: [text("part-2", "msg_step2", "**Changes**\n\nOne change.")] },
      ],
      "",
    )
    expect(answer!.text).toBe("Comparing the two drafts.\n\n**Changes**\n\nOne change.")
  })

  test("keeps a hidden recovery instruction as an assistant-turn boundary", () => {
    const text = (id: string, messageID: string, value: string) => ({
      type: "text" as const,
      id,
      sessionID: "ses_current",
      messageID,
      text: value,
      time: { start: 1, end: 2 },
    })
    const turns = toTurns(
      [
        { info: { id: "msg_user", role: "user", agent: "draft" }, parts: [text("part-user", "msg_user", "Draft this.")] },
        {
          info: { id: "msg_failed", role: "assistant", error: { name: "ProviderError", message: "Interrupted." } },
          parts: [],
        },
        {
          info: { id: "msg_recovery", role: "user" },
          parts: [
            text(
              "part-recovery",
              "msg_recovery",
              '<dochaus-recovery kind="malformed-tool" run-id="run_current" progress-token="abc123" />\nContinue.',
            ),
          ],
        },
        { info: { id: "msg_success", role: "assistant" }, parts: [text("part-success", "msg_success", "Completed.")] },
      ] as never,
      "Matter",
    )
    expect(turns.at(-1)).toMatchObject({ role: "assistant", text: "Completed.", error: undefined })
  })

  test("dispatches malformed-turn recovery with only the exact server-owned identity", async () => {
    const original = globalThis.fetch
    const requests: Request[] = []
    globalThis.fetch = ((input: string | URL | Request, init?: RequestInit) => {
      const request = input instanceof Request ? input : new Request(input, init)
      const pathname = new URL(request.url).pathname
      if (pathname === "/api/session/ses_owner/message") return Promise.resolve(Response.json({ data: [], cursor: {} }))
      if (pathname === "/api/session/ses_owner")
        return Promise.resolve(nativeSessionResponse("ses_owner", "/tmp/matter", { providerID: "google-vertex", id: "owner-model" }))
      if (pathname === "/engine/config") return Promise.resolve(Response.json({}))
      if (continuationPromptRequest(request)) {
        requests.push(request)
        return Promise.resolve(continuationAccepted())
      }
      return Promise.reject(new Error(`Unexpected fixture request: ${request.method} ${pathname}`))
    }) as typeof fetch
    try {
      const result = await sendMalformedToolRecovery("/tmp/matter", "ses_owner", {
        kind: "malformed-tool",
        runId: "ses_owner",
        progressToken: "malformed:msg_failed",
        messageID: "msg_recovery_msg_failed",
      })
      expect(result).toEqual({ admitted: true, executing: true, persisted: true })
      expect(requests).toHaveLength(1)
      expect(new URL(requests[0]!.url).pathname).toBe("/matters/matter/continuations/dispatch")
      expect(await requests[0]!.clone().json()).toEqual({
        kind: "malformed-tool",
        sessionID: "ses_owner",
        progressToken: "malformed:msg_failed",
      })
    } finally {
      globalThis.fetch = original
    }
  })

  test("deduplicates concurrent malformed-turn recovery dispatches", async () => {
    const original = globalThis.fetch
    let promptCount = 0
    globalThis.fetch = ((input: string | URL | Request, init?: RequestInit) => {
      const request = input instanceof Request ? input : new Request(input, init)
      const pathname = new URL(request.url).pathname
      if (pathname === "/api/session/ses_concurrent/message") return Promise.resolve(Response.json({ data: [], cursor: {} }))
      if (pathname === "/api/session/ses_concurrent")
        return Promise.resolve(nativeSessionResponse("ses_concurrent", "/tmp/matter-concurrent"))
      if (pathname === "/engine/config") return Promise.resolve(Response.json({}))
      if (continuationPromptRequest(request)) {
        promptCount += 1
        return new Promise<Response>((resolve) => setTimeout(() => resolve(continuationAccepted()), 10))
      }
      return Promise.reject(new Error(`Unexpected fixture request: ${request.method} ${pathname}`))
    }) as typeof fetch
    try {
      const recovery = {
        kind: "malformed-tool" as const,
        runId: "ses_concurrent",
        progressToken: "malformed:msg_failed",
        messageID: "msg_recovery_msg_failed",
      }
      await Promise.all([
        sendMalformedToolRecovery("/tmp/matter-concurrent", "ses_concurrent", recovery),
        sendMalformedToolRecovery("/tmp/matter-concurrent", "ses_concurrent", recovery),
      ])
      expect(promptCount).toBe(1)
    } finally {
      globalThis.fetch = original
    }
  })

  test("releases a failed malformed-turn dispatch so a later attempt can run", async () => {
    const original = globalThis.fetch
    let promptCount = 0
    let rejectPrompt = true
    globalThis.fetch = ((input: string | URL | Request, init?: RequestInit) => {
      const request = input instanceof Request ? input : new Request(input, init)
      const pathname = new URL(request.url).pathname
      if (pathname === "/api/session/ses_release/message") return Promise.resolve(Response.json({ data: [], cursor: {} }))
      if (pathname === "/api/session/ses_release")
        return Promise.resolve(nativeSessionResponse("ses_release", "/tmp/matter-release"))
      if (pathname === "/engine/config") return Promise.resolve(Response.json({}))
      if (continuationPromptRequest(request)) {
        promptCount += 1
        if (rejectPrompt) return Promise.resolve(Response.json({ error: "rejected" }, { status: 500 }))
        return Promise.resolve(continuationAccepted())
      }
      return Promise.reject(new Error(`Unexpected fixture request: ${request.method} ${pathname}`))
    }) as typeof fetch
    try {
      const recovery = {
        kind: "malformed-tool" as const,
        runId: "ses_release",
        progressToken: "malformed:msg_failed",
        messageID: "msg_recovery_msg_failed",
      }
      await expect(
        sendMalformedToolRecovery("/tmp/matter-release", "ses_release", recovery),
      ).rejects.toThrow("rejected")
      rejectPrompt = false
      await sendMalformedToolRecovery("/tmp/matter-release", "ses_release", recovery)
      expect(promptCount).toBe(2)
    } finally {
      globalThis.fetch = original
    }
  })

  test("keeps root history when a display-only child session cannot load", async () => {
    const native = {
      id: "msg_root_assistant",
      type: "assistant",
      agent: "draft",
      model: { providerID: "fixture-provider", id: "fixture-model" },
      time: { created: 1, completed: 2 },
      content: [{
        type: "tool",
        id: "fixture-task",
        name: "task",
        state: { status: "completed", input: {}, content: [{ type: "text", text: "done" }], metadata: { sessionId: "ses_child_slow" } },
        time: { created: 1, ran: 1, completed: 2 },
      }],
    } satisfies SessionMessageInfo
    const root: Message = {
      info: { id: native.id, sessionID: "ses_root", role: "assistant", agent: "draft", parentID: undefined, finish: undefined, error: undefined, time: native.time },
      parts: [{
        type: "tool", id: "fixture-task", sessionID: "ses_root", messageID: native.id, tool: "task",
        state: { status: "completed", input: {}, output: "done", title: "task", metadata: { sessionId: "ses_child_slow" } },
      }],
    }
    const requests: Parameters<Client["message"]["list"]>[0][] = []
    const client = {
      message: {
        list: (input: Parameters<Client["message"]["list"]>[0]) => {
          requests.push(input)
          if (input.sessionID === "ses_root" && !input.cursor)
            return Promise.resolve({ data: [native], cursor: { next: "root-terminal-page" } })
          if (input.sessionID === "ses_root" && input.cursor === "root-terminal-page")
            return Promise.resolve({ data: [], cursor: {} })
          if (input.sessionID === "ses_child_slow") return new Promise(() => undefined)
          return Promise.reject(new Error("Unexpected fixture message request"))
        },
      },
    }
    const conversation = await loadConversation(client as never, "ses_root", {
      descendantTimeoutMs: 5,
      descendantBudgetMs: 10,
    })
    expect(conversation.msgs).toEqual([root])
    expect(conversation.pool).toEqual(root.parts)
    expect(requests).toEqual([
      { sessionID: "ses_root", limit: 200, order: "asc" },
      { sessionID: "ses_root", limit: 200, cursor: "root-terminal-page" },
      { sessionID: "ses_child_slow", limit: 200, order: "asc" },
    ])
  })

  test("does not abort a quiet turn while an approval is pending", () => {
    const stalled = {
      busy: true,
      sessionID: "ses_current",
      aborting: false,
      lastActivityAt: 0,
      now: 30 * 60_000,
    }
    expect(shouldAbortStalledTurn({ ...stalled, pendingPermissions: 1 })).toBe(false)
    expect(shouldAbortStalledTurn({ ...stalled, pendingPermissions: 0 })).toBe(true)
    expect(shouldAbortStalledTurn({ ...stalled, pendingPermissions: 0, now: 30 * 60_000 - 1 })).toBe(false)
  })

  test("uses a stable token for generic printed tool syntax", () => {
    const message = {
      info: { id: "msg_malformed" },
      parts: [{ type: "reasoning", text: 'call:any-future-tool{"input":true}' }],
    }
    expect(malformedToolRecoveryToken(message as never)).toBe("malformed:msg_malformed")
    expect(malformedToolRecoveryToken(message as never)).toBe("malformed:msg_malformed")
    expect(
      malformedToolRecoveryToken({
        info: { id: "msg_text_malformed" },
        parts: [{ type: "text", text: '<|tool_call|>call:any-future-tool{"input":true}' }],
      } as never),
    ).toBe("malformed:msg_text_malformed")
    expect(
      malformedToolRecoveryToken({
        ...message,
        parts: [...message.parts, { type: "text", text: "Here is the completed answer." }],
      } as never),
    ).toBeUndefined()
  })

  test("restores the malformed recovery bound from the current user activity", () => {
    const marker = (id: string) => ({
      info: { id, role: "user" },
      parts: [
        {
          type: "text",
          text: `<dochaus-recovery kind="malformed-tool" run-id="ses_current" progress-token="malformed:${id}" />`,
        },
      ],
    })
    const history = [
      { info: { id: "first", role: "user" }, parts: [{ type: "text", text: "First request" }] },
      marker("old-recovery"),
      { info: { id: "second", role: "user" }, parts: [{ type: "text", text: "Second request" }] },
      marker("recovery-one"),
      marker("recovery-two"),
    ] as never
    expect(malformedToolRecoveryCount(history)).toBe(2)
  })

  test("recovers a printed tool call for every tool-using chat agent within the bound", () => {
    const prompt = (agent: string) => ({
      info: { id: `msg_${agent}`, role: "user", agent },
      parts: [{ type: "text", text: "Summarize the key terms." }],
    })
    const printed = (id: string) => ({
      info: { id, role: "assistant" },
      parts: [{ type: "reasoning", text: 'call:search-document{"query":"term"}' }],
    })
    const recovery = (id: string) => ({
      info: { id: `msg_recovery_${id}`, role: "user" },
      parts: [{ type: "text", text: `<dochaus-recovery kind="malformed-tool" run-id="ses_chat" progress-token="malformed:${id}" />` }],
    })
    // "draft" starts a lean run in its own sessions, so it never holds a chat turn.
    const agents = CHAT_ASSISTANTS.map((assistant) => assistant.name).filter((name) => name !== "draft")
    expect(agents.length).toBeGreaterThan(0)
    for (const agent of agents)
      expect(malformedToolRecoveryDue([prompt(agent), printed("msg_failed")] as never, "matter")).toBe("malformed:msg_failed")
    const history = [
      prompt("qa"),
      printed("msg_one"),
      recovery("msg_one"),
      printed("msg_two"),
      recovery("msg_two"),
      printed("msg_three"),
    ]
    expect(malformedToolRecoveryDue(history.slice(0, 4) as never, "matter")).toBe("malformed:msg_two")
    expect(malformedToolRecoveryDue(history as never, "matter")).toBeUndefined()
    expect(malformedToolRecoveryDue([...history, prompt("research"), printed("msg_next")] as never, "matter")).toBe("malformed:msg_next")
    expect(malformedToolRecoveryDue([prompt("qa"), printed("msg_failed")] as never, undefined)).toBeUndefined()
  })
})

describe("reasoning display", () => {
  test("drops separator-only reasoning", () => {
    expect(visibleReasoning("---\n")).toBeUndefined()
  })

  test("preserves substantive reasoning markdown", () => {
    expect(visibleReasoning("**Checking the term sheet**\n\nI will inspect the parties."))
      .toBe("**Checking the term sheet**\n\nI will inspect the parties.")
  })
})

describe("provider errors", () => {
  test("surfaces a nested quota message from a stream failure", () => {
    expect(
      errorOf({
        name: "APIError",
        data: { error: { message: "Quota exceeded for gemma-4-31b. Please retry in 54s." } },
      }),
    ).toBe("Quota exceeded for gemma-4-31b. Please retry in 54s.")
  })

  // A throwaway router session whose every turn ends the way the engine records
  // it: the reply carries `reply` (an error or text) and the idle marker follows.
  function routerClient(reply: { error?: Record<string, unknown>; text?: string }) {
    const prompts: string[] = []
    const client = {
      session: {
        create: async () => ({ id: `ses_router_${prompts.length}` }),
        switchAgent: async () => ({}),
        switchModel: async (input: { model: { providerID: string; id: string } }) => {
          prompts.push(`${input.model.providerID}/${input.model.id}`)
          return {}
        },
        prompt: async () => ({}),
        wait: async () => undefined,
        interrupt: async () => ({}),
        remove: async () => ({}),
      },
      message: {
        list: async () => ({
          data: [
            { type: "user", id: "msg_user", text: "Route this.", time: { created: 1 } },
            {
              type: "assistant",
              id: "msg_reply",
              agent: "router",
              content: reply.text ? [{ type: "text", text: reply.text }] : [],
              ...(reply.error ? { finish: "error", error: reply.error } : { finish: "stop" }),
              time: { created: 2, completed: 3 },
            },
            { type: "idle", id: "msg_idle", outcome: reply.error ? "failed" : "succeeded", ...(reply.error ? { error: reply.error } : {}), time: { created: 4 } },
          ],
          cursor: {},
        }),
      },
    } as unknown as Client
    return { client, prompts }
  }

  const route = (client: Client) =>
    routeAgent({
      client,
      smallModel: "opencode-go/small",
      primaryModel: "opencode-go/primary",
      candidates: CHAT_ASSISTANTS.map((a) => ({ name: a.name, description: a.description })),
      history: [],
      text: "Review the lease.",
    })

  test("Auto surfaces the provider's own error when the router's reply fails", async () => {
    const expired = {
      type: "provider.auth",
      status: 401,
      message: "This API key has expired. Create a new API key in OpenCode Console.",
    }
    const { client, prompts } = routerClient({ error: expired })
    await expect(route(client)).rejects.toThrow(expired.message)
    // The existing small-then-primary escalation, and no further attempts.
    expect(prompts).toEqual(["opencode-go/small", "opencode-go/primary"])
  })

  test("Auto returns the draft route with its closed scope on the first model", async () => {
    const scope = { documents: "one", legal_effect: "no", outside_law: "no", firm_material: "no", sources: "few" } as const
    const { client, prompts } = routerClient({ text: JSON.stringify({ route: "draft", ...scope }) })
    expect(await route(client)).toEqual({ name: "draft", scope })
    expect(prompts).toEqual(["opencode-go/small"])
  })

  test("Auto no longer reads a candidate name out of prose: both models miss and the lawyer is told", async () => {
    const { client, prompts } = routerClient({ text: "The best candidate is draft." })
    await expect(route(client)).rejects.toThrow("Auto could not choose an assistant.")
    expect(prompts).toEqual(["opencode-go/small", "opencode-go/primary"])
  })

  test("Auto keeps the could-not-choose message for an empty router reply", async () => {
    const { client } = routerClient({ text: "" })
    await expect(route(client)).rejects.toThrow(
      "Auto could not choose an assistant. Try again or select an assistant manually.",
    )
  })
})

describe("task status", () => {
  test("renders a blocked drafting handoff as an error", () => {
    expect(
      toolStepStatus({
        type: "tool",
        id: "part",
        sessionID: "session",
        messageID: "message",
        tool: "task",
        state: {
          status: "completed",
          input: {},
          output: '<handoff_gate status="blocked"></handoff_gate>',
          title: "Drafting strategist",
          metadata: {},
        },
      }),
    ).toBe("error")
  })

  test("finds every unique child session needed to reconstruct nested steps", () => {
    const task = (id: string, child: string) => ({
      type: "tool" as const,
      id,
      sessionID: "parent",
      messageID: "message",
      callID: id,
      tool: "task",
      state: {
        status: "completed" as const,
        input: {},
        output: "done",
        title: "Task",
        metadata: { sessionId: child },
        time: { start: 0, end: 1 },
      },
    })
    expect(childSessionIDs([{ parts: [task("one", "child-a"), task("two", "child-a"), task("three", "child-b")] }])).toEqual([
      "child-a",
      "child-b",
    ])
  })
})
