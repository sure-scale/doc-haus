import { expect, test } from "bun:test"
import { OpenCode, type SessionCreateInput, type SessionMessageInfo, type V2Event } from "@opencode/client"
import { clientDirectories, type Client } from "../src/api/opencode/client"
import { subscribeEvents } from "../src/api/opencode/events"
import { routeAgent } from "../src/api/opencode/routing"
import { createSession } from "../src/api/opencode/sessions"
import { projectEvent, projectMessages, type Event } from "../src/api/session-view"
import { toTurns } from "../src/components/chat/turn-model"
import { isConversation } from "../src/components/dashboard/data"
import { draftingStageSession } from "../src/components/narration/use-drafting-narration"
import { isUserConversation } from "../../../dochaus/lib/session-visibility"

function text(sessionID: string, directory = "/matter"): V2Event {
  return {
    id: `evt_${sessionID}`,
    created: 1,
    type: "session.text.delta",
    location: { directory },
    data: { sessionID, assistantMessageID: "msg_answer", ordinal: 0, delta: "Working" },
  }
}

function created(sessionID: string, directory: string): V2Event {
  return {
    id: `evt_created_${sessionID}`,
    created: 1,
    type: "session.created",
    location: { directory },
    durable: { aggregateID: sessionID, seq: 1, version: 1 },
    data: { sessionID, projectID: "shared-project", location: { directory }, slug: "conversation", version: "test" },
  }
}

function deleted(sessionID: string): V2Event {
  return {
    id: `evt_delete_${sessionID}`,
    created: 2,
    type: "session.deleted",
    location: { directory: "/matter" },
    durable: { aggregateID: sessionID, seq: 2, version: 2 },
    data: { sessionID },
  }
}

async function until(condition: () => boolean) {
  const deadline = Date.now() + 2_000
  while (!condition() && Date.now() < deadline) await Bun.sleep(5)
  expect(condition()).toBe(true)
}

async function stream(acceptsSession: (id: string) => boolean) {
  const reads: string[] = []
  const readSignals: Array<AbortSignal | null | undefined> = []
  const events: Event[] = []
  let formsChanged = 0
  let output: ReadableStreamDefaultController<Uint8Array> | undefined
  let hold: Promise<void> | undefined
  const server = Bun.serve({
    port: 0,
    async fetch(request) {
      const path = new URL(request.url).pathname
      if (path === "/api/event")
        return new Response(
          new ReadableStream<Uint8Array>({
            start(controller) {
              output = controller
            },
          }),
          { headers: { "content-type": "text/event-stream" } },
        )
      reads.push(path.split("/")[3]!)
      await hold
      return Response.json({
        data: [
          {
            id: "msg_answer",
            type: "assistant",
            agent: "draft",
            model: { providerID: "test", id: "test" },
            time: { created: 1 },
            content: [{ type: "text", text: "Working" }],
          } satisfies SessionMessageInfo,
        ],
        cursor: {},
      })
    },
  })
  const client = OpenCode.make({ baseUrl: server.url.href, fetch: ((input, init) => {
    const request = new Request(input, init)
    if (new URL(request.url).pathname.endsWith("/message")) readSignals.push(request.signal)
    return fetch(request)
  }) as typeof fetch })
  clientDirectories.set(client, "/matter")
  const controller = new AbortController()
  const listening = subscribeEvents(client, (event) => events.push(event), controller.signal, {
    acceptsSession,
    onFormsChanged: () => formsChanged++,
  })
  await until(() => Boolean(output))
  return {
    reads,
    readSignals,
    events,
    controller,
    formsChanged: () => formsChanged,
    holdReads(pending: Promise<void>) {
      hold = pending
    },
    send(event: V2Event) {
      output!.enqueue(new TextEncoder().encode(`data: ${JSON.stringify(event)}\n\n`))
    },
    async close() {
      controller.abort()
      server.stop(true)
      await listening
    },
  }
}

test("global native events read only the active chat, known children and current drafting stages", async () => {
  // An unknown router can be in the SAME matter; another conversation can share
  // the SAME project. Neither has conversation ownership.
  const admitted = new Set(["ses_chat", "ses_child"])
  const status = {
    stages: {
      drafting: {
        status: "running" as const,
        attempts: 1,
        started_at: null,
        finished_at: null,
        session_ids: ["ses_stage"],
      },
    },
  }
  const owns = (id: string) => admitted.has(id) || draftingStageSession(status, id)
  const view = await stream(owns)
  try {
    view.send(text("ses_chat"))
    await until(() => view.events.some((event) => event.type === "message.part.updated"))
    view.send(created("ses_chat", "/matter"))
    view.send(created("ses_router", "/matter"))
    view.send(created("ses_same_project_other_matter", "/other-matter"))
    view.send(text("ses_router"))
    view.send(text("ses_same_project_other_matter", "/other-matter"))
    view.send(text("ses_chat", "/other-matter"))
    view.send(text("ses_child"))
    view.send(text("ses_stage"))
    // Native events may omit location; explicit known-session ownership still
    // permits recovery rather than discarding legitimate progress.
    const withoutLocation = text("ses_chat")
    delete withoutLocation.location
    view.send(withoutLocation)
    await until(() => view.reads.length >= 4)
    await Bun.sleep(30)
    expect(view.reads.sort()).toEqual(["ses_chat", "ses_chat", "ses_child", "ses_stage"])
    expect(
      view.events
        .filter((event) => event.type === "message.part.updated")
        .every((event) => event.type === "message.part.updated" && owns(event.properties.part.sessionID)),
    ).toBe(true)
  } finally {
    await view.close()
  }
})

test("retirement cancels an admitted refresh and reports deletion without adopting late helper events", async () => {
  const view = await stream((id) => id === "ses_chat" || id === "ses_helper")
  try {
    // Positive control: the same native event produces a transcript read.
    view.send(text("ses_chat"))
    await until(() => view.reads.length === 1)
    view.send(text("ses_helper"))
    view.send(deleted("ses_helper"))
    view.send(text("ses_helper"))
    view.send(deleted("ses_chat"))
    await until(() => view.events.filter((event) => event.type === "session.deleted").length === 2)
    await Bun.sleep(140)
    expect(view.reads).toEqual(["ses_chat"])
    expect(
      view.events.filter((event) => event.type === "session.deleted").map((event) => event.properties.sessionID),
    ).toEqual(["ses_helper", "ses_chat"])
  } finally {
    await view.close()
  }
})

test("ownership is rechecked at flush and abort cancels queued reads", async () => {
  const status = {
    stages: {
      drafting: {
        status: "running" as const,
        attempts: 1,
        started_at: null,
        finished_at: null,
        session_id: "ses_old_stage",
      },
    },
  }
  const view = await stream((id) => id === "ses_chat" || draftingStageSession(status, id))
  try {
    view.send(text("ses_chat"))
    await until(() => view.reads.length === 1)
    view.send(text("ses_old_stage"))
    status.stages.drafting.session_id = "ses_new_stage"
    await Bun.sleep(140)
    expect(view.reads).toEqual(["ses_chat"])
    view.send(text("ses_chat"))
    view.controller.abort()
    await Bun.sleep(140)
    expect(view.reads).toEqual(["ses_chat"])
  } finally {
    await view.close()
  }
})

test("retirement aborts a transcript request already in flight and discards its response", async () => {
  const view = await stream((id) => id === "ses_chat")
  let release!: () => void
  view.holdReads(
    new Promise<void>((resolve) => {
      release = resolve
    }),
  )
  try {
    view.send(text("ses_chat"))
    await until(() => view.reads.length === 1)
    expect(view.readSignals[0]?.aborted).toBe(false)
    view.send(deleted("ses_chat"))
    await until(() => view.events.some((event) => event.type === "session.deleted"))
    expect(view.readSignals[0]?.aborted).toBe(true)
    release()
    await Bun.sleep(30)
    expect(view.events.map((event) => event.type)).toEqual(["session.deleted"])
  } finally {
    release()
    await view.close()
  }
})

test("active failures and forms remain visible while foreign forms do not trigger reads", async () => {
  const view = await stream((id) => id === "ses_chat")
  try {
    view.send({
      id: "evt_failure",
      created: 1,
      type: "session.step.failed",
      location: { directory: "/matter" },
      durable: { aggregateID: "ses_chat", seq: 1, version: 1 },
      data: {
        sessionID: "ses_chat",
        assistantMessageID: "msg_answer",
        error: { type: "provider.error", message: "Unavailable" },
      },
    })
    for (const sessionID of ["ses_chat", "ses_router"])
      view.send({
        id: `evt_form_${sessionID}`,
        created: 1,
        type: "form.created",
        location: { directory: "/matter" },
        data: {
          form: { id: "form_question", sessionID, title: "Questions", fields: [{ key: "answer", type: "string" }] },
        },
      })
    for (const sessionID of ["ses_chat", "ses_router"]) {
      view.send({
        id: `evt_replied_${sessionID}`,
        created: 2,
        type: "form.replied",
        location: { directory: "/matter" },
        data: { id: "form_question", sessionID, answer: { answer: "Continue" } },
      })
      view.send({
        id: `evt_cancelled_${sessionID}`,
        created: 3,
        type: "form.cancelled",
        location: { directory: "/matter" },
        data: { id: "form_question", sessionID },
      })
    }
    await until(() => view.reads.length === 1)
    expect(view.events[0]).toMatchObject({
      type: "session.error",
      properties: { error: { data: { message: "Unavailable" } } },
    })
    expect(view.formsChanged()).toBe(3)
    expect(view.reads).toEqual(["ses_chat"])
  } finally {
    await view.close()
  }
})

const stopModel = { providerID: "test", id: "model" }
const stopTranscript = [
  { type: "user", id: "msg_u1", text: "Who are the parties?", time: { created: 1 } },
  { type: "assistant", id: "msg_a1", agent: "qa", model: stopModel, finish: "stop", time: { created: 2, completed: 3 }, content: [{ type: "text", text: "Acme Ltd and Beta LLP." }] },
  { type: "idle", id: "msg_i1", outcome: "succeeded", time: { created: 4 } },
  { type: "user", id: "msg_u2", text: "Summarise the letter.", time: { created: 5 } },
  { type: "assistant", id: "msg_a2", agent: "qa", model: stopModel, finish: "tool-calls", time: { created: 6, completed: 7 }, content: [{ type: "text", text: "I'll read the letter." }] },
  { type: "assistant", id: "msg_a3", agent: "qa", model: stopModel, finish: "error", error: { type: "aborted", message: "Step interrupted" }, time: { created: 8, completed: 9 }, content: [{ type: "text", text: "The letter demands" }] },
  { type: "idle", id: "msg_i2", outcome: "interrupted", time: { created: 10 } },
  { type: "user", id: "msg_u3", text: "When is payment due?", time: { created: 11 } },
  { type: "assistant", id: "msg_a4", agent: "qa", model: stopModel, finish: "tool-calls", time: { created: 12, completed: 13 }, content: [{ type: "text", text: "Checking the terms." }] },
  { type: "idle", id: "msg_i3", outcome: "interrupted", time: { created: 14 } },
  { type: "user", id: "msg_u4", text: "Who signed it?", time: { created: 15 } },
  { type: "assistant", id: "msg_a5", agent: "qa", model: stopModel, finish: "error", error: { type: "aborted", message: "Step interrupted" }, time: { created: 16, completed: 17 }, content: [] },
  { type: "idle", id: "msg_i4", outcome: "interrupted", time: { created: 18 } },
  { type: "user", id: "msg_u5", text: "Which invoice is unpaid?", time: { created: 19 } },
  { type: "assistant", id: "msg_a6", agent: "qa", model: stopModel, finish: "error", error: { type: "provider.unknown", message: "The provider is unavailable." }, time: { created: 20, completed: 21 }, content: [] },
  { type: "idle", id: "msg_i5", outcome: "failed", time: { created: 22 } },
] as SessionMessageInfo[]

test("an interrupted turn's last answer step is stored as stopped; other turns are not", () => {
  const stopped = projectMessages("ses", stopTranscript)
    .filter((message) => message.info.stopped)
    .map((message) => message.info.id)
  expect(stopped).toEqual(["msg_a3", "msg_a4", "msg_a5"])
})

test("a stopped answer keeps its partial output, is marked stopped, and does not read as an error", () => {
  const turns = toTurns(projectMessages("ses", stopTranscript), "matter").filter((turn) => turn.role === "assistant")
  expect(turns.map(({ text, stopped, error }) => ({ text, stopped: Boolean(stopped), error }))).toEqual([
    { text: "Acme Ltd and Beta LLP.", stopped: false, error: undefined },
    { text: "I'll read the letter.\n\nThe letter demands", stopped: true, error: undefined },
    { text: "Checking the terms.", stopped: true, error: undefined },
    { text: "", stopped: true, error: undefined },
    { text: "", stopped: false, error: "The provider is unavailable." },
  ])
})

function producerFixture() {
  const createdSessions: SessionCreateInput[] = [], removed: string[] = []
  const client = {
    session: {
      create: async (input: SessionCreateInput) => { createdSessions.push(input); return { id: "automatic", ...input } },
      switchAgent: async () => {}, switchModel: async () => {}, prompt: async () => {}, wait: async () => {}, interrupt: async () => {},
      remove: async ({ sessionID }: { sessionID: string }) => { removed.push(sessionID) },
    },
    message: { list: async () => ({ data: [{ type: "assistant", id: "answer", agent: "router", time: { created: 1, completed: 2 }, content: [{ type: "text", text: '{"route":"qa"}' }] }], cursor: {} }) },
  } as unknown as Client
  clientDirectories.set(client, "/synthetic/matter")
  return { client, created: createdSessions, removed }
}

test("actual router producer creates an explicitly owned session excluded by conversation and dashboard predicates", async () => {
  const { client, created: createdSessions, removed } = producerFixture()
  expect(await routeAgent({ client, smallModel: "fixture/fast", candidates: [{ name: "qa", description: "Answer" }], history: [], text: "Explain this clause" })).toEqual({ name: "qa" })
  expect(createdSessions).toHaveLength(1)
  expect(createdSessions[0]).toMatchObject({ location: { directory: "/synthetic/matter" }, metadata: { dochaus_role: "chat-router" } })
  expect(createdSessions[0]).not.toHaveProperty("parentID")
  const session = { metadata: createdSessions[0]!.metadata }
  expect(isUserConversation(session)).toBe(false)
  expect(isConversation(session)).toBe(false)
  expect(removed).toEqual(["automatic"])
})

test("ordinary session producer preserves human titles and metadata-free legacy visibility", async () => {
  const { client, created: createdSessions } = producerFixture()
  const title = "doc.haus grid: My own discussion"
  await createSession(client, title, 1000, [{ action: "*", resource: "*", effect: "deny" }])
  expect(createdSessions[0]).toEqual({ title, location: { directory: "/synthetic/matter" }, permissions: [{ action: "*", resource: "*", effect: "deny" }] })
  const session = { metadata: createdSessions[0]!.metadata }
  expect(isUserConversation(session)).toBe(true)
  expect(isConversation(session)).toBe(true)
})

test("native session messages retain the chat's agent, text, tools, and citations", () => {
  const messages = [
    { type: "agent-switched", id: "msg_agent", agent: "extract", time: { created: 1 } },
    { type: "user", id: "msg_user", text: "Summarise", time: { created: 2 } },
    {
      type: "assistant", id: "msg_answer", agent: "extract", model: { providerID: "test", id: "model" },
      time: { created: 3, completed: 5 },
      content: [
        { type: "tool", id: "call_search", name: "search-document", time: { created: 3, completed: 4 }, state: {
          status: "completed", input: { query: "Summarise" }, content: [{ type: "text", text: "Found passage" }],
          metadata: { citations: [{ section: "A", excerpt: "Passage" }] },
        } },
        { type: "text", text: "Answer" },
      ],
    },
  ] as SessionMessageInfo[]

  const projected = projectMessages("ses_matter", messages)
  expect(projected).toHaveLength(2)
  expect(projected[0]?.info.agent).toBe("extract")
  expect(projected[1]?.info.parentID).toBe("msg_user")
  expect(projected[1]?.parts[0]).toMatchObject({
    type: "tool", tool: "search-document", state: { status: "completed", output: "Found passage", metadata: { citations: [{ section: "A" }] } },
  })
  expect(projected[1]?.parts[1]).toMatchObject({ type: "text", text: "Answer" })
})

test("native permission events retain the decision target and document resource", () => {
  const event = {
    type: "permission.asked", id: "evt_permission", created: 1,
    data: { id: "per_1", sessionID: "ses_matter", action: "redline", resources: ["contract.docx"], source: { type: "tool", messageID: "msg_answer", id: "call_edit" } },
  } as V2Event
  expect(projectEvent(event)).toEqual([{
    type: "permission.asked",
    properties: {
      id: "per_1", sessionID: "ses_matter", permission: "redline", patterns: ["contract.docx"],
      metadata: {}, always: [], tool: { messageID: "msg_answer", callID: "call_edit" },
    },
  }])
})

test("native execution success projects to session.idle; a step end projects nothing", () => {
  const native = (type: string, data: Record<string, unknown>) =>
    ({ type, id: `evt_${type}`, created: 1, durable: { aggregateID: "ses_matter", seq: 1, version: 1 }, data: { sessionID: "ses_matter", ...data } }) as V2Event
  expect(projectEvent(native("session.step.ended", { assistantMessageID: "msg_step", finish: "tool-calls" }))).toEqual([])
  expect(projectEvent(native("session.execution.succeeded", {}))).toEqual([
    { type: "session.idle", properties: { sessionID: "ses_matter" } },
  ])
})
