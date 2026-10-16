import { expect, test } from "bun:test"
import { OpenCode, type SessionMessageInfo, type V2Event } from "@opencode/client"
import { clientDirectories } from "../src/api/opencode/client"
import { subscribeEvents } from "../src/api/opencode/events"
import { draftingStageSession } from "../src/components/narration/use-drafting-narration"
import type { Event } from "../src/api/session-view"

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
  const client = OpenCode.make({ baseUrl: server.url.href })
  clientDirectories.set(client, "/matter")
  const controller = new AbortController()
  const listening = subscribeEvents(client, (event) => events.push(event), controller.signal, {
    acceptsSession,
    onFormsChanged: () => formsChanged++,
  })
  await until(() => Boolean(output))
  return {
    reads,
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

test("retirement discards a transcript response already in flight", async () => {
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
    view.send(deleted("ses_chat"))
    await until(() => view.events.some((event) => event.type === "session.deleted"))
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
