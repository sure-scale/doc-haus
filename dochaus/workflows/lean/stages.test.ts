import { describe, expect, test } from "bun:test"
import { LEAN_CHILD_PERMISSIONS, engineSessionPort } from "./stages"

// The engine session port runs through the real v2 adapter (engineV2Fetch and the
// transcript projection). Only the engine is scripted: it stores what it is sent
// and answers like the native v2 routes.

const DIRECTORY = "/matters/lean-port"

function scriptedEngine(reply: { text: string; outcome: string; interrupted?: boolean; finalError?: boolean; onIdlePoll?: () => void }) {
  const created: Record<string, unknown>[] = []
  const admitted: Record<string, unknown>[] = []
  const interrupts: string[] = []
  const transcript: Record<string, unknown>[] = [
    { id: "msg_000000000001aaaaaaaaaaaaaa", type: "user", text: "earlier turn", time: { created: 1 } },
    { id: "msg_000000000002aaaaaaaaaaaaaa", type: "assistant", agent: "expander", time: { created: 2, completed: 3 }, content: [{ type: "text", text: "earlier reply" }] },
    { id: "idle_earlier", type: "idle", outcome: "succeeded", time: { created: 4 } },
  ]
  const queued: { id: string; text: string }[] = []
  const idlePolls = { count: 0 }
  const session = { id: "ses_child", parentID: "ses_owner", agent: "expander", location: { directory: DIRECTORY }, model: { providerID: "test", id: "scripted" } }
  const fetcher = (async (target: string | URL | Request, init?: RequestInit) => {
    const url = new URL(typeof target === "string" || target instanceof URL ? target : target.url)
    const method = init?.method ?? "GET"
    if (url.pathname === "/api/dochaus/session" && method === "POST") {
      created.push(JSON.parse(String(init?.body)))
      return Response.json({ data: session })
    }
    if (url.pathname === "/api/session/ses_child" && method === "GET") return Response.json({ data: session })
    if (url.pathname === "/api/session/ses_child/prompt" && method === "POST") {
      const body = JSON.parse(String(init?.body)) as { id: string; text: string; delivery: string }
      admitted.push(body)
      queued.push(body)
      return Response.json({ data: { id: body.id, sessionID: "ses_child", type: "user", payload: { text: body.text }, delivery: body.delivery, time: { created: 5 } } })
    }
    if (url.pathname === "/api/session/ses_child/interrupt" && method === "POST") {
      interrupts.push(url.pathname)
      // The native interrupt answers without the v2 data envelope.
      return Response.json({ interrupted: true })
    }
    if (url.pathname === "/api/session/ses_child/message" && method === "GET") {
      if (url.searchParams.get("limit") === "1") reply.onIdlePoll?.()
      // The first idle poll still sees the earlier turn's idle record; the
      // queued prompt is promoted and answered before the second.
      if (url.searchParams.get("limit") === "1" && idlePolls.count++ === 1)
        queued.splice(0).forEach(prompt => transcript.push(
          { id: prompt.id, type: "user", text: prompt.text, time: { created: 6 } },
          ...(reply.interrupted ? [
            { id: "msg_000000000007aaaaaaaaaaaaaa", type: "assistant", agent: "expander", time: { created: 6, completed: 6 }, finish: "error",
              error: { type: "provider.transport", message: "Decode error", status: 200 }, content: [] },
            { id: "msg_000000000008aaaaaaaaaaaaaa", type: "synthetic", text: "The previous response was interrupted. Continue.", time: { created: 7 } },
          ] : []),
          { id: "msg_000000000009aaaaaaaaaaaaaa", type: "assistant", agent: "expander", time: { created: 7, completed: 8 }, content: [{ type: "text", text: reply.text }],
            ...(reply.finalError ? { finish: "error", error: { type: "provider.transport", message: "Decode error", status: 200 } } : {}) },
          { id: "idle_latest", type: "idle", outcome: reply.outcome, time: { created: 9 } },
        ))
      const limit = Number(url.searchParams.get("limit") ?? transcript.length)
      return Response.json({ data: transcript.toReversed().slice(0, limit), cursor: {} })
    }
    return new Response("unhandled", { status: 500 })
  }) as typeof fetch
  return { fetcher, created, admitted, idlePolls, interrupts }
}

describe("lean engine session port", () => {
  test("creates a permission-scoped child and returns the final reply to the admitted prompt, not an earlier idle turn", async () => {
    const engine = scriptedEngine({ text: "# Draft\n\nBody.", outcome: "succeeded" })
    const port = engineSessionPort({ engineURL: "http://engine.test", fetch: engine.fetcher, pollMs: 1 })
    const sessionID = await port.createSession({
      directory: DIRECTORY, parentID: "ses_owner", agent: "expander", title: "Lean drafting: draft",
      model: { providerID: "test", modelID: "scripted" }, permissions: LEAN_CHILD_PERMISSIONS,
    })
    const reply = await port.prompt({ directory: DIRECTORY, sessionID, agent: "expander", text: "Draft it." })

    expect(sessionID).toBe("ses_child")
    expect(engine.created[0]).toMatchObject({
      parentID: "ses_owner", agent: "expander", model: { providerID: "test", id: "scripted" },
      permissions: LEAN_CHILD_PERMISSIONS, metadata: { dochaus_role: "lean-drafting" },
    })
    expect(engine.admitted).toEqual([{ id: expect.stringMatching(/^msg_[0-9a-f]{12}[0-9A-Za-z]{14}$/), text: "Draft it.", delivery: "queue", metadata: {} }])
    expect(engine.idlePolls.count).toBe(2)
    expect(reply).toBe("# Draft\n\nBody.")
  })

  test("a failed native turn is an engine error, not a reply", async () => {
    const engine = scriptedEngine({ text: "partial", outcome: "failed" })
    const port = engineSessionPort({ engineURL: "http://engine.test", fetch: engine.fetcher, pollMs: 1 })
    await expect(port.prompt({ directory: DIRECTORY, sessionID: "ses_child", agent: "expander", text: "Draft it." }))
      .rejects.toThrow("The expander session ended failed.")
  })

  test("a provider interruption the engine recovered from does not fail the turn", async () => {
    const engine = scriptedEngine({ text: "# Memo\n\nBody.", outcome: "succeeded", interrupted: true })
    const port = engineSessionPort({ engineURL: "http://engine.test", fetch: engine.fetcher, pollMs: 1 })
    expect(await port.prompt({ directory: DIRECTORY, sessionID: "ses_child", agent: "expander", text: "Draft it." })).toBe("# Memo\n\nBody.")
  })

  test("a pause interrupts the native session and stops waiting for its reply", async () => {
    const engine = scriptedEngine({ text: "# Memo\n\nBody.", outcome: "succeeded" })
    const port = engineSessionPort({ engineURL: "http://engine.test", fetch: engine.fetcher, pollMs: 1 })
    await port.interrupt!({ directory: DIRECTORY, sessionID: "ses_child" })
    expect(engine.interrupts).toEqual(["/api/session/ses_child/interrupt"])

    // Paused while the prompt was being admitted: the admitted turn is interrupted at once.
    const admitting = new AbortController()
    const reply = port.prompt({ directory: DIRECTORY, sessionID: "ses_child", agent: "expander", text: "Draft it.", signal: admitting.signal })
    admitting.abort()
    await expect(reply).rejects.toThrow("The expander session was paused.")
    expect(engine.admitted).toHaveLength(1)
    expect(engine.interrupts).toHaveLength(2)
    expect(engine.idlePolls.count).toBe(0)

    // Paused while waiting: polling stops; the coordinator has interrupted the session.
    const waiting = new AbortController()
    const polled = scriptedEngine({ text: "# Memo\n\nBody.", outcome: "succeeded", onIdlePoll: () => waiting.abort() })
    const polling = engineSessionPort({ engineURL: "http://engine.test", fetch: polled.fetcher, pollMs: 1 })
    await expect(polling.prompt({ directory: DIRECTORY, sessionID: "ses_child", agent: "expander", text: "Draft it.", signal: waiting.signal }))
      .rejects.toThrow("The expander session was paused.")
    expect(polled.idlePolls.count).toBe(1)
  })

  test("a turn whose last reply carries an error fails even when the session idles as succeeded", async () => {
    const engine = scriptedEngine({ text: "partial", outcome: "succeeded", interrupted: true, finalError: true })
    const port = engineSessionPort({ engineURL: "http://engine.test", fetch: engine.fetcher, pollMs: 1 })
    await expect(port.prompt({ directory: DIRECTORY, sessionID: "ses_child", agent: "expander", text: "Draft it." }))
      .rejects.toThrow("The expander session ended succeeded: {\"type\":\"provider.transport\"")
  })
})
