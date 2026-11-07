import { afterEach, describe, expect, test } from "bun:test"
import { providerStopKind } from "../../lib/provider-stop"
import { PACKAGE_SCOPE_QUESTION } from "./package-scope"
import { writeLeanPlan } from "./plan"
import { stageInputSha } from "./runner"
import { LEAN_CHILD_PERMISSIONS, engineSessionPort } from "./stages"
import { writeQuestions, type Question } from "./state"
import { cleanupMatters, matter } from "./test-support"

// The engine session port runs through the real v2 adapter (engineV2Fetch and the
// transcript projection). Only the engine is scripted: it stores what it is sent
// and answers like the native v2 routes.

const DIRECTORY = "/matters/lean-port"

type NativeError = { type: string; message: string; status?: number }

function scriptedEngine(reply: { text: string; outcome: string; interrupted?: boolean; finalError?: boolean | NativeError; onIdlePoll?: () => void
  /** The engine's own refusal of the prompt, before any model call. */
  refusal?: { status: number; body: Record<string, unknown> }
  /** A turn that failed before any reply: the error is on its idle record only. */
  idleError?: NativeError }) {
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
      if (reply.refusal) return Response.json(reply.refusal.body, { status: reply.refusal.status })
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
          ...(reply.idleError ? [] : [{ id: "msg_000000000009aaaaaaaaaaaaaa", type: "assistant", agent: "expander", time: { created: 7, completed: 8 }, content: [{ type: "text", text: reply.text }],
            ...(reply.finalError ? { finish: "error", error: reply.finalError === true ? { type: "provider.transport", message: "Decode error", status: 200 } : reply.finalError } : {}) }]),
          { id: "idle_latest", type: "idle", outcome: reply.outcome, ...(reply.idleError ? { error: reply.idleError } : {}), time: { created: 9 } },
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
      .rejects.toThrow("The expander session ended succeeded: Decode error.")
  })

  test("a failed turn carries the native error's typed fields as its cause", async () => {
    const expired = { type: "provider.auth", message: "This API key has expired", status: 401 }
    const engine = scriptedEngine({ text: "", outcome: "failed", finalError: expired })
    const port = engineSessionPort({ engineURL: "http://engine.test", fetch: engine.fetcher, pollMs: 1 })
    const failure = await port.prompt({ directory: DIRECTORY, sessionID: "ses_child", agent: "expander", text: "Draft it." }).catch((error: unknown) => error)
    expect(failure).toBeInstanceOf(Error)
    // The provider's one-line message, never its payload; the typed fields stay in the cause.
    expect((failure as Error).message).toBe("The expander session ended failed: This API key has expired.")
    expect((failure as Error).cause).toEqual(expired)
  })

  test("a turn that failed before any reply carries its idle record's typed error as its cause", async () => {
    const unavailable = { type: "provider.no-route", message: "Model unavailable: test/scripted" }
    const engine = scriptedEngine({ text: "", outcome: "failed", idleError: unavailable })
    const port = engineSessionPort({ engineURL: "http://engine.test", fetch: engine.fetcher, pollMs: 1 })
    const failure = await port.prompt({ directory: DIRECTORY, sessionID: "ses_child", agent: "expander", text: "Draft it." }).catch((error: unknown) => error)
    expect((failure as Error).cause).toEqual(unavailable)
    expect(providerStopKind(failure)).toBe("model")
  })

  test("an engine refusal carries its typed body as its cause", async () => {
    const refusal = { error: "Finish installation setup before generating.", code: "setup_required" }
    const engine = scriptedEngine({ text: "", outcome: "succeeded", refusal: { status: 503, body: refusal } })
    const port = engineSessionPort({ engineURL: "http://engine.test", fetch: engine.fetcher, pollMs: 1 })
    const failure = await port.prompt({ directory: DIRECTORY, sessionID: "ses_child", agent: "expander", text: "Draft it." }).catch((error: unknown) => error)
    expect((failure as Error).message).toBe("The engine refused the expander prompt (HTTP 503).")
    expect((failure as Error).cause).toEqual(refusal)
    expect(engine.admitted).toHaveLength(0)
  })
})

describe("lean stage keys", () => {
  afterEach(cleanupMatters)

  test("an answer to a document's question re-keys the step that asked it and the document's later steps, no earlier one and no preparation step; a package-scope answer the document's steps only; a run-level answer re-keys them all", () => {
    const steps = ["instructions", "discrepancy", "intake", "research", "research-review", "acts", "spine", "d1.review", "d1.change-plan", "d1.change-prepare", "d1.changes", "d1.change-review"] as const
    const keys = (questions: Question[]) => {
      const directory = matter()
      writeLeanPlan(directory, [{ title: "Supply note", purpose: "Supply note", file: "supply-note.docx", parts: [{ heading: "Delivery", scope: "Delivery terms" }] }])
      writeQuestions(directory, questions)
      return Object.fromEntries(steps.map(step => [step, stageInputSha(directory, step)]))
    }
    const answered: Question = { id: "q-late", stage: "intake", question: "Should late delivery carry a remedy?", context: "The terms are silent.", status: "answered", answer: "No remedy.", kind: "judgment" }
    const none = keys([])
    const unchanged = (keyed: Record<string, string>, same: (step: typeof steps[number]) => boolean) => {
      for (const step of steps) expect([step, keyed[step] === none[step]]).toEqual([step, same(step)])
    }
    // The changes step applies its own answer: the reviews, the change plan
    // and the preparation before it never read it and keep their keys.
    unchanged(keys([{ ...answered, stage: "changes", document: "d1" }]), step => step !== "d1.changes" && step !== "d1.change-review")
    unchanged(keys([{ ...answered, stage: "change-plan", document: "d1" }]), step => !step.startsWith("d1.change"))
    unchanged(keys([answered]), () => false)
    // The plan step's package-scope answer carries the step's document, as the
    // runner records every document step's questions: it re-keys the document's
    // steps after its draft, not the act record or the spine. The plan step
    // settles it itself from questions.json; it directs no act of its own.
    const scope = keys([{ ...answered, id: `${PACKAGE_SCOPE_QUESTION}late`, stage: "draft", document: "d1", kind: "scope" }])
    unchanged(scope, step => !step.startsWith("d1."))
    // The gate's package-scope question carries no document and is asked once
    // the plan is saved: its answer re-keys no preparation step, the act record
    // included; the gate reads it itself and the document through its review.
    const gate = keys([{ ...answered, id: `${PACKAGE_SCOPE_QUESTION}gate`, stage: "gate", kind: "scope", mixed: true }])
    unchanged(gate, step => !step.startsWith("d1."))
  })
})
