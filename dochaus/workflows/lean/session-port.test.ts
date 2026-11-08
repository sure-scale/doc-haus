import { expect, test } from "bun:test"
import { engineSessionPort, nativeSessionWaits } from "./session-port"

const directory = "/synthetic-matter"
const agent = "lean-drafting-notes"
const session = { id: "ses_writer", location: { directory }, agent, model: { providerID: "selected", id: "model" } }
const timeout = () => new DOMException("The operation timed out.", "TimeoutError")

function fixture(options: { promotionRace?: "succeeded" | "failed"; transcriptPages?: boolean; availabilityText?: boolean; availabilityStatus?: number; permanentResult?: Record<string, unknown>; invalidIdle?: boolean; availabilityFailures?: number; quotaResult?: boolean; waitLimitMs?: number; noWaitLimit?: boolean; retryDelayMs?: number; pollMs?: number; readbackFailures?: number; neverIdle?: boolean; pauseDuringWait?: AbortController; lostAck?: boolean; ackBodyReset?: boolean; malformedAck?: boolean; malformedResult?: boolean; invalidResult?: boolean; bodyReset?: boolean; queued?: boolean; missing?: boolean; wrongText?: boolean; readFailures?: number; bodyFailures?: number; idleFailures?: number; ownerFailures?: number; failed?: boolean; superseded?: boolean; foreign?: boolean; refused?: boolean; pauseOnAdmission?: AbortController; pauseOnReadback?: AbortController } = {}) {
  const observed = { posts: 0, interrupts: 0, results: 0, olderReads: 0, idleReads: 0, ownerReads: 0, inboxReads: 0, exactReads: 0, id: "" }
  let messages: Record<string, any>[] = [], pending: Record<string, any>[] = []
  const promote = () => {
    if (options.missing || !observed.id) return
    pending = []
    messages = [
      { id: observed.id, sessionID: session.id, type: "user", text: options.wrongText ? "Different assignment" : "Private synthetic assignment", time: { created: 1 } },
      { id: "msg_reply", type: "assistant", agent, time: { created: 2, completed: 3 }, finish: "stop", content: [{ type: "text", text: "Checked later by the ordinary checker." }] },
      ...(options.superseded ? [{ id: "msg_newer", type: "user", text: "A later assignment", time: { created: 4 } }] : []),
      { id: "msg_idle", type: "idle", outcome: options.failed ? "failed" : "succeeded", time: { created: 5 } },
    ]
  }
  const transport = (async (input, init) => {
    const url = new URL(String(input)), method = init?.method ?? "GET"
    if (url.pathname === "/api/session/ses_writer") {
      if (++observed.ownerReads <= (options.ownerFailures ?? 0)) throw timeout()
      return Response.json({ data: options.foreign ? { ...session, location: { directory: "/other-matter" } } : session })
    }
    if (url.pathname.endsWith("/prompt")) {
      observed.posts++
      const body = JSON.parse(String(init?.body)); observed.id = body.id
      if (options.refused) return Response.json({ error: { type: "provider.authentication", message: "Selected connection needs attention" } }, { status: 401 })
      pending = options.missing ? [] : [{ id: body.id, sessionID: session.id, type: "user", delivery: "queue", payload: { text: options.wrongText ? "Different assignment" : body.text } }]
      if (options.promotionRace) messages = [
        { id: "msg_old_user", type: "user", text: "Earlier request", time: { created: 0 } },
        { id: "msg_old_idle", type: "idle", outcome: options.promotionRace, time: { created: 1 } },
      ]
      else if (!options.queued) promote()
      if (options.pauseOnAdmission) { options.pauseOnAdmission.abort(); throw new DOMException("Paused while acknowledgement was in flight", "AbortError") }
      if (options.lostAck) throw timeout()
      if (options.ackBodyReset) return new Response(new ReadableStream({ start(controller) { controller.error(new TypeError("Socket reset while consuming response")) } }))
      if (options.malformedAck) return new Response("{broken")
      return Response.json({ data: { id: body.id, sessionID: session.id, type: "user", delivery: "queue", payload: { text: body.text } } })
    }
    if (url.pathname.endsWith("/interrupt")) {
      observed.interrupts++
      expect(init?.signal?.aborted).toBe(false)
      return Response.json({ interrupted: true })
    }
    if (url.pathname.endsWith("/inbox")) {
      observed.inboxReads++
      if (options.pauseOnReadback) { options.pauseOnReadback.abort(); throw new DOMException("Paused during readback", "AbortError") }
      if (observed.inboxReads <= (options.readbackFailures ?? 0)) throw timeout()
      return Response.json({ data: pending })
    }
    if (url.pathname.startsWith("/api/session/ses_writer/message/")) {
      observed.exactReads++
      const row = messages.find(row => row.id === decodeURIComponent(url.pathname.split("/").at(-1)!))
      return row ? Response.json({ data: row }) : new Response(null, { status: 404 })
    }
    if (url.pathname.endsWith("/message")) {
      if (url.searchParams.get("limit") === "1") {
        if (++observed.idleReads <= (options.idleFailures ?? 0)) throw timeout()
        if (options.promotionRace && observed.idleReads > 1) promote()
        if (options.invalidIdle) return Response.json({ data: "invalid" })
        if (!messages.length) promote()
        if (options.neverIdle) return Response.json({ data: messages.slice(-2, -1) })
        return Response.json({ data: messages.slice(-1) })
      }
      if (url.searchParams.has("cursor")) {
        observed.olderReads++
        return Response.json({ data: [
          { id: "msg_older_idle", type: "idle", outcome: "failed", time: { created: 0 } },
          { id: "msg_older_user", type: "user", text: "Older request", time: { created: 0 } },
        ], cursor: {} })
      }
      observed.results++
      if (options.promotionRace && observed.results === 1) messages.push(
        { id: observed.id, type: "user", text: "Private synthetic assignment", time: { created: 2 } },
        { id: "msg_intermediate", type: "assistant", agent, time: { created: 3, completed: 4 }, content: [{ type: "text", text: "Intermediate text while the native turn is still running." }] },
      )
      if (options.permanentResult) return Response.json(options.permanentResult, { status: 503 })
      if (options.quotaResult) return Response.json({ type: "provider.quota", message: "Selected account quota exhausted" }, { status: 503 })
      if (observed.results <= (options.availabilityFailures ?? 0)) return options.availabilityText
        ? new Response("<html>Temporarily unavailable</html>", { status: options.availabilityStatus ?? 503, headers: { "content-type": "text/html" } })
        : Response.json({ error: "Engine temporarily unavailable" }, { status: options.availabilityStatus ?? 503 })
      if (observed.results <= (options.readFailures ?? 0)) {
        if (options.pauseDuringWait) setTimeout(() => options.pauseDuringWait!.abort(), 5)
        throw timeout()
      }
      if (observed.results <= (options.bodyFailures ?? 0)) return new Response(new ReadableStream({ start(controller) { controller.error(options.bodyReset ? new TypeError("Socket reset while consuming response") : timeout()) } }))
      if (options.malformedResult) return new Response("{broken")
      if (options.invalidResult) return Response.json({ data: [{ id: "msg_invalid", type: "assistant", time: { created: 1 }, content: null }] })
      return Response.json({ data: messages.toReversed(), cursor: options.transcriptPages ? { next: "older" } : {} })
    }
    throw new Error(`Unexpected synthetic request: ${method} ${url.pathname}`)
  }) as typeof fetch
  const port = engineSessionPort({ engineURL: "http://synthetic.invalid", fetch: transport, pollMs: options.pollMs ?? 0,
    ...(options.noWaitLimit ? {} : { timeoutMs: options.waitLimitMs ?? 1_000 }), retryDelayMs: options.retryDelayMs ?? 1 })
  return { port, observed, run: () => port.prompt({ directory, sessionID: session.id, agent, text: "Private synthetic assignment" }) }
}

test("a lost prompt acknowledgement reconciles the exact promoted admission without another POST", async () => {
  const f = fixture({ lostAck: true })
  expect(await f.run()).toBe("Checked later by the ordinary checker.")
  expect(f.observed.posts).toBe(1)
  expect(f.observed.exactReads).toBeGreaterThan(0)
})

test("a queued exact admission also survives a lost acknowledgement without duplicate generation", async () => {
  const f = fixture({ lostAck: true, queued: true })
  expect(await f.run()).toBe("Checked later by the ordinary checker.")
  expect(f.observed.posts).toBe(1)
  expect(f.observed.inboxReads).toBeGreaterThan(0)
})

test("an unproven admission stops with phase and durable identity, without exposing the prompt", async () => {
  const f = fixture({ lostAck: true, missing: true })
  let error: any
  try { await f.run() } catch (caught) { error = caught }
  expect(error).toBeInstanceOf(Error)
  expect(error.message).toContain("prompt-admission")
  expect(error.message).toContain(session.id)
  expect(error.message).toContain(f.observed.id)
  expect(error.message).toContain("timeout=30000ms")
  expect(error.message).toContain("elapsed=")
  expect(error.message).toContain("deadline=")
  expect(error.message).not.toContain("Private synthetic assignment")
  expect(f.observed.posts).toBe(1)
})

test.each([false, true])("a different durable payload is never treated as this prompt (queued=%s)", async queued => {
  const f = fixture({ lostAck: true, wrongText: true, queued })
  await expect(f.run()).rejects.toThrow()
  expect(f.observed.posts).toBe(1)
  expect(f.observed.results).toBe(0)
})

test("a transient completed-result read retries GET only", async () => {
  const f = fixture({ readFailures: 1 })
  expect(await f.run()).toBe("Checked later by the ordinary checker.")
  expect(f.observed.results).toBe(2)
  expect(f.observed.posts).toBe(1)
})

test("an explicit diagnostic wait limit stops an outage with the last request's phase and exact turn", async () => {
  const f = fixture({ readFailures: 100, waitLimitMs: 40 })
  let error: any
  try { await f.run() } catch (caught) { error = caught }
  expect(error.message).toContain("explicit")
  expect(error.message).toContain("wait limit")
  expect(error.message).toContain(session.id)
  expect(error.message).toContain(f.observed.id)
  expect(error.message).not.toContain("Private synthetic assignment")
  expect(f.observed.results).toBeGreaterThan(2)
  expect(f.observed.results).toBeLessThan(15)
  expect(f.observed.posts).toBe(1)
  expect(error.cause).toMatchObject({ phase: "result-read", sessionID: session.id, messageID: f.observed.id, timeoutMs: 30_000 })
  expect(error.cause.elapsedMs).toBeGreaterThanOrEqual(0)
  expect(error.cause.deadlineAt).toBeGreaterThan(0)
  expect(error.cause.context).toBeUndefined()
})

test("reconciled admission is not successful execution or acceptance", async () => {
  const f = fixture({ lostAck: true, failed: true })
  await expect(f.run()).rejects.toThrow("ended failed")
  expect(f.observed.posts).toBe(1)
})

test("another matter remains outside the prompt and reconciliation boundaries", async () => {
  const f = fixture({ foreign: true })
  await expect(f.run()).rejects.toThrow()
  expect(f.observed.posts).toBe(0)
  expect(f.observed.inboxReads).toBe(0)
})

test("a superseding durable user turn cannot supply this result", async () => {
  const f = fixture({ superseded: true })
  await expect(f.run()).rejects.toThrow("superseded")
  expect(f.observed.posts).toBe(1)
})

test("a completed-result body timeout is also recovered by GET only", async () => {
  const f = fixture({ bodyFailures: 1 })
  expect(await f.run()).toBe("Checked later by the ordinary checker.")
  expect(f.observed.results).toBe(2)
  expect(f.observed.posts).toBe(1)
})

test("a lost acknowledgement body socket reset reconciles the committed prompt without another POST", async () => {
  const f = fixture({ ackBodyReset: true })
  expect(await f.run()).toBe("Checked later by the ordinary checker.")
  expect(f.observed.posts).toBe(1)
  expect(f.observed.exactReads).toBe(1)
})

test("a completed-result body socket reset retries only the read", async () => {
  const f = fixture({ bodyFailures: 1, bodyReset: true })
  expect(await f.run()).toBe("Checked later by the ordinary checker.")
  expect(f.observed.results).toBe(2)
  expect(f.observed.posts).toBe(1)
})

test("malformed acknowledgement JSON remains a syntax failure without reconciliation", async () => {
  const f = fixture({ malformedAck: true })
  await expect(f.run()).rejects.toBeInstanceOf(SyntaxError)
  expect(f.observed.posts).toBe(1)
  expect(f.observed.inboxReads).toBe(0)
  expect(f.observed.exactReads).toBe(0)
  expect(f.observed.results).toBe(0)
})

test("malformed result JSON remains a syntax failure without read retry", async () => {
  const f = fixture({ malformedResult: true })
  await expect(f.run()).rejects.toBeInstanceOf(SyntaxError)
  expect(f.observed.results).toBe(1)
  expect(f.observed.posts).toBe(1)
})

test("a native projection contract failure is never retried as a body transport error", async () => {
  const f = fixture({ invalidResult: true })
  await expect(f.run()).rejects.toThrow("invalid assistant content")
  expect(f.observed.results).toBe(1)
  expect(f.observed.posts).toBe(1)
})

test("idle observation reconnects after more than two failed reads", async () => {
  const f = fixture({ idleFailures: 4 })
  expect(await f.run()).toBe("Checked later by the ordinary checker.")
  expect(f.observed.idleReads).toBe(5)
  expect(f.observed.posts).toBe(1)
})

test("failure before ownership readback cannot manufacture durable admission", async () => {
  const f = fixture({ ownerFailures: 1, missing: true })
  await expect(f.run()).rejects.toThrow("prompt-admission")
  expect(f.observed.posts).toBe(0)
})

test("native refusal retains its typed cause and is never retried or reconciled as admission", async () => {
  const f = fixture({ refused: true })
  let error: any
  try { await f.run() } catch (caught) { error = caught }
  expect(error.message).toContain("HTTP 401")
  expect(error.cause).toMatchObject({ error: { type: "provider.authentication" } })
  expect(f.observed.posts).toBe(1)
  expect(f.observed.inboxReads).toBe(0)
})

test("an already paused run admits no prompt and performs no recovery", async () => {
  const f = fixture()
  await expect(f.port.prompt({ directory, sessionID: session.id, agent, text: "Private synthetic assignment", signal: AbortSignal.abort() })).rejects.toThrow("paused")
  expect(f.observed.posts).toBe(0)
  expect(f.observed.ownerReads).toBe(0)
  expect(f.observed.inboxReads).toBe(0)
})

test("a pause after durable admission interrupts the same native session with a fresh deadline", async () => {
  const pause = new AbortController(), f = fixture({ pauseOnAdmission: pause })
  await expect(f.port.prompt({ directory, sessionID: session.id, agent, text: "Private synthetic assignment", signal: pause.signal })).rejects.toThrow("paused")
  expect(f.observed.posts).toBe(1)
  expect(f.observed.interrupts).toBe(1)
  expect(f.observed.inboxReads).toBe(0)
})

test("a pause during uncertain-admission readback also interrupts without reusing its aborted signal", async () => {
  const pause = new AbortController(), f = fixture({ lostAck: true, pauseOnReadback: pause })
  await expect(f.port.prompt({ directory, sessionID: session.id, agent, text: "Private synthetic assignment", signal: pause.signal })).rejects.toThrow("paused")
  expect(f.observed.posts).toBe(1)
  expect(f.observed.interrupts).toBe(1)
  expect(f.observed.inboxReads).toBe(1)
})

test("default observation waits through repeated connection loss and reads the same completed turn", async () => {
  const f = fixture({ readFailures: 5, noWaitLimit: true })
  expect(await f.run()).toBe("Checked later by the ordinary checker.")
  expect(f.observed.results).toBe(6)
  expect(f.observed.posts).toBe(1)
})

test("uncertain prompt acknowledgement survives a long readback outage without another POST", async () => {
  const f = fixture({ lostAck: true, queued: true, readbackFailures: 5 })
  expect(await f.run()).toBe("Checked later by the ordinary checker.")
  expect(f.observed.inboxReads).toBe(6)
  expect(f.observed.posts).toBe(1)
})

test("pause cancels connection backoff promptly and interrupts the same native session", async () => {
  const pause = new AbortController(), f = fixture({ readFailures: 100, pauseDuringWait: pause, retryDelayMs: 200 })
  const started = performance.now()
  await expect(f.port.prompt({ directory, sessionID: session.id, agent, text: "Private synthetic assignment", signal: pause.signal })).rejects.toThrow("paused")
  expect(performance.now() - started).toBeLessThan(150)
  expect(f.observed.results).toBe(1)
  expect(f.observed.posts).toBe(1)
  expect(f.observed.interrupts).toBe(1)
})

test("explicit wait limit also bounds ordinary polling without a busy spin", async () => {
  const f = fixture({ neverIdle: true, pollMs: 5, waitLimitMs: 35 })
  await expect(f.run()).rejects.toThrow("wait limit")
  expect(f.observed.idleReads).toBeGreaterThan(1)
  expect(f.observed.idleReads).toBeLessThan(10)
  expect(f.observed.posts).toBe(1)
})

function creationFixture(mode: "before" | "after", mismatched = false, retryDelayMs = 1) {
  const posts: string[] = [], observed = { reads: 0 }
  let stored: Record<string, unknown> | undefined
  const input = { directory, parentID: "ses_owner", agent, title: "Focused synthetic writer", model: { providerID: "selected", modelID: "model" },
    permissions: [{ action: "*", resource: "*", effect: "deny" as const }], metadata: { assignment: "synthetic" } }
  const transport = (async (target, init) => {
    const url = new URL(String(target))
    if (url.pathname === "/api/dochaus/session" && init?.method === "POST") {
      const body = JSON.parse(String(init.body)); posts.push(body.id)
      if (mode === "before" && posts.length === 1) throw timeout()
      stored = { ...body, id: body.id, location: { directory }, ...(mismatched ? { agent: "other-agent" } : {}) }
      if (posts.length === 1) throw timeout()
      return Response.json({ data: stored })
    }
    if (url.pathname.startsWith("/api/session/") && (init?.method ?? "GET") === "GET") {
      if (++observed.reads <= 3) throw timeout()
      return stored ? Response.json({ data: stored }) : new Response(null, { status: 404 })
    }
    throw new Error("Unexpected synthetic creation request")
  }) as typeof fetch
  return { input, posts, observed, port: engineSessionPort({ engineURL: "http://synthetic.invalid", fetch: transport, retryDelayMs, timeoutMs: 1_000 }) }
}

test("creation reconnects after a lost committed response and returns the exact same native ID", async () => {
  const f = creationFixture("after")
  const id = await f.port.createSession(f.input)
  expect(id).toBe(f.posts[0]!)
  expect(id).toMatch(/^ses_/)
  expect(f.posts).toHaveLength(1)
  expect(f.observed.reads).toBe(4)
})

test("creation lost before native commit retries only the same idempotent native ID after readback absence", async () => {
  const f = creationFixture("before")
  const id = await f.port.createSession(f.input)
  expect(f.posts).toEqual([id, id])
  expect(id).toMatch(/^ses_/)
  expect(f.observed.reads).toBe(4)
})

test("creation recovery cannot adopt a same-ID record with another agent or permission selection", async () => {
  const f = creationFixture("after", true)
  await expect(f.port.createSession(f.input)).rejects.toThrow("match")
  expect(f.posts).toHaveLength(1)
})

test("malformed raw idle envelope fails clearly rather than waiting forever", async () => {
  const f = fixture({ invalidIdle: true, waitLimitMs: 30 })
  await expect(f.run()).rejects.toThrow("invalid native idle")
  expect(f.observed.idleReads).toBe(1)
})

test("temporary HTTP availability loss reconnects while provider quota remains terminal", async () => {
  const unavailable = fixture({ availabilityFailures: 4 })
  expect(await unavailable.run()).toBe("Checked later by the ordinary checker.")
  expect(unavailable.observed.results).toBe(5)
  expect(unavailable.observed.posts).toBe(1)
  const quota = fixture({ quotaResult: true })
  await expect(quota.run()).rejects.toThrow("HTTP 503")
  expect(quota.observed.results).toBe(1)
})

test.each([502, 503, 504])("temporary HTML HTTP %s waits for reconnection without another generation POST", async availabilityStatus => {
  const f = fixture({ availabilityFailures: 3, availabilityStatus, availabilityText: true })
  expect(await f.run()).toBe("Checked later by the ordinary checker.")
  expect(f.observed.results).toBe(4)
  expect(f.observed.posts).toBe(1)
})

test.each([{ type: "provider.auth" }, { type: "provider.quota" }, { code: "setup_required" }])("typed permanent HTTP 503 body remains terminal: %j", async permanentResult => {
  const f = fixture({ permanentResult })
  await expect(f.run()).rejects.toThrow("HTTP 503")
  expect(f.observed.results).toBe(1)
  expect(f.observed.posts).toBe(1)
})

test("waiting display is defensive process memory and clears when the same result is collected", async () => {
  const f = fixture({ readFailures: 1, retryDelayMs: 40, noWaitLimit: true })
  const completed = f.run()
  await Bun.sleep(5)
  const [waiting] = nativeSessionWaits(directory)
  expect(waiting).toMatchObject({ sessionID: session.id, phase: "result-read" })
  expect(waiting!.lastSuccessfulContact).toBeGreaterThan(0)
  expect(waiting!.nextRetry).toBeGreaterThan(Date.now())
  waiting!.nextRetry = 0
  expect(nativeSessionWaits(directory)[0]!.nextRetry).toBeGreaterThan(0)
  expect(await completed).toBe("Checked later by the ordinary checker.")
  expect(nativeSessionWaits(directory)).toEqual([])
})

test("a paused creation wait exposes its owner only transiently and sends no prompt", async () => {
  const f = creationFixture("after", false, 100), pause = new AbortController()
  const completed = f.port.createSession({ ...f.input, signal: pause.signal })
  await Bun.sleep(5)
  expect(nativeSessionWaits(directory)).toEqual([expect.objectContaining({ sessionID: f.posts[0], parentID: f.input.parentID, phase: "session-create-readback" })])
  pause.abort()
  await expect(completed).rejects.toThrow("paused")
  expect(f.posts).toHaveLength(1)
  expect(nativeSessionWaits(directory)).toEqual([])
})

test.each(["succeeded", "failed"] as const)("a previous %s idle cannot decide a current turn promoted during transcript collection", async promotionRace => {
  const f = fixture({ promotionRace })
  expect(await f.run()).toBe("Checked later by the ordinary checker.")
  expect(f.observed.results).toBe(2)
  expect(f.observed.posts).toBe(1)
})

test("a prior successful idle cannot accept intermediate prose before the current turn fails", async () => {
  const f = fixture({ promotionRace: "succeeded", failed: true })
  await expect(f.run()).rejects.toThrow("ended failed")
  expect(f.observed.results).toBe(2)
  expect(f.observed.posts).toBe(1)
})

test("older transcript pages cannot overwrite the terminal head of the current native snapshot", async () => {
  const f = fixture({ transcriptPages: true })
  expect(await f.run()).toBe("Checked later by the ordinary checker.")
  expect(f.observed.olderReads).toBe(1)
  expect(f.observed.posts).toBe(1)
})
