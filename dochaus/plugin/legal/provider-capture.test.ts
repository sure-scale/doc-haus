import { expect, test } from "bun:test"
import { createHash } from "node:crypto"
import { chmod, lstat, mkdir, mkdtemp, readFile, readdir, realpath, rm, symlink } from "node:fs/promises"
import { tmpdir } from "node:os"
import path from "node:path"
import { createProviderCapture } from "./provider-capture"

const configDirectory = path.resolve(import.meta.dir, "../..")
const identity = { sessionID: "ses_observation", agent: "summarizer", kind: "primary", model: { providerID: "provider", id: "model", variant: "low" } }
const digest = (bytes: Uint8Array) => createHash("sha256").update(bytes).digest("hex")

async function fixture() {
  const base = await realpath(await mkdtemp(path.join(tmpdir(), "provider-capture-")))
  const directory = path.join(base, "matter")
  const captureDirectory = path.join(base, "capture")
  await mkdir(directory, { mode: 0o700 })
  await mkdir(captureDirectory, { mode: 0o700 })
  const capture = await createProviderCapture({ directory, configDirectory, captureDirectory })
  if (!capture) throw new Error("Capture fixture setup failed")
  return { base, directory, captureDirectory, capture }
}

async function completed(root: string, count = 1) {
  for (let attempt = 0; attempt < 200; attempt++) {
    const records = await Promise.all((await readdir(root)).map(async name => {
      const folder = path.join(root, name)
      if (!await Bun.file(path.join(folder, "receipt.json")).exists()) return undefined
      return { folder, request: await Bun.file(path.join(folder, "request.json")).json(), receipt: await Bun.file(path.join(folder, "receipt.json")).json() }
    }))
    const found = records.filter(record => record !== undefined)
    if (found.length === count) return found
    await Bun.sleep(10)
  }
  throw new Error("Capture did not finish")
}

test("final request and raw stream stay exact and private, linked to provisional observation", async () => {
  const f = await fixture()
  try {
    const body = '{"model":"model","messages":[{"role":"system","content":"neutral instruction"}],"tools":[],"max_tokens":100}'
    const request = new Request("https://provider.invalid/v1/chat?credential=hidden", { method: "POST", body, headers: { authorization: "hidden-auth", cookie: "hidden-cookie", "content-type": "application/json", "x-private": "hidden-private", "x-session-affinity": "ses_route", "x-opencode-session": "ses_route", "user-agent": "native-client", "idempotency-key": "hidden-idempotency" } })
    f.capture.request({ ...identity, request }, "msg_assistant_observation")
    let release: (() => void) | undefined
    const gate = new Promise<void>(resolve => { release = resolve })
    const chunks = [new TextEncoder().encode("data: {\"text\":\"sample\"}\n\n"), new Uint8Array([0, 255, 10]), new TextEncoder().encode("data: [DONE]\n\n")]
    let index = 0
    const event = { ...identity, request, response: new Response(new ReadableStream({ async pull(controller) { await gate; if (index === chunks.length) { controller.close(); return }; controller.enqueue(chunks[index++]) } }, { highWaterMark: 0 }), { status: 200, headers: { "content-type": "text/event-stream", "x-request-id": "req_observation", "set-cookie": "hidden-cookie" } }) }
    f.capture.response(event, "msg_assistant_observation")
    expect(index).toBe(0)
    release?.()
    const delivered = new Uint8Array(await event.response.arrayBuffer())
    expect(delivered).toEqual(Buffer.concat(chunks))
    expect(event.response.headers.get("set-cookie")).toBe("hidden-cookie")
    expect(request.bodyUsed).toBe(false)
    const records = await completed(f.captureDirectory, 2)
    const final = records.find(record => record.request.request_fidelity === "final")!
    const provisional = records.find(record => record.request.request_fidelity === "provisional")!
    expect(final.request.provisional_capture_id).toBe(provisional.request.capture_id)
    expect(final.request.request_started_at).toBe(provisional.request.request_started_at)
    expect(final.request.model).toEqual(identity.model)
    expect(final.request.assistant_message_id).toBe("msg_assistant_observation")
    expect(provisional.request.assistant_message_id).toBe("msg_assistant_observation")
    expect(final.request.source_commit).toMatch(/^[a-f0-9]{40}$/)
    expect(Object.keys(final.request.source_files_sha256)).toHaveLength(4)
    expect(final.request.endpoint).toEqual({ origin: "https://provider.invalid", path: "/v1/chat", query_present: true })
    expect(final.request.replay_supported).toBe(false)
    expect(final.request.headers).toEqual({ "content-type": "application/json", "x-session-affinity": "ses_route", "x-opencode-session": "ses_route", "user-agent": "native-client" })
    expect(final.request.agent_source_sha256).toMatch(/^[a-f0-9]{64}$/)
    expect(await Bun.file(path.join(final.folder, "request-body.bin")).text()).toBe(body)
    expect(await readFile(path.join(final.folder, "response-body.bin"))).toEqual(Buffer.from(delivered))
    expect(final.receipt.complete).toBe(true)
    expect(final.receipt.response_sha256).toBe(digest(delivered))
    expect(final.receipt.response_bytes).toBe(delivered.byteLength)
    expect(final.receipt.request_id).toBe("req_observation")
    expect(final.receipt.chunk_timings).toHaveLength(3)
    expect(final.receipt.first_byte_ms).toBeGreaterThanOrEqual(0)
    expect(final.receipt.elapsed_ms).toBeGreaterThanOrEqual(final.receipt.first_byte_ms)
    expect(JSON.stringify(records)).not.toContain("hidden")
    for (const record of records) {
      expect((await lstat(record.folder)).mode & 0o777).toBe(0o700)
      for (const name of await readdir(record.folder)) expect((await lstat(path.join(record.folder, name))).mode & 0o777).toBe(0o600)
    }
  } finally { await rm(f.base, { recursive: true, force: true }) }
})

test("cancel and upstream error keep original reasons and incomplete receipts", async () => {
  const f = await fixture()
  try {
    const reason = new Error("consumer cancellation")
    let cancelled: unknown
    const request = new Request("https://provider.invalid/v1/chat", { method: "POST", body: "{}" })
    const event = { ...identity, request, response: new Response(new ReadableStream<Uint8Array>({ cancel(value) { cancelled = value } }, { highWaterMark: 0 })) }
    f.capture.response(event)
    await event.response.body!.cancel(reason)
    expect(cancelled).toBe(reason)
    const failure = new Error("upstream failure")
    const failed = { ...identity, request: request.clone(), response: new Response(new ReadableStream<Uint8Array>({ pull(controller) { controller.error(failure) } }, { highWaterMark: 0 })) }
    f.capture.response(failed)
    await expect(failed.response.text()).rejects.toBe(failure)
    const records = await completed(f.captureDirectory, 2)
    expect(records.every(record => !record.receipt.complete)).toBe(true)
    expect(records.map(record => record.receipt.capture_error).sort()).toEqual(["cancelled", "upstream_error"])
  } finally { await rm(f.base, { recursive: true, force: true }) }
})

test("large provider requests and responses are captured completely while bytes arrive unchanged", async () => {
  const f = await fixture()
  try {
    const chunk = new Uint8Array(64 * 1024)
    const chunks = Array.from({ length: 48 }, (_, index) => chunk.map((_, offset) => (index + offset) % 251))
    let index = 0
    const body = "x".repeat(4 * 1024 * 1024 + 1)
    const event = { ...identity, request: new Request("https://provider.invalid/chat", { method: "POST", body }),
      response: new Response(new ReadableStream<Uint8Array>({ pull(controller) { if (index === chunks.length) { controller.close(); return }; controller.enqueue(chunks[index++]) } }, { highWaterMark: 0 })) }
    f.capture.response(event)
    const delivered = new Uint8Array(await event.response.arrayBuffer())
    expect(delivered).toEqual(new Uint8Array(Buffer.concat(chunks)))
    const [record] = await completed(f.captureDirectory)
    expect(record.receipt.complete).toBe(true)
    expect(record.receipt.capture_error).toBeUndefined()
    expect(record.receipt.response_bytes).toBe(delivered.byteLength)
    expect(await readFile(path.join(record.folder, "response-body.bin"))).toEqual(Buffer.from(delivered))
    expect(record.request.request_complete).toBe(true)
    expect(record.request.body_bytes).toBe(body.length)
    expect(await Bun.file(path.join(record.folder, "request-body.bin")).text()).toBe(body)
    expect(event.request.bodyUsed).toBe(false)
  } finally { await rm(f.base, { recursive: true, force: true }) }
})

test("sink disappearance does not change generation", async () => {
  const f = await fixture()
  try {
    await rm(f.captureDirectory, { recursive: true })
    const event = { ...identity, request: new Request("https://provider.invalid/chat", { method: "POST", body: "{}" }), response: new Response("unchanged") }
    f.capture.response(event)
    expect(await event.response.text()).toBe("unchanged")
    await Bun.sleep(20)
    await mkdir(f.captureDirectory, { mode: 0o700 })
    const next = { ...identity, request: new Request("https://provider.invalid/chat", { method: "POST", body: "{}" }), response: new Response("ok") }
    f.capture.response(next)
    expect(await next.response.text()).toBe("ok")
    expect((await completed(f.captureDirectory)).length).toBe(1)
  } finally { await rm(f.base, { recursive: true, force: true }) }
})

test("tracing defaults off, preserves earlier captures, and rejects unsafe explicit roots loudly", async () => {
  const f = await fixture()
  const previous = { capture: process.env.DOCHAUS_PROVIDER_CAPTURE_DIR, runtime: process.env.DOCHAUS_RUNTIME_DIR }
  try {
    process.env.DOCHAUS_PROVIDER_CAPTURE_DIR = "off"
    expect(await createProviderCapture({ directory: f.directory, configDirectory })).toBeUndefined()
    delete process.env.DOCHAUS_PROVIDER_CAPTURE_DIR
    delete process.env.DOCHAUS_RUNTIME_DIR
    expect(await createProviderCapture({ directory: f.directory, configDirectory })).toBeUndefined()
    process.env.DOCHAUS_RUNTIME_DIR = path.join(f.base, "runtime")
    const defaulted = await createProviderCapture({ directory: f.directory, configDirectory })
    expect(defaulted).toBeUndefined()
    process.env.DOCHAUS_PROVIDER_CAPTURE_DIR = f.captureDirectory
    expect((await createProviderCapture({ directory: f.directory, configDirectory }))?.root).toBe(f.captureDirectory)
    process.env.DOCHAUS_PROVIDER_CAPTURE_DIR = "off"
    expect(await createProviderCapture({ directory: f.directory, configDirectory })).toBeUndefined()
    expect((await lstat(f.captureDirectory)).isDirectory()).toBe(true)
    await expect(createProviderCapture({ directory: f.directory, configDirectory, captureDirectory: "relative" })).rejects.toThrow("absolute")
    await expect(createProviderCapture({ directory: f.directory, configDirectory, captureDirectory: f.directory })).rejects.toThrow("overlaps")
    await expect(createProviderCapture({ directory: f.directory, configDirectory, captureDirectory: path.join(f.directory, "nested") })).rejects.toThrow("overlaps")
    await expect(createProviderCapture({ directory: f.directory, configDirectory, captureDirectory: path.resolve(import.meta.dir, "../../../trace-in-repo") })).rejects.toThrow("not git-ignored")
    const loose = path.join(f.base, "capture-loose")
    await mkdir(loose, { mode: 0o755 })
    await chmod(loose, 0o755)
    await expect(createProviderCapture({ directory: f.directory, configDirectory, captureDirectory: loose })).rejects.toThrow("expected 700")
    expect(await readdir(loose)).toEqual([])
    // A misplaced root is rejected before any directory is created.
    await expect(createProviderCapture({ directory: f.directory, configDirectory, captureDirectory: path.join(f.directory, "new", "trace") })).rejects.toThrow("overlaps")
    expect(await Bun.file(path.join(f.directory, "new")).exists()).toBe(false)
    expect(await readdir(f.directory)).toEqual([])
    const link = path.join(f.base, "capture-link")
    await symlink(f.captureDirectory, link)
    await expect(createProviderCapture({ directory: f.directory, configDirectory, captureDirectory: link })).rejects.toThrow("not a symlink")
    expect(await readdir(f.captureDirectory)).toEqual([])
  } finally {
    for (const [name, value] of [["DOCHAUS_PROVIDER_CAPTURE_DIR", previous.capture], ["DOCHAUS_RUNTIME_DIR", previous.runtime]] as const) {
      if (value === undefined) delete process.env[name]
      if (value !== undefined) process.env[name] = value
    }
    await rm(path.resolve(import.meta.dir, "../../../trace-in-repo"), { recursive: true, force: true })
    await rm(f.base, { recursive: true, force: true })
  }
})

test("tool executions are recorded in full before and after legal hooks, with session, message and call identity", async () => {
  const f = await fixture()
  try {
    const output = Array.from({ length: 6000 }, (_, line) => `line ${line} ${"y".repeat(40)}`).join("\n")
    expect(output.length).toBeGreaterThan(50 * 1024)
    const base = { tool: "read-matter", sessionID: "ses_child_observation", agent: "drafter", messageID: "msg_tool_observation", id: "call_observation", input: { path: "notes.txt" } }
    const event = { ...base, status: "completed" as const, result: { content: output, metadata: { title: "notes" } } }
    f.capture.tool("raw", event)
    event.result = { content: "rewritten by a later hook", metadata: { title: "notes" } }
    f.capture.tool("final", event)
    f.capture.tool("raw", { ...base, id: "call_failed", status: "error", error: { message: "neutral failure", metadata: { reason: "missing" } } })
    const file = path.join(f.captureDirectory, "streams", "tool-calls", "ses_child_observation.jsonl")
    for (let attempt = 0; attempt < 200 && (await Bun.file(file).text().catch(() => "")).split("\n").filter(Boolean).length < 3; attempt++) await Bun.sleep(10)
    const lines = (await Bun.file(file).text()).split("\n").filter(Boolean).map(line => JSON.parse(line))
    expect(lines.map(line => [line.phase, line.call_id, line.status])).toEqual([["raw", "call_observation", "completed"], ["final", "call_observation", "completed"], ["raw", "call_failed", "error"]])
    expect(lines[0].result.content).toBe(output)
    expect(lines[1].result.content).toBe("rewritten by a later hook")
    expect(lines[0]).toMatchObject({ session_id: "ses_child_observation", message_id: "msg_tool_observation", tool: "read-matter", agent: "drafter", input: { path: "notes.txt" } })
    expect(lines[2].error).toEqual({ message: "neutral failure", metadata: { reason: "missing" } })
    for (const entry of [path.join(f.captureDirectory, "streams"), path.dirname(file)]) expect((await lstat(entry)).mode & 0o777).toBe(0o700)
    expect((await lstat(file)).mode & 0o777).toBe(0o600)
  } finally { await rm(f.base, { recursive: true, force: true }) }
})

test("credentials and auth headers are redacted at write time across every trace stream", async () => {
  const f = await fixture()
  try {
    const secrets = {
      bearer: "Bearer abcdefghijklmnopqrstuvwxyz0123456789",
      adc: "ya29.a0AfB_byC1234567890abcdefghijklmnopqrstuvwxyz",
      apiKey: "AIzaSyA1234567890abcdefghijklmnopqrstuv",
      providerKey: "sk-proj-abcdefghijklmnopqrstuvwxyz012345",
      refresh: "1//0abcdefghijklmnopqrstuvwxyz0123456789",
      privateKey: "-----BEGIN PRIVATE KEY-----\nMIIEvQIBADANBgkqhkiG9w0BAQEFAASC\n-----END PRIVATE KEY-----",
      lowerBearer: "bearer abcdefghijklmnop0123456789",
      aws: "AKIAABCDEFGHIJKLMNOP",
      oauthClient: "GOCSPX-abcdefghijklmnopqrstuv12",
      jwt: "eyJhbGciOiJIUzI1NiJ9.eyJzdWIiOiJuZXV0cmFsIn0.c2lnbmF0dXJlLXZhbHVl",
      github: "ghp_abcdefghijklmnopqrstuvwxyz0123456789",
    }
    const body = JSON.stringify({ model: "model", messages: [{ role: "user", content: `Keep the clause. ${secrets.bearer} ${secrets.adc}` }] })
    const request = new Request(`https://provider.invalid/v1/chat?key=${secrets.apiKey}`, { method: "POST", body, headers: { authorization: secrets.bearer, "x-goog-api-key": secrets.apiKey, "content-type": "application/json" } })
    const event = { ...identity, request, response: new Response("data: {\"text\":\"ok\"}\n\n") }
    f.capture.response(event)
    await event.response.text()
    f.capture.tool("raw", { tool: "fetch", sessionID: "ses_redaction", agent: "drafter", messageID: "msg_redaction", id: "call_redaction",
      input: { api_key: "plain-key-value", accessToken: "camel-token-value", client_secret: "snake-secret-value", token: "generic-token-value", secret: "generic-secret-value", headers: { Authorization: "Basic dXNlcjpwYXNz" }, url: "https://example.invalid" },
      status: "completed", result: { content: `Fetched. ${secrets.providerKey} ${secrets.privateKey} ${secrets.refresh} ${secrets.lowerBearer} ${secrets.aws} ${secrets.oauthClient} ${secrets.jwt} ${secrets.github} The bearer instruments require a basic understanding of the agreement.`, metadata: { title: "fetch", bytes: new Uint8Array([1, 2, 3]) } } })
    const cyclic: Record<string, unknown> = { title: "cyclic" }
    cyclic.self = cyclic
    expect(() => f.capture.tool("final", { tool: "fetch", sessionID: "ses_redaction", agent: "drafter", messageID: "msg_redaction", id: "call_cyclic", input: {}, status: "completed", result: { content: "x", metadata: cyclic } })).not.toThrow()
    const [record] = await completed(f.captureDirectory)
    expect(record.request.credential_redactions).toBe(2)
    expect(record.request.replay_supported).toBe(false)
    const tools = path.join(f.captureDirectory, "streams", "tool-calls", "ses_redaction.jsonl")
    for (let attempt = 0; attempt < 200 && !await Bun.file(tools).exists(); attempt++) await Bun.sleep(10)
    const files = [path.join(record.folder, "request.json"), path.join(record.folder, "request-body.bin"), path.join(record.folder, "receipt.json"), tools]
    for (let attempt = 0; attempt < 200 && !(await Bun.file(tools).text()).includes("call_cyclic"); attempt++) await Bun.sleep(10)
    const toolLines = (await Bun.file(tools).text()).split("\n").filter(Boolean).map(line => JSON.parse(line))
    expect(toolLines.find(line => line.call_id === "call_cyclic")).toMatchObject({ phase: "final", session_id: "ses_redaction", tool: "fetch" })
    expect(toolLines.find(line => line.call_id === "call_cyclic").serialization_error).toBeString()
    expect(toolLines.find(line => line.call_id === "call_redaction").result.metadata.bytes).toEqual({ bytes_base64: "AQID" })
    const written = (await Promise.all(files.map(file => Bun.file(file).text()))).join("\n")
    for (const secret of [...Object.values(secrets), "plain-key-value", "camel-token-value", "snake-secret-value", "generic-token-value", "generic-secret-value", "dXNlcjpwYXNz"]) expect(written).not.toContain(secret)
    expect(written).toContain("The bearer instruments require a basic understanding of the agreement.")
    expect(written).toContain("[REDACTED_CREDENTIAL]")
    expect(written).toContain("Keep the clause.")
    expect(written).toContain("Fetched.")
    expect(JSON.parse((await Bun.file(record.folder + "/request-body.bin").text())).messages[0].content).toBe("Keep the clause. [REDACTED_CREDENTIAL] [REDACTED_CREDENTIAL]")
  } finally { await rm(f.base, { recursive: true, force: true }) }
})

test("a stalled sink never delays the consumed stream and marks the capture incomplete", async () => {
  const f = await fixture()
  try {
    const chunks = Array.from({ length: 160 }, (_, index) => new Uint8Array(64 * 1024).fill(index % 251))
    const source = () => {
      let index = 0
      return new ReadableStream<Uint8Array>({ pull(controller) { if (index === chunks.length) { controller.close(); return }; controller.enqueue(chunks[index++]) } }, { highWaterMark: 0 })
    }
    const baselineStarted = performance.now()
    await new Response(source()).arrayBuffer()
    const baseline = performance.now() - baselineStarted
    // The request body never ends, so the capture sink cannot open its files.
    const stall: { release?: () => void } = {}
    const body = new ReadableStream<Uint8Array>({ start(controller) { stall.release = () => { controller.enqueue(new TextEncoder().encode("{}")); controller.close() } } })
    const event = { ...identity, request: new Request("https://provider.invalid/chat", { method: "POST", body, duplex: "half" } as RequestInit), response: new Response(source()) }
    f.capture.response(event)
    const started = performance.now()
    const delivered = new Uint8Array(await event.response.arrayBuffer())
    const elapsed = performance.now() - started
    expect(delivered).toEqual(new Uint8Array(Buffer.concat(chunks)))
    expect(elapsed).toBeLessThan(baseline + 1000)
    // Delivery finished while the sink was still stalled: no receipt can exist yet.
    expect(await Promise.all((await readdir(f.captureDirectory)).map(name => Bun.file(path.join(f.captureDirectory, name, "receipt.json")).exists()))).not.toContain(true)
    stall.release?.()
    const [record] = await completed(f.captureDirectory)
    expect(record.receipt.complete).toBe(false)
    expect(record.receipt.capture_error).toBe("buffer_limit")
    expect(record.receipt.response_truncated).toBe(true)
    expect(record.receipt.observed_response_bytes).toBe(delivered.byteLength)
    const saved = await readFile(path.join(record.folder, "response-body.bin"))
    expect(saved.byteLength).toBe(8 * 1024 * 1024)
    expect(saved).toEqual(Buffer.from(delivered.subarray(0, saved.byteLength)))
    expect(record.receipt.response_sha256).toBe(digest(saved))
  } finally { await rm(f.base, { recursive: true, force: true }) }
})
