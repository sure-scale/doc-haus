import { createHash, randomUUID } from "node:crypto"
import { constants } from "node:fs"
import { mkdir, open, readFile } from "node:fs/promises"
import path from "node:path"
import { appendTrace, describeError, ensureTraceRoot, redactCredentials, resolveTraceRoot, sha256, traceSpaceAvailable, warnOnce } from "../../lib/trace-sink"

// Unwritten response bytes held in memory per response. The consumed stream is
// never paced by the sink: past this backlog, capture of that response stops and
// its receipt is marked incomplete (buffer_limit) while delivery continues.
const bufferLimit = 8 * 1024 * 1024
const timingLimit = 2048
const requestHeaders = [
  "content-type", "accept", "x-session-affinity", "x-session-id", "x-parent-session-id",
  "x-opencode-project", "x-opencode-session", "x-opencode-client", "user-agent",
  "anthropic-version", "anthropic-beta",
]
const responseHeaders = ["content-type", "x-request-id", "request-id"]

type CaptureIdentity = {
  sessionID: string
  agent: string
  kind: string
  model: { providerID: string; id: string; variant?: string }
  request: Request
}

type ToolObservation = {
  tool: string
  sessionID: string
  agent: string
  messageID: string
  id: string
  input: unknown
} & ({ status: "completed"; result: { output?: unknown; content?: unknown; metadata?: unknown } } | { status: "error"; error: { message: string; metadata?: unknown } })

// An observational sink only: no authority, retries, tools or provider rewrites.
// Disabled unless explicitly requested; a configured but unusable
// destination throws instead of silently disabling the trace.
export async function createProviderCapture(input: {
  directory: string
  configDirectory: string
  captureDirectory?: string
  workspaceDirectory?: string
}) {
  const destination = input.captureDirectory ?? resolveTraceRoot()
  if (!destination) return undefined
  const repository = path.resolve(import.meta.dir, "../../..")
  const root = await ensureTraceRoot(destination, {
    repository,
    protectedPaths: [input.directory, input.configDirectory, input.workspaceDirectory ?? process.env.WORKSPACE_ROOT],
  })
  // Provenance is best effort: a missing git binary or file never disables tracing.
  const sourceCommit = await (async () => {
    const source = Bun.spawn(["git", "-C", repository, "rev-parse", "HEAD"], { stdout: "pipe", stderr: "ignore" })
    const revision = (await new Response(source.stdout).text()).trim()
    return await source.exited === 0 && /^[a-f0-9]{40}$/.test(revision) ? revision : null
  })().catch(() => null)
  const sourceFiles = Object.fromEntries(await Promise.all([
    ["dochaus/plugin/legal/provider-capture.ts", import.meta.filename],
    ["dochaus/plugin/legal-v2.ts", path.resolve(import.meta.dir, "../legal-v2.ts")],
    ["dochaus/plugin/legal/v2-model-policy.ts", path.join(import.meta.dir, "v2-model-policy.ts")],
    ["config/opencode.json", path.join(input.configDirectory, "opencode.json")],
  ].map(async ([name, file]) => [name, await readFile(file).then(sha256, () => null)])))
  const provisional = new WeakMap<Request, { id: string; startedAt: string; started: number }>()
  const start = (event: CaptureIdentity, fidelity: "final" | "provisional", messageID: string | undefined) => {
    const id = randomUUID()
    const previous = provisional.get(event.request)
    const startedAt = previous?.startedAt ?? new Date().toISOString()
    const started = previous?.started ?? performance.now()
    if (fidelity === "provisional") provisional.set(event.request, { id, startedAt, started })
    const folder = path.join(root, id)
    const ready = (async () => {
      if (!await traceSpaceAvailable(root)) throw new Error("Trace capture suspended for low disk space")
      await mkdir(folder, { mode: 0o700 })
      const url = new URL(event.request.url)
      const body = redactBody(new Uint8Array(await event.request.clone().arrayBuffer()))
      const agentSourceHash = /^[a-zA-Z0-9_-]+$/.test(event.agent)
        ? await readFile(path.join(input.configDirectory, "agent", `${event.agent}.md`)).then(sha256).catch(() => undefined)
        : undefined
      await createFile(path.join(folder, "request-body.bin"), body.bytes)
      await createFile(path.join(folder, "request.json"), JSON.stringify({
        version: 2, capture_id: id, provisional_capture_id: fidelity === "final" ? previous?.id : undefined,
        session_id: event.sessionID, assistant_message_id: messageID, agent: event.agent, kind: event.kind,
        model: { providerID: event.model.providerID, id: event.model.id, variant: event.model.variant },
        source_commit: sourceCommit, source_files_sha256: sourceFiles, agent_source_sha256: agentSourceHash,
        endpoint: { origin: url.origin, path: url.pathname, query_present: !!url.search },
        method: event.request.method, headers: allowedHeaders(event.request.headers, requestHeaders),
        body_sha256: sha256(body.bytes), body_bytes: body.bytes.byteLength, request_complete: true,
        credential_redactions: body.redactions,
        request_fidelity: fidelity, replay_supported: fidelity === "final" && body.redactions === 0 && !url.search && !url.username && !url.password && event.request.method === "POST" && event.request.headers.get("content-type")?.includes("application/json") === true,
        request_started_at: startedAt, timing_origin: previous || fidelity === "provisional" ? "request_hook" : "response_hook",
      }))
      return folder
    })()
    // Observe rejections immediately and loudly, including when a transport never responds.
    void ready.catch(error => warnOnce("capture", `[drafting-trace] PROVIDER CAPTURE FAILED in ${root}: ${describeError(error)}`))
    return { ready, startedAt, started }
  }
  return {
    root,
    request(event: CaptureIdentity, messageID?: string) {
      try {
        const capture = start(event, "provisional", messageID)
        void capture.ready.then(folder => createFile(path.join(folder, "receipt.json"), JSON.stringify({
          version: 2, request_started_at: capture.startedAt, complete: false,
          capture_error: "no_response_observed", request_fidelity: "provisional",
        }))).catch(() => undefined)
      } catch { /* Diagnostics cannot fail generation. */ }
    },
    response(event: CaptureIdentity & { response: Response }, messageID?: string) {
      try {
        const capture = start(event, "final", messageID)
        const responseStartedAt = new Date().toISOString()
        const original = event.response
        const hash = createHash("sha256")
        const timings: Array<{ offset_ms: number; bytes: number }> = []
        let pending = capture.ready.then(folder => open(path.join(folder, "response-body.bin"), constants.O_WRONLY | constants.O_CREAT | constants.O_EXCL | constants.O_NOFOLLOW, 0o600)).catch(() => { captureError = "sink_failure"; return undefined })
        let buffered = 0
        let saved = 0
        let observed = 0
        let firstByte: number | undefined
        let captureError: string | undefined
        let sinkFailed = false
        let ended = false
        const receive = (chunk: Uint8Array) => {
          firstByte ??= performance.now() - capture.started
          observed += chunk.byteLength
          if (timings.length < timingLimit) timings.push({ offset_ms: performance.now() - capture.started, bytes: chunk.byteLength })
          if (captureError) return
          if (buffered + chunk.byteLength > bufferLimit) {
            captureError = "buffer_limit"
            return
          }
          const bytes = chunk.slice()
          buffered += bytes.byteLength
          pending = pending.then(async file => {
            // After a write failure later chunks are not written, so the file never has gaps.
            // After buffer_limit the already-queued contiguous prefix is still written.
            if (!file || sinkFailed) return file
            try {
              // FileHandle.write can write a short prefix; retain exact bytes.
              for (let offset = 0; offset < bytes.byteLength;) {
                const written = (await file.write(bytes, offset, bytes.byteLength - offset)).bytesWritten
                if (!written) throw new Error("Capture write made no progress")
                hash.update(bytes.subarray(offset, offset + written))
                saved += written
                offset += written
              }
            } catch {
              sinkFailed = true
              captureError = "sink_failure"
            }
            return file
          }).finally(() => { buffered -= bytes.byteLength })
          void pending.catch(() => { captureError ??= "sink_failure" })
        }
        const finish = (complete: boolean, reason?: string) => {
          if (ended) return
          ended = true
          const elapsed = performance.now() - capture.started
          void (async () => {
            const file = await pending
            await file?.close().catch(() => { captureError ??= "sink_failure" })
            const folder = await capture.ready
            if (captureError) warnOnce("capture", `[drafting-trace] PROVIDER RESPONSE CAPTURE INCOMPLETE in ${folder}: ${captureError}`)
            await createFile(path.join(folder, "receipt.json"), JSON.stringify({
              version: 2, status: original.status,
              headers: allowedHeaders(original.headers, responseHeaders),
              request_id: original.headers.get("x-request-id") ?? original.headers.get("request-id") ?? undefined,
              request_started_at: capture.startedAt, response_started_at: responseStartedAt,
              first_byte_ms: firstByte, elapsed_ms: elapsed,
              response_bytes: saved, observed_response_bytes: observed, response_truncated: saved < observed, response_sha256: hash.digest("hex"),
              complete: complete && !captureError, capture_error: captureError ?? reason,
              chunk_timings: timings, chunk_timings_complete: timings.length < timingLimit,
            }))
          })().catch(() => undefined)
        }
        if (!original.body) { finish(true); return }
        const reader = original.body.getReader()
        event.response = new Response(new ReadableStream<Uint8Array>({
          async pull(controller) {
            try {
              const next = await reader.read()
              if (next.done) { controller.close(); finish(true); return }
              try { receive(next.value) } catch { captureError = "sink_failure" }
              controller.enqueue(next.value)
            } catch (error) { finish(false, "upstream_error"); controller.error(error) }
          },
          cancel(reason) { finish(false, "cancelled"); return reader.cancel(reason) },
        }, { highWaterMark: 0 }), { status: original.status, statusText: original.statusText, headers: original.headers })
      } catch { /* Diagnostics cannot fail generation. */ }
    },
    /**
     * Record a tool execution in full. `raw` runs before legal after-hooks and
     * holds the tool's own untruncated result; `final` runs after them and
     * holds the result handed to native truncation and storage.
     */
    tool(phase: "raw" | "final", event: ToolObservation) {
      try {
        void appendTrace(root, "tool-calls", event.sessionID, {
          phase, session_id: event.sessionID, message_id: event.messageID, call_id: event.id, tool: event.tool, agent: event.agent,
          status: event.status, input: event.input,
          ...(event.status === "completed"
            ? { result: event.result }
            : { error: { message: event.error.message, metadata: event.error.metadata } }),
        })
      } catch { /* Diagnostics cannot fail tool execution. */ }
    },
  }
}

function allowedHeaders(headers: Headers, names: string[]) {
  return Object.fromEntries(names.flatMap(name => headers.has(name) ? [[name, headers.get(name)]] : []))
}

async function createFile(file: string, content: string | Uint8Array) {
  const handle = await open(file, constants.O_WRONLY | constants.O_CREAT | constants.O_EXCL | constants.O_NOFOLLOW, 0o600)
  try { await handle.writeFile(content) } finally { await handle.close() }
}

// Credentials are redacted before the body reaches disk; any redaction marks
// the capture non-replayable because the stored bytes differ from the wire.
function redactBody(bytes: Uint8Array) {
  const text = (() => { try { return new TextDecoder("utf-8", { fatal: true }).decode(bytes) } catch { return undefined } })()
  if (text === undefined) return { bytes, redactions: 0 }
  const redacted = redactCredentials(text)
  return redacted.count === 0 ? { bytes, redactions: 0 } : { bytes: new TextEncoder().encode(redacted.text), redactions: redacted.count }
}
