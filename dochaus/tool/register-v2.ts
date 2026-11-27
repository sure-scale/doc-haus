import { observeStartup, startupObservation } from "../lib/startup-observation"
import { OpenCode } from "@opencode/client"
import type { Plugin } from "@opencode/plugin"
import { legalTools } from "./registry"
import { tool, type LegalToolContext, type LegalToolResult } from "./legacy-definition"

type Context = Parameters<Plugin.Plugin["setup"]>[0]

// The legal assistants thread helper findings into their next step, so `task`
// runs to completion unless the caller asks for background work.
const SUBAGENT_DESCRIPTION = [
  "Spawns an agent in a child session to work on the specified task.",
  "The output includes a sessionID you can pass back later to continue that specific conversation with the subagent.",
  "New child sessions start with fresh context, so include all relevant context and instructions when you don't pass a sessionID.",
  "Foreground (default) runs the subagent to completion and returns its final response.",
  "Background mode (background=true) launches it asynchronously and returns immediately; you are notified when it finishes.",
  "Use background only for independent work that can run while you continue elsewhere.",
].join("\n")

const NO_SUBAGENT_TEXT = "Subagent completed without a text response."

export async function registerLegalTools(context: Context) {
  const definitions = await observeStartup("tools.import", context.location.directory, () => Promise.all(legalTools.map(async ([name, load]) => [name, (await load()).default] as const)))
  startupObservation("tools.transform", "enter", context.location.directory, { tools: definitions.length })
  return context.tool.transform((editor) => {
    // Keep legal policy vocabulary while v2 owns child execution and recovery.
    const subagent = editor.get("subagent")
    if (!subagent) throw new Error("Native subagent execution is unavailable")
    editor.remove("subagent")
    editor.add({
      name: "task",
      description: SUBAGENT_DESCRIPTION,
      output: tool.schema.object({
        sessionID: tool.schema.string(),
        status: tool.schema.union([tool.schema.literal("completed"), tool.schema.literal("running")]),
        output: tool.schema.string(),
      }),
      input: tool.schema.object({
        description: tool.schema.string(),
        prompt: tool.schema.string(),
        subagent_type: tool.schema.string(),
        task_id: tool.schema.string().optional(),
        background: tool.schema.boolean().optional(),
      }),
      options: { codemode: false, permission: "subagent" },
      execute: async (input, runtime) => {
        const startedAt = Date.now()
        const result = await subagent.execute({
          agent: input.subagent_type,
          description: input.description,
          prompt: input.prompt,
          ...(input.task_id ? { sessionID: input.task_id } : {}),
          ...(input.background !== undefined ? { background: input.background } : {}),
        }, {
          ...runtime,
          progress: update => runtime.progress({ ...update, sessionId: update.sessionID, parentSessionId: runtime.sessionID }),
        })
        const sessionID = childSessionID(result)
        const metadata = normalizeToolMetadata({ ...result.metadata, sessionID, sessionId: sessionID, parentSessionId: runtime.sessionID })
        if (input.background === true) return { ...result, metadata }
        const text = await settleSubagent(sessionID, startedAt, runtime.signal)
        return {
          output: { sessionID, status: "completed" as const, output: text },
          content: `<subagent sessionID="${sessionID}" state="completed">\n${text}\n</subagent>`,
          metadata: { ...metadata, status: "completed" },
        }
      },
    })
    definitions.forEach(([name, definition]) => editor.add({
      name,
      description: definition.description,
      input: tool.schema.object(definition.args as never),
      options: { codemode: false },
      execute: async (input, runtime) => {
        const legacy: LegalToolContext = {
          directory: context.location.directory,
          worktree: context.location.project.directory,
          agent: runtime.agent,
          sessionID: runtime.sessionID,
          messageID: runtime.messageID,
          callID: runtime.id,
          abort: runtime.signal,
          metadata: (update) => { void runtime.progress(normalizeToolMetadata({ ...update.metadata, ...(update.title ? { title: update.title } : {}) })) },
          ask: (request) => assertPermission(context, runtime, request),
        }
        const result = await definition.execute(input as never, legacy) as LegalToolResult | undefined
        if (result === undefined) throw new Error(runtime.signal.aborted ? "Tool execution interrupted" : "Tool returned no result")
        if (typeof result === "string") return { content: result }
        return { content: result.output, metadata: normalizeToolMetadata({ ...result.metadata, ...(result.title ? { title: result.title } : {}) }) }
      },
    }))
  }).then(value => {
    startupObservation("tools.transform", "exit", context.location.directory, { tools: definitions.length })
    return value
  }, error => {
    startupObservation("tools.transform", "failed", context.location.directory)
    throw error
  })
}

function engineEndpoint() {
  return (process.env.OPENCODE_URL ?? `http://127.0.0.1:${process.env.DOC_HAUS_ENGINE_PORT ?? "4097"}`).replace(/\/$/, "")
}

function engineHeaders(): Record<string, string> {
  const password = process.env.OPENCODE_SERVER_PASSWORD
  return password ? { authorization: `Basic ${btoa(`opencode:${password}`)}` } : {}
}

export function childSessionID(result: { metadata?: Record<string, unknown>; output?: unknown }) {
  const fromMetadata = result.metadata?.sessionID
  if (typeof fromMetadata === "string") return fromMetadata
  const fromOutput = (result.output as { sessionID?: unknown } | undefined)?.sessionID
  if (typeof fromOutput === "string") return fromOutput
  throw new Error("Native subagent did not return a session ID")
}

/**
 * The native tool now always launches children in the background. Foreground
 * callers wait for the child to go idle, then read its final response so the
 * legal orchestrators can thread it into their next step.
 */
async function settleSubagent(sessionID: string, startedAt: number, signal: AbortSignal) {
  const client = OpenCode.make({ baseUrl: engineEndpoint(), headers: engineHeaders() })
  let onAbort: (() => void) | undefined
  const interrupted = new Promise<never>((_, reject) => {
    if (signal.aborted) return reject(new Error("Tool execution interrupted"))
    onAbort = () => reject(new Error("Tool execution interrupted"))
    signal.addEventListener("abort", onAbort, { once: true })
  })
  void interrupted.catch(() => {})
  try {
    // `wait` returns at once when the child has not been admitted yet, and a
    // follow-up child already carries the previous turn's idle marker. Only an
    // idle record created since this call proves that this turn completed, so
    // re-check a missed admission instead of reporting an earlier answer.
    for (let attempt = 0; attempt < 40; attempt++) {
      await Promise.race([client.experimental.session.wait({ sessionID }), interrupted])
      const page = await client.message.list({ sessionID, order: "desc", limit: 8 })
      const newest = page.data[0] as { type: string; time?: { created?: number } } | undefined
      if (newest?.type === "idle" && (newest.time?.created ?? 0) >= startedAt) return subagentReport(page.data, sessionID)
      // The child's job fiber registers its execution just after the tool
      // returns; give it a moment rather than spinning on an idle child.
      if (attempt < 39) await Promise.race([Bun.sleep(250), interrupted])
    }
    throw new Error("Subagent completion could not be confirmed")
  } catch (error) {
    if (signal.aborted) {
      // Cancelling the job as well as the session stops a settlement from
      // reporting this cancelled child to the interrupted parent.
      await Promise.allSettled([
        client.session.interrupt({ sessionID }),
        client.experimental.job.cancel({ id: sessionID }),
      ])
    }
    throw error
  } finally {
    if (onAbort) signal.removeEventListener("abort", onAbort)
  }
}

type SubagentMessage = {
  readonly type: string
  readonly outcome?: string
  readonly error?: { readonly message?: string }
  readonly content?: ReadonlyArray<{ readonly type: string; readonly text?: string }>
}

/** Final assistant text for the child's latest turn, or a failure for a turn that did not succeed. */
export function subagentReport(messages: ReadonlyArray<SubagentMessage>, sessionID?: string) {
  // Newest first. A stale idle marker would report an earlier turn's answer.
  const idle = messages[0]
  if (idle?.type !== "idle") throw new Error("Subagent completion could not be confirmed")
  const at = sessionID ? ` (sessionID: ${sessionID})` : ""
  if (idle.outcome === "interrupted") throw new Error(`Subagent cancelled${at}`)
  if (idle.outcome === "failed") throw new Error(`Subagent failed${at}: ${idle.error?.message ?? "unknown error"}`)
  // Only this turn's messages follow the new idle marker: stop at the previous
  // idle so a follow-up that wrote no text cannot report an earlier answer.
  let assistant: SubagentMessage | undefined
  for (const message of messages.slice(1)) {
    if (message.type === "idle") break
    if (message.type === "assistant") {
      assistant = message
      break
    }
  }
  return (
    assistant?.content
      ?.filter((part) => part.type === "text")
      .map((part) => part.text ?? "")
      .filter((text) => text.length > 0)
      .join("\n\n") || NO_SUBAGENT_TEXT
  )
}

export async function assertPermission(
  context: Context,
  runtime: { sessionID: string; agent: string; messageID: string; id: string; signal: AbortSignal },
  request: Parameters<LegalToolContext["ask"]>[0],
) {
  const base = engineEndpoint()
  const password = process.env.OPENCODE_SERVER_PASSWORD
  const headers = engineHeaders()
  const controller = new AbortController()
  const events = OpenCode.make({ baseUrl: base, headers }).event.subscribe({ signal: controller.signal })[Symbol.asyncIterator]()
  let onAbort: (() => void) | undefined
  const interrupted = new Promise<never>((_, reject) => {
    if (runtime.signal.aborted) return reject(new Error("Permission request interrupted"))
    onAbort = () => reject(new Error("Permission request interrupted"))
    runtime.signal.addEventListener("abort", onAbort, { once: true })
  })
  try {
    // The initial event confirms that the live subscription is installed before
    // the permission request can be answered. A missed reply must fail closed.
    const ready = await Promise.race([events.next(), interrupted])
    if (ready.done || ready.value.type !== "server.connected") throw new Error("Permission event stream unavailable")
    const id = `per_${crypto.randomUUID()}`
    const reply = (async () => {
      for (;;) {
        const next = await events.next()
        if (next.done) throw new Error("Permission event stream closed")
        if (next.value.type !== "permission.replied" || (next.value.data.requestID !== id || next.value.data.sessionID !== runtime.sessionID)) continue
        if (next.value.data.reply === "reject") throw new Error(`Permission ${request.permission} declined`)
        return
      }
    })()
    void reply.catch(() => {})
    const response = await fetch(`${base}/api/session/${encodeURIComponent(runtime.sessionID)}/permission`, {
      method: "POST",
      headers: {
        "content-type": "application/json",
        ...(password ? { authorization: `Basic ${btoa(`opencode:${password}`)}` } : {}),
      },
      body: JSON.stringify({
        id,
        action: request.permission,
        resources: request.patterns,
        save: request.always,
        metadata: request.metadata,
        source: { type: "tool", messageID: runtime.messageID, id: runtime.id },
        agent: runtime.agent,
      }),
      signal: runtime.signal,
    })
    if (!response.ok) throw new Error(`Permission request failed (${response.status})`)
    const payload = await response.json() as { data?: { id?: string; effect?: "allow" | "ask" | "deny" } }
    if (payload.data?.id !== id) throw new Error("Permission response did not match request")
    if (payload.data.effect === "allow") return
    if (payload.data.effect === "deny") throw new Error(`Permission ${request.permission} denied`)
    if (payload.data.effect !== "ask") throw new Error("Permission response was incomplete")
    await Promise.race([reply, interrupted])
  } finally {
    if (onAbort) runtime.signal.removeEventListener("abort", onAbort)
    controller.abort()
    await events.return?.()
  }
}

// The v1 HTTP boundary omitted optional undefined fields. Native plugin results
// enter the event store directly, where nested undefined is not valid JSON.
// Preserve the same transport representation, and reject nonserializable data.
export function normalizeToolMetadata(metadata: Record<string, unknown>): Record<string, unknown> {
  return JSON.parse(JSON.stringify(metadata))
}
