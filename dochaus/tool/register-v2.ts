import { observeStartup, startupObservation } from "../lib/startup-observation"
import { OpenCode } from "@opencode/client"
import type { Plugin } from "@opencode/plugin"
import { legalTools } from "./registry"
import { tool, type LegalToolContext, type LegalToolResult } from "./legacy-definition"

type Context = Parameters<Plugin.Plugin["setup"]>[0]

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
      description: subagent.description,
      output: subagent.output,
      input: tool.schema.object({
        description: tool.schema.string(),
        prompt: tool.schema.string(),
        subagent_type: tool.schema.string(),
        task_id: tool.schema.string().optional(),
        background: tool.schema.boolean().optional(),
      }),
      options: { codemode: false, permission: "subagent" },
      execute: async (input, runtime) => {
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
        return { ...result, metadata: normalizeToolMetadata({ ...result.metadata, sessionId: result.metadata?.sessionID, parentSessionId: runtime.sessionID }) }
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

export async function assertPermission(
  context: Context,
  runtime: { sessionID: string; agent: string; messageID: string; id: string; signal: AbortSignal },
  request: Parameters<LegalToolContext["ask"]>[0],
) {
  const base = (process.env.OPENCODE_URL ?? `http://127.0.0.1:${process.env.DOC_HAUS_ENGINE_PORT ?? "4097"}`).replace(/\/$/, "")
  const password = process.env.OPENCODE_SERVER_PASSWORD
  const headers: Record<string, string> = password ? { authorization: `Basic ${btoa(`opencode:${password}`)}` } : {}
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
