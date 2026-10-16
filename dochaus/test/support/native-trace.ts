// Explicitly invoked native event recorder for synthetic integration tests.
import path from "node:path"
import { mkdirSync, createWriteStream } from "node:fs"
import type { OpenCodeClient, V2Event } from "@opencode/client"

type Tokens = { input: number; output: number; reasoning: number; cache: { read: number; write: number } }

// Captures the whole engine stream (root session plus every subagent child
// session); callers use an isolated engine for one test run. Returns a stop
// handle: call stop() once the run finishes to close the subscription, flush the
// JSONL, and write a token-accounting summary next to it.
export function startTrace(client: OpenCodeClient, file: string) {
  mkdirSync(path.dirname(file), { recursive: true })
  const sink = createWriteStream(file, { flags: "a" })
  const controller = new AbortController()
  const totals = new Map<string, Tokens>()
  const accounted = new Set<string>()
  let seq = 0
  const write = (record: Record<string, unknown>) => sink.write(JSON.stringify({ seq: seq++, at: Date.now(), ...record }) + "\n")

  const pump = (async () => {
    try {
      const res = client.event.subscribe({ signal: controller.signal })
      for await (const event of res) {
        if (controller.signal.aborted) break
        record(event, write, totals, accounted)
      }
    } catch (error) {
      if (!controller.signal.aborted) write({ type: "trace.error", error: String(error) })
    }
  })()

  return {
    async stop() {
      controller.abort()
      await pump
      const grand = [...totals.values()].reduce(
        (acc, t) => ({ input: acc.input + t.input, output: acc.output + t.output, reasoning: acc.reasoning + t.reasoning }),
        { input: 0, output: 0, reasoning: 0 },
      )
      write({ type: "trace.summary", sessions: totals.size, perSession: Object.fromEntries(totals), grand })
      await new Promise((resolve) => sink.end(resolve))
      return { file, grand, sessions: totals.size }
    },
  }
}

// Observe native execution events. Text deltas and bulk tool output stay in the
// paginated durable transcript; step endings carry authoritative token usage.
function record(event: V2Event, write: (r: Record<string, unknown>) => void, totals: Map<string, Tokens>, accounted: Set<string>) {
  if (
    event.type === "session.step.ended" || event.type === "session.step.failed" ||
    event.type === "session.compaction.ended" || event.type === "session.compaction.failed"
  ) {
    const tokens = event.data.tokens
    if (tokens && !accounted.has(event.id)) {
      accounted.add(event.id)
      const sid = event.data.sessionID
      const prev = totals.get(sid) ?? { input: 0, output: 0, reasoning: 0, cache: { read: 0, write: 0 } }
      totals.set(sid, {
        input: prev.input + tokens.input,
        output: prev.output + tokens.output,
        reasoning: prev.reasoning + tokens.reasoning,
        cache: { read: prev.cache.read + tokens.cache.read, write: prev.cache.write + tokens.cache.write },
      })
    }
  }
  if (
    event.type.startsWith("session.step.") ||
    event.type.startsWith("session.tool.") ||
    event.type.startsWith("permission.") ||
    (event.type.startsWith("session.compaction.") && event.type !== "session.compaction.delta") ||
    event.type === "session.status" ||
    event.type === "session.idle"
  ) {
    const data = event.data as Record<string, unknown>
    write({ type: event.type, eventID: event.id, data: { ...data, content: undefined, text: undefined, recent: undefined } })
  }
}
