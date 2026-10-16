import type { Part } from "../../api/session-view"
import { getMessages } from "../../api/opencode/sessions"
import type { Client } from "../../api/opencode/client"

export function childSessionIDs(msgs: { parts: Part[] }[]) {
  return [
    ...new Set(
      msgs.flatMap((m) =>
        m.parts.flatMap((p) =>
          p.type === "tool" && p.tool === "task"
            ? [(p.state as { metadata?: { sessionId?: string } }).metadata?.sessionId].filter(
                (id): id is string => Boolean(id),
              )
            : [],
        ),
      ),
    ),
  ]
}

// Stored task parts point at separate child sessions. Load the full descendant
// graph whenever reconstructing a conversation so a remount/reopen produces the
// same nested timeline as the live SSE view. Without this, the parent-only load
// replaced completed subagent steps as soon as navigation or an error remounted
// the panel.
export async function loadConversation(
  client: Client,
  sessionID: string,
  options: { descendantTimeoutMs?: number; descendantBudgetMs?: number; descendantLimit?: number } = {},
) {
  const msgs = await getMessages(client, sessionID)
  const seen = new Set([sessionID])
  let queue = childSessionIDs(msgs)
  const descendants: Awaited<ReturnType<typeof getMessages>> = []
  const deadline = Date.now() + (options.descendantBudgetMs ?? 8_000)
  const limit = options.descendantLimit ?? 256
  while (queue.length && seen.size <= limit && Date.now() < deadline) {
    const level = [...new Set(queue.filter((id) => !seen.has(id)))].slice(0, Math.max(0, limit - seen.size + 1))
    if (!level.length) break
    level.forEach((id) => seen.add(id))
    const timeout = Math.max(1, Math.min(options.descendantTimeoutMs ?? 1_500, deadline - Date.now()))
    const children = await Promise.all(
      level.map((id) => bounded(getMessages(client, id), timeout, "Child conversation").catch(() => [])),
    )
    descendants.push(...children.flat())
    queue = children.flatMap(childSessionIDs)
  }
  return { msgs, pool: [...msgs, ...descendants].flatMap((m) => m.parts) }
}

export function bounded<T>(promise: Promise<T>, milliseconds: number, label: string) {
  let timeout: ReturnType<typeof setTimeout> | undefined
  return Promise.race([
    promise,
    new Promise<T>((_, reject) => {
      timeout = setTimeout(() => reject(new Error(`${label} timed out.`)), milliseconds)
    }),
  ]).finally(() => {
    if (timeout) clearTimeout(timeout)
  })
}
