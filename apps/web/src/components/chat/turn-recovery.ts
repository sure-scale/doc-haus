import type { Part } from "../../api/session-view"
import { getMessages } from "../../api/opencode/sessions"
import { isDraftingRecoveryMessage } from "./drafting-recovery"
import { errorOf } from "./turn-model"

export const MAX_MALFORMED_TOOL_RECOVERIES = 2
const STALL_TIMEOUT_MS = 30 * 60_000

export function shouldAbortStalledTurn(input: {
  busy: boolean
  sessionID: string
  aborting: boolean
  pendingPermissions: number
  lastActivityAt: number
  now: number
}) {
  return (
    input.busy &&
    Boolean(input.sessionID) &&
    !input.aborting &&
    input.pendingPermissions === 0 &&
    input.now - input.lastActivityAt >= STALL_TIMEOUT_MS
  )
}

export function stalledAbortFailureMessage(error: unknown) {
  const detail = error instanceof Error && error.message ? ` ${error.message}` : ""
  return `The engine did not respond while stopping the stalled turn.${detail} Check the connection, then retry.`
}

export function malformedToolRecoveryToken(message: { info: { id: string }; parts: Part[] }) {
  const toolSyntax = /(?:<\|tool_call\|>|call:[a-z][\w-]*\s*\{)/i
  const printedToolSyntax = message.parts.some(
    (part) =>
      (part.type === "reasoning" || (part.type === "text" && !part.synthetic)) && toolSyntax.test(part.text),
  )
  const hasAnswer = message.parts.some(
    (part) =>
      part.type === "text" &&
      !part.synthetic &&
      /[\p{L}\p{N}]/u.test(part.text) &&
      !toolSyntax.test(part.text),
  )
  return printedToolSyntax && !hasAnswer ? `malformed:${message.info.id}` : undefined
}

// The progress token of a matter chat whose latest model turn printed tool-call
// syntax instead of calling the tool, while recoveries remain for the current
// prompt. Any chat agent qualifies; the durable per-prompt bound and the
// matter's continuation route are the limits.
export function malformedToolRecoveryDue(messages: Awaited<ReturnType<typeof getMessages>>, matterId?: string) {
  const last = messages.at(-1)
  if (!matterId || last?.info.role !== "assistant" || errorOf(last.info.error)) return undefined
  if (malformedToolRecoveryCount(messages) >= MAX_MALFORMED_TOOL_RECOVERIES) return undefined
  return malformedToolRecoveryToken(last)
}

// Count malformed-tool recovery admissions since the latest real user prompt.
// The count lives in durable session history, so reopening the conversation
// cannot reset the bounded retry fuse.
export function malformedToolRecoveryCount(messages: Awaited<ReturnType<typeof getMessages>>) {
  const boundary = messages
    .map((message) => message.info.role === "user" && !isDraftingRecoveryMessage(message))
    .lastIndexOf(true)
  return messages.slice(boundary + 1).filter(
    (message) =>
      message.info.role === "user" &&
      message.parts.some(
        (part) =>
          part.type === "text" &&
          /<dochaus-recovery\b[^>]*kind="malformed-tool"[^>]*run-id="[^"]+"[^>]*progress-token="[^"]+"/i.test(
            part.text,
          ),
      ),
  ).length
}
