import type { Part } from "../../api/session-view"
import { getMessages } from "../../api/opencode/sessions"

export function isDraftingRecoveryMessage(message: { parts: Part[] }) {
  return message.parts.some(
    (part) => part.type === "text" && /<dochaus-recovery\b[^>]*run-id="[^"]+"[^>]*progress-token="[^"]+"/i.test(part.text),
  )
}

export function lastVisibleUserMessage(messages: Awaited<ReturnType<typeof getMessages>>) {
  return [...messages].reverse().find((message) => message.info.role === "user" && !isDraftingRecoveryMessage(message))
}

// A chat turn that printed tool-call syntax instead of executing it is the
// one turn the conversation recovers itself, through ingest's continuation route.
export function draftingTurnRecoveryMode(malformedProgressToken?: string): "malformed-tool" | "none" {
  return malformedProgressToken ? "malformed-tool" : "none"
}
