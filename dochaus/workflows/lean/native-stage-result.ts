import { isProjectedUserAdmission } from "../../lib/engine-v2-projection"

/** A reply the provider cut off at its output limit before any text or tool call. */
export class OutputLimitError extends Error {}

/** The stage result is the text of the latest completed, error-free assistant
 * message the stage agent produced for the exact admitted user turn. A reply
 * the provider cut off at its output limit before any text or tool call has no
 * answer to check, and a retry would only repeat it: that is a runtime stop,
 * not a reply. A cut-off tool call goes to the checker's retry as before. */
export function nativeStageResult(messages: unknown[], parentID: string, agent: string) {
  const latestUser = messages.filter(record).findLast(isProjectedUserAdmission)
  if (!latestUser || !record(latestUser.info) || latestUser.info.id !== parentID) return ""
  const last = messages.filter(record).findLast(message => record(message.info) &&
    message.info.role === "assistant" && message.info.parentID === parentID && message.info.agent === agent &&
    record(message.info.time) && message.info.time.completed && !message.info.error && Array.isArray(message.parts))
  if (!last || !record(last.info)) return ""
  const parts = (last.parts as unknown[]).filter(record)
  const text = parts.filter(part => part.type === "text" && typeof part.text === "string").map(part => part.text as string)
  if (last.info.finish === "length" && !text.some(part => part.trim()) && !parts.some(part => part.type === "tool")) {
    const tokens = record(last.info.tokens) ? last.info.tokens : {}
    const spent = (count: unknown, kind: string) => typeof count === "number" && count > 0 ? ` (${count.toLocaleString("en-US")} ${kind} tokens)` : undefined
    throw new OutputLimitError(`The model's reply reached its output limit before giving any answer${spent(tokens.reasoning, "reasoning") ?? spent(tokens.output, "output") ?? ""}.`)
  }
  return text.join("\n")
}

function record(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value)
}
