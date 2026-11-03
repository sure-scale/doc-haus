import { isProjectedUserAdmission } from "../../lib/engine-v2-projection"

/** The stage result is the text of the latest completed, error-free assistant
 * message the stage agent produced for the exact admitted user turn. A reply
 * the provider cut off at its output limit before any text has no answer to
 * check, and a retry would only repeat it: that is a runtime stop, not a reply. */
export function nativeStageResult(messages: unknown[], parentID: string, agent: string) {
  const latestUser = messages.filter(record).findLast(isProjectedUserAdmission)
  if (!latestUser || !record(latestUser.info) || latestUser.info.id !== parentID) return ""
  const last = messages.filter(record).findLast(message => record(message.info) &&
    message.info.role === "assistant" && message.info.parentID === parentID && message.info.agent === agent &&
    record(message.info.time) && message.info.time.completed && !message.info.error && Array.isArray(message.parts))
  if (!last || !record(last.info)) return ""
  const text = (last.parts as unknown[]).filter(record).filter(part => part.type === "text" && typeof part.text === "string").map(part => part.text as string)
  if (last.info.finish === "length" && !text.some(part => part.trim())) {
    const reasoning = record(last.info.tokens) && typeof last.info.tokens.reasoning === "number" ? last.info.tokens.reasoning : undefined
    throw new Error(`The model's reply reached its output limit before giving any answer${reasoning ? ` (${reasoning.toLocaleString("en-US")} reasoning tokens)` : ""}.`)
  }
  return text.join("\n")
}

function record(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value)
}
