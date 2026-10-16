import { isProjectedUserAdmission } from "../../lib/engine-v2-projection"

/** The stage result is the text of the latest completed, error-free assistant
 * message the stage agent produced for the exact admitted user turn. */
export function nativeStageResult(messages: unknown[], parentID: string, agent: string) {
  const latestUser = messages.filter(record).findLast(isProjectedUserAdmission)
  if (!latestUser || !record(latestUser.info) || latestUser.info.id !== parentID) return ""
  const last = messages.filter(record).findLast(message => record(message.info) &&
    message.info.role === "assistant" && message.info.parentID === parentID && message.info.agent === agent &&
    record(message.info.time) && message.info.time.completed && !message.info.error && Array.isArray(message.parts))
  return last ? (last.parts as unknown[]).filter(record).filter(part => part.type === "text" && typeof part.text === "string").map(part => part.text).join("\n") : ""
}

function record(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value)
}
