import type { SessionContext } from "@opencode/plugin/promise/session"
import type { Message, Part } from "./policy-types"
import { MAX_STEPS_PROMPT } from "@opencode/core/session/runner/max-steps"
type History = Array<{ info: Message; parts: Part[] }>

// Modify the native context in place: keep provider reasoning and attachments,
// and apply policy removals to both sides of each tool call/result pair.
export function applyLegalContextChanges(event: SessionContext, before: History, after: History) {
  const originals = new Map(before.map(message => [message.info.id, message]))
  const changes = new Map(after.map(message => [message.info.id, message]))
  const removedTools = new Set(before.flatMap(message => message.parts.flatMap(part =>
    part.type === "tool" && !changes.get(message.info.id)?.parts.some(next => next.type === "tool" && next.callID === part.callID)
      ? [part.callID] : [],
  )))
  const outputs = new Map(after.flatMap(message => message.parts.flatMap((part): Array<readonly [string, { type: "text" | "error"; value: string }]> => {
    if (part.type !== "tool") return []
    const prior = originals.get(message.info.id)?.parts.find(item => item.type === "tool" && item.callID === part.callID)
    if (prior?.type !== "tool") return []
    if (part.state.status === "completed" && prior.state.status === "completed" && prior.state.output !== part.state.output)
      return [[part.callID, { type: "text" as const, value: part.state.output }] as const]
    if (part.state.status === "error" && prior.state.status === "error" && prior.state.error !== part.state.error)
      return [[part.callID, { type: "error" as const, value: part.state.error }] as const]
    return []
  })))
  event.messages = event.messages.flatMap(message => {
    const original = message.id ? originals.get(message.id) : undefined
    const changed = message.id ? changes.get(message.id) : undefined
    if (original && !changed) return []
    const used = new Set<string>()
    const content = message.content.flatMap((part): typeof message.content => {
      if (part.type === "tool-result") {
        if (removedTools.has(part.id)) return []
        const result = outputs.get(part.id)
        return [result ? { ...part, result } : part]
      }
      if (!original || !changed) return [part]
      const prior = original.parts.find(item => !used.has(item.id) && (
        part.type === "tool-call" ? item.type === "tool" && item.callID === part.id :
        (part.type === "text" || part.type === "reasoning") && item.type === part.type && item.text === part.text
      ))
      if (!prior) return [part]
      used.add(prior.id)
      const next = changed.parts.find(item => item.id === prior.id)
      if (!next) return []
      if (next.type === "text" && (part.type !== "text" || next.text !== part.text))
        return [{ type: "text", text: next.text }]
      return [part]
    })
    if (original && changed) content.push(...changed.parts.flatMap(part =>
      part.type === "text" && !original.parts.some(prior => prior.id === part.id)
        ? [{ type: "text" as const, text: part.text }] : [],
    ))
    return content.length ? [{ ...message, content }] : []
  })
  // A compaction checkpoint may replace the original user boundary. Rehydrate
  // only newly prepared evidence, never replay the original prompt or outputs.
  const represented = new Set(event.messages.map(message => message.id))
  for (const message of after) {
    if (message.info.role !== "user" || represented.has(message.info.id)) continue
    const prior = originals.get(message.info.id)
    if (!prior) continue
    const content = message.parts.flatMap(part => part.type === "text" &&
      !prior.parts.some(item => item.id === part.id) ? [{ type: "text" as const, text: part.text }] : [])
    if (content.length) event.messages.push({ id: message.info.id, role: "user", content })
  }
  // The native tool limit must not replace a stage's checked output with an
  // unparseable progress summary. Adapt only the actual transient control;
  // durable user messages and the native tool/step limits remain unchanged.
  if (!event.agent?.startsWith("lean-")) return
  event.messages = event.messages.map(message => {
    const part = message.content[0]
    if (message.id !== undefined || message.role !== "user" || message.content.length !== 1 ||
      part?.type !== "text" || part.text !== MAX_STEPS_PROMPT) return message
    return { ...message, content: [{ ...part, text:
      "MAXIMUM STEPS REACHED. Tools remain disabled; make no tool calls. Return this stage's required output format using only the inspected evidence and supplied context. The limit changes tool availability only, not the existing output contract or source checks. Follow that contract's requirements for remaining uncertainty, questions and limitations. Do not invent evidence or answers, omit essential requirements silently, or present unfinished work as complete.",
    }] }
  })
}
