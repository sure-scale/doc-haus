import { nativeReadEvidence } from "./native-read-evidence"
/** Stable document-workflow projection of the native engine transcript.
 * Transport and execution remain v2; this keeps document validation independent
 * of a UI SDK. Native control messages are not fabricated into user prompts.
 */
export type ProjectedSession = Record<string, unknown> & {
  id: string
  directory: string
  agent?: string
  parentID?: string
  model?: { providerID: string; id: string; modelID: string }
}

export type ProjectedMessage = {
  info: Record<string, unknown> & {
    id: string
    sessionID: string
    role: "user" | "assistant"
    agent?: string
    parentID?: string
    time: { created: number; completed?: number }
  }
  parts: Array<Record<string, unknown> & { type: string; id: string; sessionID: string; messageID: string }>
}

/** Native synthetic continuation rows carry instructions, never prompt admission. */
export function isProjectedUserAdmission(value: unknown): boolean {
  return record(value) && record(value.info) && value.info.role === "user" &&
    value.info.type !== "synthetic" &&
    !(value.info.type === undefined && Array.isArray(value.parts) && value.parts.length > 0 && value.parts.every(part =>
      record(part) && part.type === "text" && part.synthetic === true))
}

export function projectSession(value: unknown): ProjectedSession {
  if (!record(value) || typeof value.id !== "string" || !record(value.location) || typeof value.location.directory !== "string")
    throw new Error("Engine returned an invalid session location.")
  const model = record(value.model) && typeof value.model.providerID === "string" && typeof value.model.id === "string"
    ? { providerID: value.model.providerID, id: value.model.id, modelID: value.model.id }
    : undefined
  return {
    ...value, id: value.id, directory: value.location.directory,
    ...(typeof value.agent === "string" ? { agent: value.agent } : {}),
    ...(typeof value.parentID === "string" ? { parentID: value.parentID } : {}),
    model,
  }
}

export function projectSessionMessages(
  messages: readonly unknown[],
  session: { id: string; agent?: string; model?: { providerID: string; id?: string; modelID?: string } },
): ProjectedMessage[] {
  let parentID: string | undefined
  let agent = session.agent
  let model = session.model
  return messages.flatMap((message): ProjectedMessage[] => {
    if (!record(message) || typeof message.id !== "string" || !record(message.time))
      throw new Error("Engine returned an invalid transcript message.")
    if (message.type === "agent-switched") {
      if (typeof message.agent === "string") agent = message.agent
      return []
    }
    if (message.type === "model-switched") {
      if (record(message.model) && typeof message.model.providerID === "string" && typeof message.model.id === "string")
        model = { providerID: message.model.providerID, id: message.model.id }
      return []
    }
    if (message.type !== "user" && message.type !== "assistant" && message.type !== "synthetic" && message.type !== "compaction") return []
    const role = message.type === "assistant" || message.type === "compaction" ? "assistant" : "user"
    const id = message.id
    const identity = { sessionID: session.id, messageID: id }
    const time = { created: millis(message.time.created), ...(message.time.completed !== undefined ? { completed: millis(message.time.completed) } : {}) }
    const info = {
      ...message, id, sessionID: session.id, role, time,
      agent: typeof message.agent === "string" ? message.agent : agent,
      ...(role === "assistant" && parentID ? { parentID } : {}),
      ...(record(message.model) ? { providerID: message.model.providerID, modelID: message.model.id } : { model: model ? { providerID: model.providerID, modelID: model.id ?? model.modelID } : undefined }),
      ...(message.type === "compaction" ? { summary: true } : {}),
      ...(record(message.metadata) && record(message.metadata.dochaus_tools) ? { tools: message.metadata.dochaus_tools } : {}),
    } satisfies ProjectedMessage["info"]
    if (message.type === "user") parentID = id
    const contents = message.type === "assistant" ? message.content : [{ type: "text", text: message.type === "compaction" ? message.summary : message.text }]
    if (!Array.isArray(contents)) throw new Error("Engine returned invalid assistant content.")
    const parts = contents.flatMap((content, index): ProjectedMessage["parts"] => {
      if (!record(content) || typeof content.type !== "string") throw new Error("Engine returned invalid message content.")
      const providerID = record(message.model) ? message.model.providerID : model?.providerID
      const providerMetadataKey = providerID === "google-vertex" ? "vertex" : providerID === "google" ? "google" : undefined
      const part = { ...content, ...identity, type: content.type, id: typeof content.id === "string" ? content.id : `${id}_${index}`,
        ...(providerMetadataKey && record(content.providerState) ? { metadata: { ...(record(content.metadata) ? content.metadata : {}), [providerMetadataKey]: content.providerState } } : {}),
      }
      if (content.type !== "tool") return [{ ...part, ...(message.type === "synthetic" ? { synthetic: true } : {}) }]
      if (!record(content.state)) throw new Error("Engine returned invalid tool state.")
      const state = content.state
      const output = Array.isArray(state.content) ? state.content.filter(record).filter(item => item.type === "text").map(item => item.text).join("\n") : undefined
      const read = content.name === "read" && record(state.input) && typeof output === "string"
        ? nativeReadEvidence(state.input, output, record(state.metadata) ? state.metadata : undefined) : undefined
      return [{ ...part, tool: content.name, callID: content.id, state: {
        ...state, status: state.status === "streaming" ? "pending" : state.status,
        ...(output !== undefined ? { output: read?.output ?? output } : {}),
        ...(read ? { input: read.input } : {}),
        ...(record(state.error) ? { error: String(state.error.message ?? state.error.type ?? "Tool failed") } : {}),
        ...(record(content.time) ? { time: { start: millis(content.time.created), ...(content.time.completed !== undefined ? { end: millis(content.time.completed) } : {}) } } : {}),
      } }]
    })
    return [{ info, parts }]
  })
}

function millis(value: unknown) {
  if (typeof value === "number" && Number.isFinite(value)) return value
  if (typeof value === "string") {
    const parsed = Date.parse(value)
    if (Number.isFinite(parsed)) return parsed
  }
  throw new Error("Engine returned an invalid message timestamp.")
}

function record(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value)
}
