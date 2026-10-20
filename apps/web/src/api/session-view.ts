import type { SessionMessageInfo, SessionMessageAssistantTool, V2Event } from "@opencode/client"

// Keep the chat's display model local to the browser. The engine remains the
// source of truth: these values are projected from its public v2 transcript.
export type Part =
  | { type: "text"; id: string; sessionID: string; messageID: string; text: string; synthetic?: boolean }
  | { type: "reasoning"; id: string; sessionID: string; messageID: string; text: string; time: { start: number; end?: number } }
  | {
      type: "tool"
      id: string
      sessionID: string
      messageID: string
      tool: string
      state:
        | { status: "pending"; input: Record<string, unknown> }
        | { status: "running"; input: Record<string, unknown>; title: string; metadata?: Record<string, unknown> }
        | { status: "completed"; input: Record<string, unknown>; output: string; title: string; metadata?: Record<string, unknown> }
        | { status: "error"; input: Record<string, unknown>; error: string; title: string; metadata?: Record<string, unknown> }
    }
  | { type: "compaction"; id: string; sessionID: string; messageID: string; tail_start_id?: string }

export type Message = {
  info: {
    id: string
    sessionID: string
    role: "user" | "assistant"
    agent?: string
    parentID?: string
    summary?: boolean
    finish?: string
    error?: { name: string; data?: { message?: string } }
    time: { created: number; completed?: number }
  }
  parts: Part[]
}

function toolPart(tool: SessionMessageAssistantTool, sessionID: string, messageID: string): Part {
  const input = tool.state.status === "streaming" ? {} : tool.state.input
  const metadata = tool.state.status === "streaming" ? undefined : tool.state.metadata
  const title = typeof metadata?.title === "string" ? metadata.title : tool.name
  const base = { type: "tool" as const, id: tool.id, sessionID, messageID, tool: tool.name }
  if (tool.state.status === "completed")
    return {
      ...base,
      state: {
        status: "completed",
        input,
        output: tool.state.content.filter((item) => item.type === "text").map((item) => item.text).join("\n"),
        title,
        metadata,
      },
    }
  if (tool.state.status === "error")
    return { ...base, state: { status: "error", input, error: tool.state.error.message, title, metadata } }
  if (tool.state.status === "running")
    return { ...base, state: { status: "running", input, title, metadata } }
  return { ...base, state: { status: "pending", input } }
}

export function contentParts(sessionID: string, messageID: string, content: Extract<SessionMessageInfo, { type: "assistant" }>["content"]): Part[] {
  return content.map((part, index) => {
    if (part.type === "tool") return toolPart(part, sessionID, messageID)
    if (part.type === "reasoning")
      return {
        type: "reasoning" as const,
        id: `${messageID}:reasoning:${index}`,
        sessionID,
        messageID,
        text: part.text,
        time: { start: part.time?.created ?? 0, end: part.time?.completed },
      }
    return { type: "text" as const, id: `${messageID}:text:${index}`, sessionID, messageID, text: part.text }
  })
}

export function projectMessages(sessionID: string, messages: SessionMessageInfo[]): Message[] {
  let latestUser: string | undefined
  let currentAgent: string | undefined
  return messages.flatMap((message): Message[] => {
    if (message.type === "agent-switched") {
      currentAgent = message.agent
      return []
    }
    if (message.type === "user" || message.type === "synthetic") {
      latestUser = message.id
      return [{
        info: { id: message.id, sessionID, role: "user" as const, agent: currentAgent, time: message.time },
        parts: [{ type: "text" as const, id: `${message.id}:text`, sessionID, messageID: message.id, text: message.text, synthetic: message.type === "synthetic" }],
      }]
    }
    if (message.type !== "assistant") return []
    return [{
      info: {
        id: message.id,
        sessionID,
        role: "assistant" as const,
        agent: message.agent,
        parentID: latestUser,
        finish: message.finish,
        error: message.error && { name: message.error.type, data: { message: message.error.message } },
        time: message.time,
      },
      parts: contentParts(sessionID, message.id, message.content),
    }]
  })
}

export type Event =
  | { type: "permission.asked"; properties: { id: string; sessionID: string; permission: string; patterns: string[]; metadata: Record<string, unknown>; always: string[]; tool?: { messageID: string; callID: string } } }
  | { type: "permission.replied"; properties: { sessionID: string; requestID: string; reply: "once" | "always" | "reject" } }
  | { type: "message.updated"; properties: { info: { id: string; sessionID: string; role: "assistant" } } }
  | { type: "message.part.updated"; properties: { part: Part } }
  | { type: "session.status"; properties: { sessionID: string; status: { type: "busy" | "idle" | "retry" } } }
  | { type: "session.idle"; properties: { sessionID: string } }
  | { type: "session.deleted"; properties: { sessionID: string } }
  | { type: "session.error"; properties: { sessionID: string; error: { name: string; data: { message: string } } } }

export function projectEvent(event: V2Event): Event[] {
  if (event.type === "permission.asked")
    return [{ type: event.type, properties: {
      id: event.data.id,
      sessionID: event.data.sessionID,
      permission: event.data.action,
      patterns: event.data.resources,
      metadata: event.data.metadata ?? {},
      always: event.data.save ?? [],
      tool: event.data.source && { messageID: event.data.source.messageID, callID: event.data.source.id },
    } }]
  if (event.type === "permission.replied") return [{ type: event.type, properties: event.data }]
  if (event.type === "session.step.started")
    return [{ type: "message.updated", properties: { info: { id: event.data.assistantMessageID, sessionID: event.data.sessionID, role: "assistant" } } }]
  if (event.type === "session.status") return [{ type: event.type, properties: event.data }]
  if (event.type === "session.deleted") return [{ type: event.type, properties: event.data }]
  if (event.type === "session.idle") return [{ type: event.type, properties: event.data }]
  // The engine ends a turn with its execution outcome, not session.idle. A
  // success settles the chat; a failure or interruption rejects sendPrompt.
  if (event.type === "session.execution.succeeded") return [{ type: "session.idle", properties: event.data }]
  if (event.type === "session.step.failed")
    return [{ type: "session.error", properties: { sessionID: event.data.sessionID, error: { name: event.data.error.type, data: { message: event.data.error.message } } } }]
  return []
}
