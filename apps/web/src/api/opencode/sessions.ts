import type { SessionMessageInfo } from "@opencode/client"
import { projectMessages } from "../session-view"
import { boundedRequest, clientDirectories, location, type Client } from "./client"
import { getConfig, modelSpec } from "./config"

export async function listAgents(client: Client) {
  const res = await client.agent.list({ location: location(client) })
  return res.data.map((agent) => ({ name: agent.id, description: agent.description ?? "", mode: agent.mode, hidden: agent.hidden }))
}

export async function disposeInstance(_client: Client) {
  // V2 owns location lifecycle; closing a page does not evict other sessions.
}

// Session control is the web flow's control plane. A lost engine response must
// not pin the composer, router, credential probe, or cleanup path indefinitely.
export async function createSession(client: Client, title: string, timeoutMs = 10_000) {
  return boundedRequest(client.session.create({ title, location: location(client) }), timeoutMs, "Session creation")
}

// Every session the engine holds for this matter (scoped by the client's
// directory header). Subagent runs carry a parentID; top-level chats do not.
export async function listSessions(client: Client) {
  const sessions: Awaited<ReturnType<typeof client.session.list>>["data"] = []
  let cursor: string | undefined
  do {
    const page = await client.session.list({ limit: 200, ...(cursor ? { cursor } : { order: "desc", directory: clientDirectories.get(client) }) })
    sessions.push(...page.data)
    cursor = page.cursor.next ?? undefined
  } while (cursor)
  return sessions
}

// Delete a session and its messages from the engine. Irreversible.
export async function deleteSession(client: Client, sessionID: string, timeoutMs = 10_000) {
  return boundedRequest(client.session.remove({ sessionID }), timeoutMs, "Session deletion")
}

export async function abortSession(client: Client, sessionID: string, timeoutMs = 10_000) {
  return boundedRequest(client.session.interrupt({ sessionID }), timeoutMs, "Session abort")
}

// Resolves once the session has no running execution.
export async function waitForSession(client: Client, sessionID: string, timeoutMs = 10_000) {
  return boundedRequest(client.session.wait({ sessionID }), timeoutMs, "Session settle")
}

// Settled messages for one session, each as { info, parts }, used to replay a
// past conversation back into the chat panel.
export async function getMessages(client: Client, sessionID: string) {
  const messages: SessionMessageInfo[] = []
  let cursor: string | undefined
  do {
    const page = await client.message.list({ sessionID, limit: 200, ...(cursor ? { cursor } : { order: "asc" }) })
    messages.push(...page.data)
    cursor = page.cursor.next ?? undefined
  } while (cursor)
  return projectMessages(sessionID, messages)
}

export async function getSessionStatus(client: Client, sessionID: string): Promise<{ type: "busy" | "idle" | "retry" }> {
  const active = await client.session.active()
  return active[sessionID] ? { type: "busy" as const } : { type: "idle" as const }
}

// Fire a prompt to a named agent. Resolves when the assistant turn completes;
// live progress arrives separately through subscribeEvents.
export async function sendPrompt(client: Client, sessionID: string, agent: string, text: string) {
  // The selected product model is part of admission. A failed settings read
  // cannot silently select a different engine default.
  const model = modelSpec(await boundedRequest(getConfig(), 3_000, "Model configuration"))
  if (!model) throw new Error("Select a valid generation model before sending a message.")
  await client.session.switchAgent({ sessionID, agent })
  await client.session.switchModel({ sessionID, model: { providerID: model.providerID, id: model.modelID } })
  const admitted = await client.session.prompt({ sessionID, text, delivery: "queue" })
  await client.session.wait({ sessionID })
  const settled = (await client.message.list({ sessionID, order: "desc", limit: 2 })).data
  const terminal = settled[0]
  if (terminal?.type !== "idle") throw new Error("Session completion could not be confirmed.")
  if (terminal.outcome !== "succeeded") {
    const failed = settled.find(message => "error" in message && message.error)
    throw new Error(failed && "error" in failed ? failed.error!.message :
      terminal.outcome === "interrupted" ? "The session was interrupted." : "The session failed.")
  }
  return admitted
}

// Revert a user message and everything after it, rolling the session back to
// just before that turn. Used by edit/retry: revert, then send the new prompt
// so the assistant answer is regenerated from the edited question.
export async function revertMessage(client: Client, sessionID: string, messageID: string) {
  await client.session.revert.stage({ sessionID, messageID, files: true })
  return client.session.revert.commit({ sessionID })
}
