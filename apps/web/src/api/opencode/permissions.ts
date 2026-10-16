import { matterClient } from "./client"

// A pending permission request from the engine. The edit tools (word-integration,
// tracked-changes, redline) call ctx.ask before mutating or proposing against a
// document; the engine parks the tool call and emits a permission.asked event,
// then resumes (or fails) it when we reply. metadata carries what the tool is
// about to do (document, find/replace or clause/replacement) for display.
export type PermissionRequest = {
  id: string
  sessionID: string
  permission: string
  patterns: string[]
  metadata: Record<string, unknown>
  always: string[]
  tool?: { messageID: string; callID: string }
}

export type PermissionReply = "once" | "always" | "reject"

// The permission events the engine emits on the same SSE stream as everything
// else. The v1 SDK's Event union does not model them, so the subscriber narrows
// raw events through this type (see ChatPanel.onEvent).
export type PermissionEvent =
  | { type: "permission.asked"; properties: PermissionRequest }
  | { type: "permission.replied"; properties: { sessionID: string; requestID: string; reply: PermissionReply } }

// Permission requests and replies use the native client in the active matter.
export async function listPermissions(directory: string) {
  const res = await matterClient(directory).permission.request.list({ location: { directory } })
  return res.data.map((request) => ({
    id: request.id,
    sessionID: request.sessionID,
    permission: request.action,
    patterns: request.resources,
    metadata: request.metadata ?? {},
    always: request.save ?? [],
    tool: request.source && { messageID: request.source.messageID, callID: request.source.id },
  }))
}

export async function replyPermission(directory: string, requestID: string, reply: PermissionReply) {
  const client = matterClient(directory)
  const request = (await client.permission.request.list({ location: { directory } })).data.find((item) => item.id === requestID)
  if (!request) throw new Error("The permission request is no longer pending.")
  await client.permission.reply({ sessionID: request.sessionID, requestID, decision: reply })
}
