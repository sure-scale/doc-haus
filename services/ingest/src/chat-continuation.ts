import { createHash } from "node:crypto"
import { diagnosticErrorMetadata } from "../../../dochaus/lib/diagnostic-metadata"
import { engineV2Fetch } from "../../../dochaus/lib/engine-v2"
import { openDb } from "./db"
import { boundedEngineRequest, engineURL, serializeMatter, type EngineRequestOptions } from "./matter-session"

// Chat recovery for a model turn that printed tool-call syntax instead of
// calling the tool. The ingest service owns both admission and acknowledgement:
// one server-derived continuation per failed turn, claimed durably beside the
// matter, admitted to the native inbox under a stable message ID, and
// acknowledged only once the exact claimed payload is visible in session
// history. A transport acknowledgement alone is not proof.

export type ChatContinuationInput = {
  kind: "malformed-tool"
  runId: string
  progressToken: string
  sessionID: string
  messageID: string
  text: string
  agent?: string
  model?: { providerID: string; modelID: string }
}

export type ChatContinuationResult = {
  accepted: boolean
  attempted: boolean
  existing: boolean
  executing: boolean
  persisted: boolean
  pending?: boolean
  uncertain: boolean
  exhausted?: boolean
  ownerState?: SessionExecutionState
}

export type ChatContinuationOptions = EngineRequestOptions & { requireIdleOwner?: boolean }

export type SessionExecutionState = "idle" | "busy" | "retry" | "unavailable"

export class ChatContinuationConflictError extends Error {}

type Claim = {
  message_id: string
  session_id: string
  kind: string
  run_id: string
  progress_token: string
  payload_hash: string
  status: "claimed" | "accepted" | "observed" | "uncertain"
  attempts: number
}

const MAX_SINGLE_ID_TRANSPORT_ATTEMPTS = 2
const MAX_MALFORMED_TOOL_RECOVERIES_PER_TURN = 2
const MALFORMED_TOOL_CONTINUATION =
  "The previous response ended before its intended continuation completed. Continue once from the durable workflow checkpoint without repeating completed work."
const IDENTITY = /^[A-Za-z0-9_-]{1,512}$/
const OWNER_EXECUTION_TAIL_LIMIT = 16

export function malformedToolContinuationInput(sessionID: string, failedMessageID: string): ChatContinuationInput {
  if (!IDENTITY.test(sessionID) || !IDENTITY.test(failedMessageID))
    throw new ChatContinuationConflictError("The malformed-turn continuation identity is invalid.")
  const progressToken = `malformed:${failedMessageID}`
  return {
    kind: "malformed-tool",
    runId: sessionID,
    progressToken,
    sessionID,
    messageID: `msg_recovery_${failedMessageID}`,
    text:
      `<dochaus-recovery kind="malformed-tool" run-id="${sessionID}" progress-token="${progressToken}" />\n` +
      MALFORMED_TOOL_CONTINUATION,
  }
}

export function dispatchChatContinuation(
  directory: string,
  input: ChatContinuationInput,
  options: ChatContinuationOptions = {},
): Promise<ChatContinuationResult> {
  return serializeMatter(directory, async () => {
    const hash = payloadHash(input)
    const db = openDb(directory)
    try {
      const prior = db.query("SELECT * FROM continuation_dispatches WHERE message_id = ?").get(input.messageID) as Claim | null
      if (prior) {
        requireMatchingClaim(prior, input, hash)
        return await reconcileExistingClaim(directory, prior, input, options)
      }
      // An existing exact claim is always safe to observe, even after its
      // admitted turn made the owner busy. Only a new claim waits for an idle
      // owner, which closes the check-then-insert race.
      if (options.requireIdleOwner) {
        const ownerState = await readSessionExecutionState(directory, input.sessionID, options)
        if (ownerState !== "idle")
          return {
            accepted: false, attempted: false, existing: false,
            executing: ownerState === "busy" || ownerState === "retry",
            persisted: false, uncertain: ownerState === "unavailable", ownerState,
          }
      }
      await requireCurrentContinuation(directory, input, options)
      const now = Date.now()
      const inserted = db.query(`
        INSERT INTO continuation_dispatches (
          message_id, session_id, kind, run_id, progress_token, payload_hash, status, attempts, created_at, updated_at
        ) VALUES (?, ?, ?, ?, ?, ?, 'claimed', 1, ?, ?)
        ON CONFLICT(message_id) DO NOTHING
        RETURNING message_id
      `).get(input.messageID, input.sessionID, input.kind, input.runId, input.progressToken, hash, now, now)
      if (!inserted) {
        const concurrent = db.query("SELECT * FROM continuation_dispatches WHERE message_id = ?").get(input.messageID) as Claim | null
        if (!concurrent) throw new ChatContinuationConflictError("The continuation claim could not be read.")
        requireMatchingClaim(concurrent, input, hash)
        return await reconcileExistingClaim(directory, concurrent, input, options)
      }
      const response = await dispatchToEngine(directory, input, options).catch(() => undefined)
      if (response?.status === 409)
        throw new ChatContinuationConflictError("The engine rejected a conflicting continuation identity or session selection.")
      setClaimStatus(db, input.messageID, response?.ok ? "accepted" : "uncertain")
      // V2 responds only after durable admission; queued work may not yet run.
      if (response?.ok)
        return { accepted: true, attempted: true, existing: false, executing: false, persisted: true, pending: true, uncertain: false }
      return { accepted: false, attempted: true, existing: false, executing: false, persisted: false, uncertain: true }
    } finally {
      db.close()
    }
  })
}

/** Native execution state of one session: queued inbox input, an active status,
 * a pending permission or a running subagent task is work in progress; an
 * unreadable or absent owner is unknown. */
export async function readSessionExecutionState(
  directory: string,
  sessionID: string,
  options: EngineRequestOptions = {},
): Promise<SessionExecutionState> {
  const session = `/api/session/${encodeURIComponent(sessionID)}`
  const [active, owner, inbox, permissions, recent] = await Promise.all([
    engineJson(directory, "/api/session/active", options),
    engineJson(directory, session, options),
    engineJson(directory, `${session}/inbox`, options),
    engineJson(directory, "/api/permission/request", options),
    engineJson(directory, `${session}/message?limit=${OWNER_EXECUTION_TAIL_LIMIT}`, options),
  ])
  if (
    !isRecord(active) ||
    !isRecord(owner) ||
    owner.id !== sessionID ||
    !Array.isArray(inbox) ||
    !Array.isArray(permissions) ||
    !Array.isArray(recent)
  ) return "unavailable"
  if (inbox.length) return "busy"
  const current = active[sessionID]
  if (current !== undefined) {
    if (!isRecord(current) || typeof current.type !== "string") return "unavailable"
    if (current.type === "busy" || current.type === "retry") return current.type
    if (current.type !== "idle") return "unavailable"
  }
  const waiting = new Set(permissions.flatMap(permission =>
    isRecord(permission) && typeof permission.sessionID === "string" ? [permission.sessionID] : []))
  if (waiting.has(sessionID)) return "busy"
  // A running subagent task keeps the owner's turn open while its child works.
  const running = recent.some(message =>
    isRecord(message) &&
    isRecord(message.info) &&
    message.info.role === "assistant" &&
    Array.isArray(message.parts) &&
    message.parts.some(part => {
      if (!isRecord(part) || part.type !== "tool" || part.tool !== "task" || !isRecord(part.state)) return false
      if (part.state.status !== "pending" && part.state.status !== "running") return false
      const child = isRecord(part.state.metadata) ? part.state.metadata.sessionId : undefined
      if (typeof child !== "string") return false
      const status = active[child]
      return waiting.has(child) || (isRecord(status) && (status.type === "busy" || status.type === "retry"))
    }))
  return running ? "busy" : "idle"
}

async function reconcileExistingClaim(
  directory: string,
  claim: Claim,
  input: ChatContinuationInput,
  options: ChatContinuationOptions,
): Promise<ChatContinuationResult> {
  const observed = await observeEngineContinuation(directory, claim, options)
  const db = openDb(directory)
  try {
    if (observed.persisted) {
      setClaimStatus(db, claim.message_id, "observed")
      return {
        accepted: true, attempted: true, existing: true,
        executing: observed.executing, persisted: true,
        pending: observed.pending, uncertain: false,
        ...(!observed.executing && !observed.pending ? { exhausted: true } : {}),
      }
    }
    if (observed.messageIDPresent)
      throw new ChatContinuationConflictError("The engine admission conflicts with this legal continuation.")
    // An exact retry is reconciled by the native durable inbox even when the
    // first HTTP response was lost. Never invent a replacement identity here.
    await requireCurrentContinuation(directory, input, options)
    if (claim.attempts >= MAX_SINGLE_ID_TRANSPORT_ATTEMPTS)
      return {
        accepted: claim.status === "accepted" || claim.status === "observed",
        attempted: false, existing: true, executing: observed.executing,
        persisted: false, uncertain: true, exhausted: true,
      }
    db.query("UPDATE continuation_dispatches SET attempts = attempts + 1, updated_at = ? WHERE message_id = ?")
      .run(Date.now(), claim.message_id)
    const response = await dispatchToEngine(directory, input, options).catch(() => undefined)
    if (response?.status === 409)
      throw new ChatContinuationConflictError("The engine rejected a conflicting continuation identity or session selection.")
    if (response?.ok) setClaimStatus(db, claim.message_id, "accepted")
    return {
      accepted: Boolean(response?.ok), attempted: true, existing: true,
      executing: observed.executing, persisted: Boolean(response?.ok),
      pending: Boolean(response?.ok), uncertain: !response?.ok,
    }
  } finally {
    db.close()
  }
}

async function observeEngineContinuation(directory: string, claim: Claim, options: EngineRequestOptions) {
  const session = `/api/session/${encodeURIComponent(claim.session_id)}`
  const [message, inbox, statuses] = await Promise.all([
    engineJson(directory, `${session}/message/${encodeURIComponent(claim.message_id)}`, options),
    engineJson(directory, `${session}/inbox`, options),
    engineJson(directory, "/api/session/active", options),
  ])
  const queued = Array.isArray(inbox) ? inbox.find(item => isRecord(item) && item.id === claim.message_id && item.sessionID === claim.session_id) : undefined
  const delivered = isRecord(message) && isRecord(message.info) && message.info.role === "user" && message.info.id === claim.message_id
  const text = isRecord(queued) && isRecord(queued.payload) && typeof queued.payload.text === "string"
    ? queued.payload.text
    : delivered && Array.isArray(message.parts)
      ? message.parts.filter(isRecord).filter(part => part.type === "text").map(part => typeof part.text === "string" ? part.text : "").join("")
      : undefined
  const current = isRecord(statuses) ? statuses[claim.session_id] : undefined
  return {
    persisted: text !== undefined && payloadHash({ ...input(claim), text }) === claim.payload_hash,
    messageIDPresent: Boolean(queued || delivered),
    // Pending inbox entries are work even before the execution drain starts.
    pending: Boolean(queued) || !Array.isArray(inbox) || !isRecord(statuses),
    executing: isRecord(current) && current.type === "busy",
  }
}

async function dispatchToEngine(directory: string, input: ChatContinuationInput, options: EngineRequestOptions) {
  const owner = input.agent ? { agent: input.agent, model: input.model } : await engineOwner(directory, input.sessionID, options)
  if (!owner) return
  return boundedEngineRequest(
    engineV2Fetch(options.fetch),
    `${engineURL(options.engineURL)}/api/session/${encodeURIComponent(input.sessionID)}/prompt`,
    {
      method: "POST",
      headers: { "Content-Type": "application/json", "x-opencode-directory": directory },
      body: JSON.stringify({
        messageID: input.messageID,
        agent: owner.agent,
        ...(owner.model ? { model: owner.model } : {}),
        parts: [{ type: "text", text: input.text }],
      }),
    },
  )
}

async function engineOwner(directory: string, sessionID: string, options: EngineRequestOptions) {
  const body = await engineJson(directory, `/api/session/${encodeURIComponent(sessionID)}`, options)
  if (!isRecord(body) || typeof body.agent !== "string" || !isRecord(body.model)) return
  const modelID = typeof body.model.id === "string"
    ? body.model.id
    : typeof body.model.modelID === "string" ? body.model.modelID : undefined
  if (typeof body.model.providerID !== "string" || !modelID) return
  return { agent: body.agent, model: { providerID: body.model.providerID, modelID } }
}

// The continuation must be the server-derived prompt for the session's current
// failed turn, and each user turn gets a bounded number of recoveries.
async function requireCurrentContinuation(directory: string, input: ChatContinuationInput, options: EngineRequestOptions) {
  const failedMessageID = input.progressToken.match(/^malformed:([A-Za-z0-9_-]{1,512})$/)?.[1]
  const canonical = failedMessageID ? malformedToolContinuationInput(input.sessionID, failedMessageID) : undefined
  if (
    !canonical ||
    input.kind !== canonical.kind ||
    input.runId !== canonical.runId ||
    input.messageID !== canonical.messageID ||
    input.text !== canonical.text
  ) throw new ChatContinuationConflictError("The interrupted-turn continuation is not current.")
  const [session, messages] = await Promise.all([
    engineJson(directory, `/api/session/${encodeURIComponent(input.sessionID)}`, options),
    engineJson(directory, `/api/session/${encodeURIComponent(input.sessionID)}/message`, options),
  ])
  if (!isRecord(session) || session.id !== input.sessionID)
    throw new Error("The owning session is unavailable for interrupted-turn recovery.")
  // Chat recovery owns only top-level chats. A child session (a subagent or a
  // lean drafting stage) or a runtime-owned session recovers through its runner.
  if (typeof session.parentID === "string" || (isRecord(session.metadata) && session.metadata.dochaus_role !== undefined))
    throw new ChatContinuationConflictError("Interrupted-turn recovery applies only to top-level chat sessions.")
  if (!Array.isArray(messages)) throw new Error("The owning session history is unavailable for interrupted-turn recovery.")
  const last = messages.at(-1)
  const failed = isRecord(last) && isRecord(last.info) && last.info.role === "user" && last.info.id === input.messageID
    ? messages.at(-2)
    : last
  if (
    !isRecord(failed) ||
    !isRecord(failed.info) ||
    failed.info.role !== "assistant" ||
    failed.info.id !== failedMessageID ||
    !malformedToolFailure(failed)
  ) throw new ChatContinuationConflictError("The interrupted model turn is no longer the current recoverable failure.")
  const boundary = messages.findLastIndex(message =>
    isUserMessage(message) && !/<dochaus-recovery\b/i.test(messageText(message)))
  const recoveries = messages.slice(boundary + 1).filter(message =>
    isUserMessage(message) && /<dochaus-recovery\b[^>]*kind="malformed-tool"/i.test(messageText(message))).length
  if (recoveries >= MAX_MALFORMED_TOOL_RECOVERIES_PER_TURN)
    throw new ChatContinuationConflictError("The interrupted turn exhausted its bounded malformed-output recovery attempts.")
}

function malformedToolFailure(message: Record<string, unknown>) {
  if (!Array.isArray(message.parts)) return false
  const toolSyntax = /(?:<\|tool_call\|>|call:[a-z][\w-]*\s*\{)/i
  const printedToolSyntax = message.parts.some(part =>
    isRecord(part) &&
    (part.type === "reasoning" || (part.type === "text" && part.synthetic !== true)) &&
    typeof part.text === "string" &&
    toolSyntax.test(part.text))
  const hasAnswer = message.parts.some(part =>
    isRecord(part) &&
    part.type === "text" &&
    part.synthetic !== true &&
    typeof part.text === "string" &&
    /[\p{L}\p{N}]/u.test(part.text) &&
    !toolSyntax.test(part.text))
  return printedToolSyntax && !hasAnswer
}

function requireMatchingClaim(claim: Claim, input: ChatContinuationInput, hash: string) {
  if (
    claim.session_id !== input.sessionID ||
    claim.kind !== input.kind ||
    claim.run_id !== input.runId ||
    claim.progress_token !== input.progressToken ||
    claim.payload_hash !== hash
  ) throw new ChatContinuationConflictError("The continuation message ID is already claimed by another payload.")
}

function payloadHash(input: Pick<ChatContinuationInput, "kind" | "runId" | "progressToken" | "sessionID" | "messageID" | "text">) {
  return createHash("sha256")
    .update(JSON.stringify({
      kind: input.kind,
      runId: input.runId,
      progressToken: input.progressToken,
      sessionID: input.sessionID,
      messageID: input.messageID,
      text: input.text,
    }))
    .digest("hex")
}

function input(claim: Claim) {
  return { kind: claim.kind as ChatContinuationInput["kind"], runId: claim.run_id, progressToken: claim.progress_token, sessionID: claim.session_id, messageID: claim.message_id }
}

// The status records what this service last saw of the engine; the engine
// stays the truth, and the next call for the same claim reads it again. A
// status write that finds the database locked (an upload commit holds the lock
// for seconds) is skipped, so a turn the engine already accepted is never
// reported as an error.
function setClaimStatus(db: ReturnType<typeof openDb>, messageID: string, status: Claim["status"]) {
  try {
    db.query("UPDATE continuation_dispatches SET status = ?, updated_at = ? WHERE message_id = ?").run(status, Date.now(), messageID)
  } catch (error) {
    if (!String((error as { code?: unknown }).code).startsWith("SQLITE_BUSY")) throw error
    console.error("continuation status not recorded", diagnosticErrorMetadata(error))
  }
}

async function engineJson(directory: string, endpoint: string, options: EngineRequestOptions): Promise<unknown> {
  const response = await boundedEngineRequest(engineV2Fetch(options.fetch), `${engineURL(options.engineURL)}${endpoint}`, {
    headers: { "x-opencode-directory": directory },
  }).catch(() => undefined)
  return response?.ok ? response.json().catch(() => undefined) : undefined
}

function isUserMessage(message: unknown): message is Record<string, unknown> {
  return isRecord(message) && isRecord(message.info) && message.info.role === "user"
}

function messageText(message: Record<string, unknown>) {
  if (!Array.isArray(message.parts)) return ""
  return message.parts
    .filter(part => isRecord(part) && part.type === "text" && typeof part.text === "string")
    .map(part => String(part.text))
    .join("")
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value)
}
