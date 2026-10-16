import { dispatchMatterContinuation } from "../ingest/drafting-run"
import type { Part } from "../session-view"
import { boundedRequest, matterClient, type Client } from "./client"
import { getConfig, modelSpec } from "./config"
import { getMessages, getSessionStatus } from "./sessions"

type ContinuationOwner = {
  agent?: string
  model?: { providerID: string; modelID: string }
}

function continuationOwner(value: unknown): ContinuationOwner {
  if (typeof value !== "object" || value === null) return {}
  const record = value as Record<string, unknown>
  const model = typeof record.model === "object" && record.model !== null
    ? record.model as Record<string, unknown>
    : undefined
  const modelID = typeof model?.modelID === "string"
    ? model.modelID
    : typeof model?.id === "string"
      ? model.id
      : undefined
  return {
    agent: typeof record.agent === "string" ? record.agent : undefined,
    model: typeof model?.providerID === "string" && modelID
      ? { providerID: model.providerID, modelID }
      : undefined,
  }
}

// The browser recovers only the explicit interrupted-turn case it directly
// observed in the mounted conversation, through ingest's continuation route.
export async function sendMalformedToolRecovery(
  directory: string,
  sessionID: string,
  recovery: MalformedToolRecovery,
) {
  if (recovery.runId !== sessionID)
    throw new Error("Interrupted-turn recovery must stay on its owning session.")
  if (typeof recovery.messageID !== "string" || !recovery.messageID.startsWith("msg_"))
    throw new Error("Workflow recovery requires its durable message ID.")
  const failedMessageID = recovery.progressToken.match(/^malformed:([A-Za-z0-9_-]{1,512})$/)?.[1]
  if (!failedMessageID || recovery.messageID !== `msg_recovery_${failedMessageID}`)
    throw new Error("Interrupted-turn recovery requires its exact failed message identity.")
  const marker =
    `<dochaus-recovery kind="malformed-tool" run-id="${recovery.runId}" ` +
    `progress-token="${recovery.progressToken}" />`
  return sendMarkedPrompt(directory, sessionID, marker, recovery)
}

export type MalformedToolRecovery = {
  kind: "malformed-tool"
  runId: string
  progressToken: string
  messageID: string
}

export type MarkedPromptResult = { admitted: boolean; executing: boolean; persisted: boolean }

const MARKED_PROMPTS = new Map<string, Promise<MarkedPromptResult>>()

async function sendMarkedPrompt(
  directory: string,
  sessionID: string,
  marker: string,
  recovery: MalformedToolRecovery,
) {
  const key = `${directory}:${sessionID}:${marker}`
  const current = MARKED_PROMPTS.get(key)
  if (current) return current
  const pending = withBrowserLock(`dochaus-continuation:${key}`, async () => {
    const client = matterClient(directory)
    const [messages, owner] = await Promise.all([
      boundedRequest(getMessages(client, sessionID), 3_000, "Session history"),
      boundedRequest(client.session.get({ sessionID }), 3_000, "Session context").catch(() => undefined),
    ])
    const existing = markerState(messages, marker)
    const exact = recoveryMessageState(messages, recovery.messageID)
    if (exact) {
      const status = await boundedRequest(
        getSessionStatus(client, sessionID),
        3_000,
        "Session status",
      ).catch(() => undefined)
      return {
        admitted: false,
        executing: status?.type === "busy",
        persisted: true,
      }
    }
    if (existing.admissions >= 2)
      throw new Error("The workflow continuation was admitted twice without durable progress; manual review is required.")
    const configuredModel = modelSpec(
      await boundedRequest(getConfig(), 3_000, "Model configuration").catch(() => undefined),
    )
    const latestUser = [...messages].reverse().find((message) => message.info.role === "user")
    const ownerContext = continuationOwner(owner)
    const messageContext = continuationOwner(latestUser?.info)
    const model = ownerContext.model ?? messageContext.model ?? configuredModel
    if (model && (existing.admissions === 1 || recoveryHistoryNeedsCompaction(messages)))
      await boundedRequest(
        client.session.compact({ sessionID }),
        120_000,
        "Session compaction",
      ).catch(() => undefined)
    const dispatched = await dispatchMatterContinuation(matterIDFromDirectory(directory), {
      kind: "malformed-tool",
      sessionID,
      progressToken: recovery.progressToken,
    })
    const visible = dispatched.persisted
      ? true
      : dispatched.accepted
        ? await waitForContinuationVisibility(client, sessionID, recovery.messageID)
        : false
    return {
      admitted: dispatched.attempted,
      // A reconciled claim can be accepted and persisted even though its turn
      // has already settled. Only a fresh accepted admission or current engine
      // execution proves that this attempt is running.
      executing: dispatched.executing || (!dispatched.existing && dispatched.accepted),
      persisted: dispatched.persisted || visible,
    }
  })
  MARKED_PROMPTS.set(key, pending)
  return pending.finally(() => MARKED_PROMPTS.delete(key))
}

function matterIDFromDirectory(directory: string) {
  const matterID = directory.replace(/[\\/]+$/, "").split(/[\\/]/).at(-1)
  if (!matterID) throw new Error("Workflow continuation requires a matter directory.")
  return matterID
}

async function waitForContinuationVisibility(client: Client, sessionID: string, messageID: string) {
  const deadline = Date.now() + 2_000
  while (Date.now() < deadline) {
    const [messages, status] = await Promise.all([
      boundedRequest(getMessages(client, sessionID), 1_000, "Continuation history").catch(() => []),
      boundedRequest(getSessionStatus(client, sessionID), 1_000, "Continuation status").catch(() => undefined),
    ])
    if (messages.some(
      (message) =>
        message.info.role === "user" &&
        message.info.id === messageID,
    )) return true
    // Busy is useful execution evidence for auto-recovery, but it does not
    // prove that this exact continuation was durably admitted. Keep polling
    // history so an acknowledgement cannot race a different in-flight turn.
    void status
    await new Promise((resolve) => setTimeout(resolve, 100))
  }
  return false
}

function recoveryMessageState(messages: Awaited<ReturnType<typeof getMessages>>, messageID: string) {
  const index = messages.findIndex((message) => message.info.id === messageID && message.info.role === "user")
  if (index < 0) return undefined
  return recoveryAttempt(messages, index)
}

export function recoveryHistoryNeedsCompaction(messages: Awaited<ReturnType<typeof getMessages>>) {
  const latest = [...messages].reverse().find((message) => message.info.role === "assistant")
  if (latest?.info.role === "assistant" && latest.info.finish === "length") return true
  const characters = activeRecoveryHistory(messages).reduce(
    (messageTotal, message) =>
      messageTotal +
      message.parts.reduce((partTotal, part) => {
        if (part.type === "text" || part.type === "reasoning") return partTotal + part.text.length
        if (part.type !== "tool") return partTotal
        const input = JSON.stringify(part.state.input ?? {})
        const output = part.state.status === "completed" ? part.state.output : part.state.status === "error" ? part.state.error : ""
        return partTotal + input.length + output.length
      }, 0),
    0,
  )
  // Recovery is the safe boundary for compressing a long activity. At roughly
  // 256k characters of persisted prompts/tool results, stale stage plans can
  // dominate even when the provider still has raw context capacity.
  return characters >= 256_000
}

// The session message endpoint returns the full persisted transcript even after
// compaction. Count only the history the engine would project to the model;
// otherwise every later recovery sees the same old bytes and summarizes again
// forever. A malformed retained-tail pointer fails open to the full transcript,
// matching the engine's own compaction filter.
function activeRecoveryHistory(messages: Awaited<ReturnType<typeof getMessages>>) {
  for (let index = messages.length - 1; index >= 0; index -= 1) {
    const message = messages[index]!
    if (message.info.role !== "user") continue
    const part = message.parts.find((candidate) => candidate.type === "compaction") as
      | (Part & { type: "compaction"; tail_start_id?: string })
      | undefined
    if (!part) continue
    const summary = messages.some(
      (candidate, candidateIndex) =>
        candidateIndex > index &&
        candidate.info.role === "assistant" &&
        candidate.info.parentID === message.info.id &&
        candidate.info.summary === true &&
        Boolean(candidate.info.finish) &&
        !candidate.info.error,
    )
    if (!summary) continue
    if (!part.tail_start_id) return messages.slice(index)
    const tailIndex = messages.findIndex((candidate) => candidate.info.id === part.tail_start_id)
    if (tailIndex < 0 || tailIndex >= index) return messages
    return messages.slice(tailIndex)
  }
  return messages
}

async function withBrowserLock<T>(name: string, run: () => Promise<T>) {
  if (typeof navigator === "undefined" || !navigator.locks) return run()
  return navigator.locks.request(name, run)
}

function markerState(messages: Awaited<ReturnType<typeof getMessages>>, marker: string) {
  const admissions = messages.flatMap((message, index) =>
    message.info.role === "user" && message.parts.some((part) => part.type === "text" && part.text.includes(marker))
      ? [index]
      : [],
  )
  const attempts = admissions.map((index) => recoveryAttempt(messages, index))
  return {
    admissions: admissions.length,
    latestSettled: attempts.at(-1)?.settled ?? false,
  }
}

function recoveryAttempt(messages: Awaited<ReturnType<typeof getMessages>>, admissionIndex: number) {
  const nextUserOffset = messages.slice(admissionIndex + 1).findIndex((message) => message.info.role === "user")
  const end = nextUserOffset < 0 ? messages.length : admissionIndex + 1 + nextUserOffset
  const assistants = messages.slice(admissionIndex + 1, end).filter((message) => message.info.role === "assistant")
  return {
    settled: assistants.length > 0,
    productive: assistants.some((message) =>
      message.parts.some(
        (part) =>
          part.type === "tool" ||
          (part.type === "text" && !part.synthetic && /[\p{L}\p{N}]/u.test(part.text)),
      ),
    ),
  }
}
