import { deleteSessionAnswerDrafts } from "./answer-drafts"
import { engineV2Fetch } from "../../../dochaus/lib/engine-v2"
import { leanDraftingStatus } from "./lean-drafting"

export class MatterSessionDeletionConflictError extends Error {}

export class MatterSessionDeletionError extends Error {}

export type EngineRequestOptions = {
  engineURL?: string
  fetch?: typeof fetch
}

const ENGINE_REQUEST_TIMEOUT_MS = 10_000
const MATTER_QUEUES = new Map<string, Promise<void>>()

// Destructive conversation operations stay behind the matter-aware service. It
// refuses to delete the owner of an unfinished lean drafting run (its stage
// sessions are children of that owner), then forwards permitted deletions to the
// engine with the same matter scope used by every other session operation.
export function deleteMatterSession(directory: string, sessionID: string, options: EngineRequestOptions = {}) {
  return serializeMatter(directory, async () => {
    if (!sessionID.trim() || sessionID.length > 512)
      throw new MatterSessionDeletionConflictError("The conversation identity is invalid.")
    const lean = leanDraftingStatus(directory)
    if (lean?.owner_session_id === sessionID && (lean.in_flight || (lean.status !== "completed" && lean.status !== "failed")))
      throw new MatterSessionDeletionConflictError("This conversation owns an unfinished drafting run and cannot be deleted.")
    const response = await boundedEngineRequest(
      engineV2Fetch(options.fetch),
      `${engineURL(options.engineURL)}/api/session/${encodeURIComponent(sessionID)}`,
      { method: "DELETE", headers: { "x-opencode-directory": directory } },
    ).catch(() => undefined)
    if (!response) throw new MatterSessionDeletionError("The conversation service is unavailable.")
    if (!response.ok && response.status !== 404)
      throw new MatterSessionDeletionError(`The conversation service rejected deletion (${response.status}).`)
    deleteSessionAnswerDrafts(directory, sessionID)
    return { ok: true, deleted: response.status !== 404 }
  })
}

export function engineURL(value?: string) {
  return (value ?? process.env.OPENCODE_URL ?? `http://127.0.0.1:${process.env.DOC_HAUS_ENGINE_PORT ?? "4097"}`)
    .replace(/\/$/, "")
}

export function boundedEngineRequest(request: typeof fetch, url: string, init: RequestInit = {}) {
  return request(url, { ...init, signal: init.signal ?? AbortSignal.timeout(ENGINE_REQUEST_TIMEOUT_MS) })
}

/** Runs session operations for one matter one at a time, in arrival order. */
export function serializeMatter<T>(directory: string, run: () => Promise<T>) {
  const prior = MATTER_QUEUES.get(directory) ?? Promise.resolve()
  const pending = prior.catch(() => undefined).then(run)
  const tail = pending.then(() => undefined, () => undefined)
  MATTER_QUEUES.set(directory, tail)
  return pending.finally(() => {
    if (MATTER_QUEUES.get(directory) === tail) MATTER_QUEUES.delete(directory)
  })
}
