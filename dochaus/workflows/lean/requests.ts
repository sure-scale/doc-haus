import { randomUUID } from "node:crypto"
import { closeSync, existsSync, fsyncSync, linkSync, mkdirSync, openSync, readdirSync, unlinkSync, writeSync } from "node:fs"
import path from "node:path"
import { readDurableJson } from "../../lib/durable-json"
import { appendLeanEvent } from "./events"
import { LEAN_PROMPT_VERSION, stageInputs } from "./stages"
import { checklistRepairItems, LEAN_MAX_SUBSTANTIVE_REPAIRS, readRepairRound, startRepairRound } from "./repairs"
import { LeanQuestionError, hashStageInputs, answerQuestion, leanDir, readLeanJson, readLeanText, writeLeanText, type LeanState, type OpenItem } from "./state"

// Lawyer requests to a lean drafting run. Continue and answer POSTs only append a
// request file; the lean runner is the single writer of the run's state and
// drains the requests itself at stage boundaries, so a request never races the
// stage that is running. Each request is created under a unique name with an
// exclusive link (it can never overwrite another), processed once in arrival
// order, and moved to requests/done/ with its outcome.

export type ChecklistRefinementSelection = { parent_id: string; scope: "source_collection" | "cited_sources" | "target_body" }
export type LeanRequest = { kind: "continue" } | { kind: "answer"; question_id: string; answer: string } | { kind: "update"; text: string } | { kind: "clarification"; text: string } | { kind: "refine_checklist"; parents: ChecklistRefinementSelection[] }
export type StoredLeanRequest = LeanRequest & { id: string; received_at: string }
export type DrainOutcome = { answered: number; updated: number; repairRound: boolean }

const sequence = { next: 0 }

function requestsDir(directory: string) {
  return path.join(leanDir(directory), "requests")
}

export function enqueueLeanRequest(directory: string, request: LeanRequest): StoredLeanRequest {
  const dir = requestsDir(directory)
  mkdirSync(path.join(dir, ".incoming"), { recursive: true, mode: 0o700 })
  const id = `${String(Date.now()).padStart(15, "0")}-${String(sequence.next++ % 1_000_000).padStart(6, "0")}-${randomUUID().slice(0, 8)}`
  const stored: StoredLeanRequest = { ...request, id, received_at: new Date().toISOString() }
  const incoming = path.join(dir, ".incoming", randomUUID())
  const fd = openSync(incoming, "wx", 0o600)
  try {
    writeSync(fd, `${JSON.stringify(stored)}\n`)
    fsyncSync(fd)
  } finally {
    closeSync(fd)
  }
  // link(2) fails rather than replace an existing name, so no request is lost.
  linkSync(incoming, path.join(dir, `${id}.json`))
  unlinkSync(incoming)
  appendLeanEvent(directory, { stage: "run", event: "request_received", request_id: id, kind: request.kind })
  return stored
}

/** Names of the requests not yet processed, oldest first. */
export function pendingLeanRequests(directory: string) {
  const dir = requestsDir(directory)
  if (!existsSync(dir)) return []
  return readdirSync(dir).filter(name => name.endsWith(".json")).toSorted()
}

/** Question ids named by answers still waiting in the queue. */
export function queuedAnswerIDs(directory: string) {
  return new Set(pendingLeanRequests(directory).flatMap(name => {
    const request = readDurableJson(path.join(requestsDir(directory), name))
    return isAnswer(request) ? [request.question_id] : []
  }))
}

/** Processes pending requests in order. Answers are always applied. A Continue
 * is processed only when `stopped` is the run's settled state: on a blocked gate
 * it starts one repair round; otherwise it simply resumes. While a pass is in
 * progress a Continue stays queued until the pass settles. */
export function drainLeanRequests(directory: string, stopped?: LeanState): DrainOutcome {
  const pending = pendingLeanRequests(directory)
  if (pending.length) mkdirSync(path.join(requestsDir(directory), "done"), { recursive: true, mode: 0o700 })
  return pending.reduce<DrainOutcome>((total, name) => {
    const file = path.join(requestsDir(directory), name)
    const request = readDurableJson(file)
    if ((isContinue(request) || isChecklistRefinement(request) || isClarification(request)) && !stopped) return total
    const outcome = processRequest(directory, request, stopped)
    writeLeanText(directory, path.join("requests", "done", name), `${JSON.stringify({
      ...(record(request) ? request : { invalid: true }), outcome, processed_at: new Date().toISOString(),
    }, null, 2)}\n`)
    unlinkSync(file)
    appendLeanEvent(directory, { stage: "run", event: "request_processed", request_id: name.replace(/\.json$/, ""), outcome: outcome.status })
    return {
      answered: total.answered + (outcome.status === "answered" ? 1 : 0),
      updated: total.updated + (["included", "clarification_included"].includes(outcome.status) ? 1 : 0),
      repairRound: total.repairRound || outcome.status === "repair_round",
    }
  }, { answered: 0, updated: 0, repairRound: false })
}

function processRequest(directory: string, request: unknown, stopped: LeanState | undefined): { status: string; detail?: string } {
  if (isChecklistRefinement(request)) {
    if (stopped?.status !== "blocked" || stopped.stages["checklist-check"]?.status !== "blocked")
      return { status: "rejected", detail: "Checklist refinement requires a settled blocked checklist stage." }
    return { status: "refinement_included" }
  }
  if (isClarification(request)) {
    if (stopped?.status !== "blocked" || stopped.stages["checklist-check"]?.status !== "blocked")
      return { status: "rejected", detail: "Clarification requires a settled blocked checklist; no document edit was admitted." }
    // The processed request is the context. Do not rewrite the brief or body.
    return { status: "clarification_included" }
  }
  if (isUpdate(request) && typeof request.id === "string") {
    const previous = readLeanText(directory, "updated-notes.md") ?? ""
    const marker = `<!-- update:${request.id} -->`
    if (!previous.includes(marker)) writeLeanText(directory, "updated-notes.md", `${previous}${marker}\n${request.text}\n\n`)
    return { status: "included" }
  }
  if (isAnswer(request)) {
    try {
      answerQuestion(directory, request.question_id, request.answer)
      return { status: "answered" }
    } catch (error) {
      if (error instanceof LeanQuestionError) return { status: "rejected", detail: `${error.code}: ${error.message}` }
      throw error
    }
  }
  if (!isContinue(request)) return { status: "rejected", detail: "unreadable request" }
  if (stopped?.status === "blocked" && stopped.stages["checklist-check"]?.status === "blocked") {
    // Changed source, answer or Word inputs must be inspected before an old
    // adverse verdict can consume a substantive repair attempt.
    const current = hashStageInputs("checklist-check", LEAN_PROMPT_VERSION, stageInputs(directory, "checklist-check"))
    const findings = stopped.stages["checklist-check"]?.input_sha === current ? checklistRepairItems(directory) : []
    if (findings.length) {
      if ((readRepairRound(directory)?.substantive_attempts ?? 0) >= LEAN_MAX_SUBSTANTIVE_REPAIRS)
        return { status: "rejected", detail: `Checklist corrections remain after ${LEAN_MAX_SUBSTANTIVE_REPAIRS} substantive repair rounds; no new body repair was admitted.` }
      const round = startRepairRound(directory, findings, true)
      return { status: "repair_round", detail: `round ${round.round}: ${findings.length} checklist finding(s) submitted for source adjudication` }
    }
    return { status: "resume" }
  }
  const blocking = stopped?.status === "blocked" ? blockedGateItems(directory) : []
  if (!blocking.length) return { status: "resume" }
  const round = startRepairRound(directory, blocking)
  return { status: "repair_round", detail: `round ${round.round}: ${blocking.length} item(s)` }
}

function blockedGateItems(directory: string): OpenItem[] {
  const gate = readLeanJson(directory, "gate.json")
  if (!record(gate) || gate.status !== "blocked" || !Array.isArray(gate.blocking)) return []
  return gate.blocking.filter(record) as OpenItem[]
}

function isContinue(value: unknown): value is { kind: "continue" } {
  return record(value) && value.kind === "continue"
}

function isAnswer(value: unknown): value is { kind: "answer"; question_id: string; answer: string } {
  return record(value) && value.kind === "answer" && typeof value.question_id === "string" && typeof value.answer === "string"
}

function record(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value)
}

function isUpdate(value: unknown): value is { kind: "update"; text: string; id?: unknown; outcome?: unknown } {
  return record(value) && value.kind === "update" && typeof value.text === "string"
}

function isClarification(value: unknown): value is { kind: "clarification"; text: string; id?: unknown; outcome?: unknown } {
  return record(value) && value.kind === "clarification" && typeof value.text === "string"
}

/** Human inspection context from accepted queue requests, never new drafting
 * instructions. Identical text adds no new evidence or reinspection allowance. */
export function leanClarifications(directory: string) {
  const done = path.join(requestsDir(directory), "done")
  if (!existsSync(done)) return []
  const texts = new Set<string>()
  return readdirSync(done).filter(name => name.endsWith(".json")).toSorted().flatMap(name => {
    const value = readDurableJson(path.join(done, name))
    if (!isClarification(value) || !record(value.outcome) || value.outcome.status !== "clarification_included" || texts.has(value.text)) return []
    texts.add(value.text)
    return [{ document: "request", locator: `.dochaus/lean/requests/done/${name} /text`, quote: value.text }]
  })
}

export function isChecklistRefinement(value: unknown): value is { kind: "refine_checklist"; parents: ChecklistRefinementSelection[] } {
  return record(value) && value.kind === "refine_checklist" && Array.isArray(value.parents) && value.parents.length > 0 &&
    value.parents.every(parent => record(parent) && typeof parent.parent_id === "string" && parent.parent_id.trim().length > 0 &&
      (parent.scope === "source_collection" || parent.scope === "cited_sources" || parent.scope === "target_body")) &&
    new Set(value.parents.map(parent => parent.parent_id)).size === value.parents.length
}

/** Refinement changes late verification scope, never the lawyer's instructions.
 * Only selections accepted by the sole writer participate in the stage key. */
export function checklistRefinementSelections(directory: string): ChecklistRefinementSelection[] {
  const done = path.join(requestsDir(directory), "done")
  if (!existsSync(done)) return []
  const selected = new Map<string, ChecklistRefinementSelection>()
  for (const name of readdirSync(done).filter(name => name.endsWith(".json")).toSorted()) {
    const value = readDurableJson(path.join(done, name))
    if (!record(value) || !record(value.outcome) || value.outcome.status !== "refinement_included" || !isChecklistRefinement(value)) continue
    for (const parent of value.parents) selected.set(parent.parent_id, parent)
  }
  return [...selected.values()].sort((a, b) => a.parent_id.localeCompare(b.parent_id))
}

/** Stored user updates survive reload; included means prompt context, not applied wording. */
export function leanUserUpdates(directory: string) {
  const dir = requestsDir(directory)
  const done = path.join(dir, "done")
  const files = [
    ...pendingLeanRequests(directory).map(name => ({ file: path.join(dir, name), status: "queued" as const })),
    ...(existsSync(done) ? readdirSync(done).filter(name => name.endsWith(".json")).map(name => ({ file: path.join(done, name), status: "included" as const })) : []),
  ]
  return files.flatMap(({ file, status }) => {
    const value = readDurableJson(file)
    return (isUpdate(value) || isClarification(value)) && typeof value.id === "string" &&
      (status === "queued" || record(value.outcome) && ["included", "clarification_included"].includes(String(value.outcome.status)))
      ? [{ id: value.id, text: value.text, status, ...(isClarification(value) ? { kind: "clarification" as const } : {}) }] : []
  }).sort((a, b) => a.id.localeCompare(b.id))
}
