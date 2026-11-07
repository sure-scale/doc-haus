import { parseJsonObject } from "./checks"
import { parallel } from "./checklist-flow"
import { appendLeanEvent, leanElapsedMs } from "./events"
import { OutputLimitError } from "./native-stage-result"
import { retryPrompt, tagged } from "./prompts"
import { factContext, laterFactInstructions, missingFactEvidence, originalFactHold } from "./fact-answers"
import { matterSourceIDs, quoteFound } from "./quotes"
import { LEAN_CHILD_PERMISSIONS, LEAN_MAX_RETRIES, LEAN_STAGE_AGENTS, type LeanDeps } from "./stages"
import { ANSWER_CLASSES, FACT_STATUSES, factDisposition, questionOptions, readLeanText, readQuestions, recordAnswerClasses, unclassifiedAnswer, writeQuestions, type AnswerClass, type FactStatus, type LeanState, type Question } from "./state"

// How the lawyer's own words answer a question, read once when the coordinator
// takes the answer and before any step reads it: one fresh read-only session
// per answer, with the same-session bounded retry of every lean job. The reply
// is one of a closed set code accepts (answered, delegated, declined); a
// delegating answer is then recorded as Decide for me records an adoption
// (state.ts classifiedAnswer). A reply that never passes its check stands as
// the lawyer's own answer, as every typed answer did before, and so does a reply
// cut off at its output limit; any other runtime error or a pause records
// nothing, so the next pass reads the answer again.

/** Reads every answer in the lawyer's own words not yet read and records each class. */
export async function classifyAnswers(directory: string, state: LeanState, deps: LeanDeps) {
  const unread = readQuestions(directory).filter(unclassifiedAnswer)
  if (!unread.length) return
  const input = factContext(directory)
  const read = new Map<string, { words: string; answer_class: AnswerClass; fact_status?: FactStatus; fact_context?: string }>()
  await parallel(unread, async question => {
    const classified = await classifyAnswer(directory, state, deps, question)
    if (classified) read.set(question.id, { words: question.answer!, ...classified, ...(classified.fact_status ? { fact_context: question.fact_context ?? input } : {}) })
  })
  if (deps.pause?.aborted || factContext(directory) !== input) return
  recordAnswerClasses(directory, read)
}

/** Existing waits are reconsidered on changed instructions/evidence only.
 * Missing old context is bound once, without reinterpreting a chosen option. */
export async function refreshFactAnswers(directory: string, state: LeanState, deps: LeanDeps) {
  const input = factContext(directory)
  const pending = readQuestions(directory).filter(question => question.stage === "intake" && question.kind === "fact" &&
    question.status === "answered" && (["wait", "unknown", "provisional"].includes(factDisposition(question) ?? "") ||
      factDisposition(question) === "supplied" && question.fact_resolution))
  const changes = new Map<string, { before: Question; status?: FactStatus; resolution?: Question["fact_resolution"]; clearResolution?: boolean }>()
  await parallel(pending, async question => {
    const missing = missingFactEvidence(directory, question)
    if (question.fact_context === input && !missing) return
    const safe = missing ? { ...question, fact_status: originalFactHold(directory, question), fact_resolution: undefined } : question
    const fallback = missing ? { status: safe.fact_status, clearResolution: true } : {}
    if (!question.fact_context) { changes.set(question.id, { before: question, ...fallback }); return }
    const outcome = await refreshFactAnswer(directory, state, deps, safe)
    changes.set(question.id, { before: question, ...fallback, ...(outcome ?? {}) })
  })
  if (!changes.size || deps.pause?.aborted || factContext(directory) !== input) return
  writeQuestions(directory, readQuestions(directory).map(question => {
    const change = changes.get(question.id)
    if (!change || JSON.stringify(question) !== JSON.stringify(change.before)) return question
    const { fact_resolution: _resolution, ...withoutResolution } = question
    return { ...(change.clearResolution ? withoutResolution : question), fact_context: input, ...(change.status ? { fact_status: change.status } : {}),
      ...(change.resolution ? { fact_resolution: change.resolution } : {}) }
  }))
}

type FactBasis = { source_id: string; quote: string }[]
type FactEvidence = { evidence_status: "supplied" | "missing"; reason: string; basis: FactBasis }
type FactPermission = { drafting_permission: "provisional" | "wait" | "none"; reason: string; basis: FactBasis }

async function refreshFactAnswer(directory: string, state: LeanState, deps: LeanDeps, question: Question): Promise<{ status: FactStatus; resolution?: Question["fact_resolution"]; clearResolution?: boolean } | undefined> {
  // Missing evidence and permission to draft with it missing are separate
  // questions. Neither job decides the other's answer.
  const evidence = await readFactRefresh(directory, state, deps, question, "evidence", factEvidencePrompt(directory, question), reply => checkFactEvidence(directory, question, reply))
  if (!evidence || deps.pause?.aborted) return
  const permission = await readFactRefresh(directory, state, deps, question, "permission", factPermissionPrompt(directory, question, evidence), reply => checkFactPermission(directory, question, evidence, reply))
  if (!permission) return
  const status: FactStatus = permission.drafting_permission === "wait" ? "wait" : evidence.evidence_status === "supplied" ? "supplied"
    : permission.drafting_permission === "provisional" ? "provisional" : "unknown"
  // A supplied fact's source revalidation must never look for the original
  // permission as if it were an indexed source. The original answer stays put.
  const basis = status === "supplied" ? [...evidence.basis, ...permission.basis.filter(row => row.source_id !== "original-answer")] : permission.basis
  const reason = status === "supplied" ? [evidence.reason, permission.reason].join(" ") : permission.reason
  return { status, ...(basis.length ? { resolution: { reason, basis } } : status !== factDisposition(question) ? { clearResolution: true } : {}) }
}

async function readFactRefresh<T>(directory: string, state: LeanState, deps: LeanDeps, question: Question, job: "evidence" | "permission", prompt: string,
  check: (reply: string) => { issues: string[]; value?: T }): Promise<T | undefined> {
  const agent = LEAN_STAGE_AGENTS.answer
  const at = { stage: "intake" as const, step: `reconsider fact ${question.id}: ${job}` }
  try {
    const sessionID = await deps.sessions.createSession({ directory, parentID: state.owner_session_id, agent,
      title: job === "evidence" ? "Check whether the requested fact is now supplied" : "Check the lawyer's permission to draft", model: state.model, permissions: LEAN_CHILD_PERMISSIONS })
    let issues: string[] = []
    for (let index = 0; index <= LEAN_MAX_RETRIES; index += 1) {
      appendLeanEvent(directory, { ...at, event: index ? "retry" : "model_call", attempt: index, session_id: sessionID })
      const started = performance.now()
      const reply = await deps.sessions.prompt({ directory, sessionID, agent, text: index ? retryPrompt("answer", issues) : prompt })
      const replied = performance.now()
      const checked = check(reply)
      issues = checked.issues
      appendLeanEvent(directory, { ...at, event: "checker", attempt: index, issues, wait_ms: leanElapsedMs(started, replied), check_ms: leanElapsedMs(replied) })
      if (checked.value) return checked.value
    }
  } catch (error) {
    // Exhausted model output leaves the existing hold in force under these
    // inputs. Runtime/provider interruptions stay retryable after recovery.
    if (!(error instanceof OutputLimitError)) throw error
  }
}

function factData(value: unknown) {
  return JSON.stringify(value).replaceAll("<", "\\u003c").replaceAll(">", "\\u003e").replaceAll("&", "\\u0026")
}

function factQuestion(question: Question) {
  return factData({ question: question.question, context: question.context, answer: question.answer, answer_kind: question.answer_kind,
    selected_option: questionOptions(question).find(option => option.id === question.option_id), prior_status: factDisposition(question), prior_resolution: question.fact_resolution })
}

function factEvidencePrompt(directory: string, question: Question) {
  return [
    "Check ONE factual uncertainty: is the exact fact or completed check requested by this question now supplied? Do not decide whether drafting is permitted, answer for the lawyer, invent evidence, or perform the outside task.",
    "Use targeted source lookup only for this question. Preserve the question's evidence conditions, source qualifications and contrary evidence. A repeated unverified assertion is not the completed verification requested. A request to obtain or check evidence, a source name, or permission to draft supplies no missing fact. Supplied means the requested information is actually present, not that it has been independently verified beyond what the evidence establishes. If the requested fact or check remains absent or uncertain, report missing even if a marked working draft is authorised.",
    tagged("original_fact_question", factQuestion(question)),
    tagged("current_request", factData(readLeanText(directory, "request.md") ?? "")),
    tagged("later_instructions", factData(laterFactInstructions(directory, question))),
    tagged("source_names", factData(matterSourceIDs(directory))),
    'Reply JSON only: {"evidence_status":"supplied|missing","reason":"whether the exact requested information or check is now supplied, preserving qualifications","basis":[{"source_id":"request, updated-notes or an exact matter source name","quote":"exact relevant words"}]}. Supplied requires a checked supporting quotation. Missing may have an empty basis. original-answer is a prior treatment decision, not new evidence. Later instructions contain only updates received after the original answer.',
  ].join("\n\n")
}

function factPermissionPrompt(directory: string, question: Question, evidence: FactEvidence) {
  return [
    "Decide ONE permission question: under the lawyer's original answer and later instructions, may drafting proceed with this fact in the checked state below? Do not reassess or supply the fact. Evidence availability and drafting permission are separate; a fact remaining unverified does not cancel express permission to prepare a marked working draft.",
    "Read all later instructions in receipt order; an earlier permission never overrides a later hold. Unrelated updates never revoke already authorised provisional drafting. Provisional requires express human permission to draft with the uncertainty visibly marked. A genuinely human-picked option is the lawyer's instruction, including an older option without a code-owned fact_action; an assistant recommendation alone or an adopted professional judgment supplies no human approval. The original request predates the answer and cannot override it.",
    "Keep an actual wait, approval condition or release restriction binding while its condition remains unsatisfied, even when evidence arrives. Distinguish a hold on drafting from a hold on sending, signing, serving or release: permission for a marked working draft does not release it. none means no remaining drafting restriction and no need for provisional permission: for example, supplied evidence satisfies the sole condition to wait for that fact. Evidence never creates human approval or satisfies an independent approval condition. Preserve the current treatment when no actual change is established.",
    tagged("original_fact_question", factQuestion(question)),
    tagged("later_instructions", factData(laterFactInstructions(directory, question))),
    tagged("checked_fact_evidence", factData(evidence)),
    'Reply JSON only: {"drafting_permission":"provisional|wait|none","reason":"the current permission and any condition still binding","basis":[{"source_id":"updated-notes or original-answer","quote":"exact relevant human words"}]}. A new permission or hold requires checked supporting human words under updated-notes. To preserve or reestablish express provisional permission in the original typed reply or human-picked option after an unrelated update, quote it under original-answer and explain why no later instruction supersedes it; never do so while a later hold remains unsatisfied. Keeping an existing treatment needs no new approval and basis may be empty. Do not cite source documents as permission. Missing chronology cannot establish a later override.',
  ].join("\n\n")
}

function factRefreshValue(reply: string, field: "evidence_status" | "drafting_permission", statuses: readonly string[]): { issues: string[]; value?: { status: string; reason: string; basis: FactBasis } } {
  const parsed = parseJsonObject(reply)
  if (!parsed.ok) return { issues: [`Reply with one JSON object. ${parsed.error}`] }
  const value = parsed.value
  if (!statuses.includes(value[field] as string) || typeof value.reason !== "string" || !value.reason.trim() || !Array.isArray(value.basis))
    return { issues: [`Give ${field}, its reason, and a basis array.`] }
  const basis: FactBasis = []
  for (const row of value.basis) {
    if (!row || typeof row !== "object" || typeof row.source_id !== "string" || typeof row.quote !== "string" || !row.quote.trim())
      return { issues: ["Every basis needs a known source_id and exact supporting quote."] }
    basis.push({ source_id: row.source_id, quote: row.quote })
  }
  return { issues: [], value: { status: value[field] as string, reason: value.reason.trim(), basis } }
}

function checkFactEvidence(directory: string, question: Question, reply: string): { issues: string[]; value?: FactEvidence } {
  const checked = factRefreshValue(reply, "evidence_status", ["supplied", "missing"])
  if (!checked.value) return { issues: checked.issues }
  const { status, reason, basis } = checked.value
  const allowed = new Set([...matterSourceIDs(directory), "request", "updated-notes"])
  for (const row of basis) {
    if (!allowed.has(row.source_id)) return { issues: ["Fact evidence must quote the request, a later update or an exact matter source, never original-answer."] }
    const found = row.source_id === "request" || row.source_id === "updated-notes"
      ? (row.source_id === "request" ? readLeanText(directory, "request.md") ?? "" : laterFactInstructions(directory, question)).includes(row.quote)
      : quoteFound(directory, row)
    if (!found) return { issues: [`The supporting quote is not in ${row.source_id}.`] }
  }
  if (status === "supplied" && !basis.length) return { issues: ["A supplied fact requires actual quoted evidence satisfying the question's conditions."] }
  return { issues: [], value: { evidence_status: status as FactEvidence["evidence_status"], reason, basis } }
}

function checkFactPermission(directory: string, question: Question, evidence: FactEvidence, reply: string): { issues: string[]; value?: FactPermission } {
  const checked = factRefreshValue(reply, "drafting_permission", ["provisional", "wait", "none"])
  if (!checked.value) return { issues: checked.issues }
  const { status, reason, basis } = checked.value
  const selected = questionOptions(question).find(option => option.id === question.option_id)
  const originalPermission = !question.answer_kind || question.answer_kind === "option" && selected?.text === question.answer
  for (const row of basis) {
    if (row.source_id !== "updated-notes" && row.source_id !== "original-answer")
      return { issues: ["Drafting permission must quote human instructions, never a source-derived assumption or the earlier request."] }
    if (row.source_id === "original-answer" && (status !== "provisional" || !originalPermission))
      return { issues: ["original-answer can establish only express permission in the lawyer's own reply or human-picked option, never an assistant recommendation or a new override."] }
    const text = row.source_id === "original-answer" ? question.answer ?? "" : laterFactInstructions(directory, question)
    if (!text.includes(row.quote)) return { issues: [`The supporting quote is not in ${row.source_id}.`] }
  }
  const prior = factDisposition(question)
  if (status === "provisional" && prior !== "provisional" && !basis.length)
    return { issues: ["Provisional drafting requires quoted later human permission or express original permission that no later hold supersedes."] }
  if (status === "wait" && prior !== "wait" && !basis.some(row => row.source_id === "updated-notes"))
    return { issues: ["A new drafting hold requires quoted later human instructions."] }
  if (status === "none" && evidence.evidence_status === "missing" && (prior === "wait" || prior === "provisional"))
    return { issues: ["The fact is still missing: preserve the existing wait or provisional permission unless later human instructions change it."] }
  return { issues: [], value: { drafting_permission: status as FactPermission["drafting_permission"], reason, basis } }
}

async function classifyAnswer(directory: string, state: LeanState, deps: LeanDeps, question: Question): Promise<{ answer_class: AnswerClass; fact_status?: FactStatus } | undefined> {
  const agent = LEAN_STAGE_AGENTS.answer
  const at = { stage: question.stage, document: question.document, step: `answer ${question.id}` }
  try {
    const sessionID = await deps.sessions.createSession({ directory, parentID: state.owner_session_id, agent, title: `Lean drafting: answer: ${question.id}`, model: state.model, permissions: LEAN_CHILD_PERMISSIONS })
    let issues: string[] = []
    for (let index = 0; index <= LEAN_MAX_RETRIES; index += 1) {
      appendLeanEvent(directory, { ...at, event: index ? "retry" : "model_call", attempt: index, session_id: sessionID })
      const asked = performance.now()
      const reply = await deps.sessions.prompt({ directory, sessionID, agent, text: index ? retryPrompt("answer", issues) : answerClassPrompt(question) })
      const answered = performance.now()
      const checked = checkAnswerClass(reply, question)
      issues = checked.issues
      appendLeanEvent(directory, { ...at, event: "checker", attempt: index, issues, wait_ms: leanElapsedMs(asked, answered), check_ms: leanElapsedMs(answered) })
      if (checked.answer_class) return { answer_class: checked.answer_class, ...(checked.fact_status ? { fact_status: checked.fact_status } : {}) }
    }
    return { answer_class: "answered", ...(question.kind === "fact" && question.stage === "intake" ? { fact_status: "unknown" } : {}) }
  } catch (error) {
    // A reply cut off before any answer fails like one that never passes its check.
    if (error instanceof OutputLimitError) return { answer_class: "answered", ...(question.kind === "fact" && question.stage === "intake" ? { fact_status: "unknown" } : {}) }
    // A runtime error, a pause or a settings change: the runner acts on it.
    return undefined
  }
}

export function answerClassPrompt(question: Question) {
  const options = questionOptions(question)
  const advice = question.advice
  const recommended = advice && "recommendation" in advice ? options.find(option => option.id === advice.recommendation.option) : undefined
  return [
    "Read the lawyer's reply to ONE question raised while drafting and say how it answers that question. Do not answer the question yourself and do not look anything up.",
    tagged("question", `Kind: ${question.kind ?? "not stated"}\n${question.question}`),
    ...(question.context ? [tagged("context", question.context)] : []),
    ...(options.length ? [tagged("options", options.map(option => `- ${option.text}${option === recommended ? " (recommended)" : ""}`).join("\n"))] : []),
    tagged("reply", question.answer ?? ""),
    [
      'Reply with one JSON object only: {"class":"answered"}, {"class":"delegated"} or {"class":"declined"}.',
      '- "answered": the reply gives the lawyer\'s own answer, choice, instruction, fact or condition on this question, in whole or in part.',
      '- "delegated": the reply gives no answer of its own and delegates the choice to the drafter\'s professional judgment or to the recommendation.',
      '- "declined": the reply gives no answer of its own and does not delegate the choice.',
    ].join("\n"),
    ...(question.kind === "fact" && question.stage === "intake" ? [[
      'Also give "fact_status":"supplied|provisional|wait|unknown" and "fact_quote":"exact words from the lawyer\'s reply that establish this treatment" (quote may be empty only for unknown).',
      '- supplied: the reply supplies or confirms the requested fact itself, preserving all qualifications. This is the lawyer\'s assertion, not independent verification. An instruction to obtain, check or ask someone for evidence supplies no fact.',
      '- provisional: the lawyer expressly permits drafting while this exact fact remains unknown, with its uncertainty disclosed. Do not infer permission merely from a request to obtain evidence.',
      '- wait: the lawyer says to wait, defer drafting, or obtain evidence before drafting. Keep that instruction binding until the lawyer changes it.',
      '- unknown: the fact remains missing and the reply gives no clear permission to draft provisionally or explicit waiting instruction, or its meaning is uncertain. Never turn a task request, a hypothetical future result or delegation into a confirmed fact.',
    ].join("\n")] : []),
  ].join("\n\n")
}

/** The class a reply gives, or the issue for a retry. */
export function checkAnswerClass(reply: string, question?: Question): { issues: string[]; answer_class?: AnswerClass; fact_status?: FactStatus } {
  const parsed = parseJsonObject(reply)
  if (!parsed.ok) return { issues: [`Reply with one JSON object: ${parsed.error}.`] }
  const value = parsed.value.class
  if (!(ANSWER_CLASSES as readonly unknown[]).includes(value)) return { issues: ['"class" must be "answered", "delegated" or "declined".'] }
  if (question?.kind !== "fact" || question.stage !== "intake") return { issues: [], answer_class: value as AnswerClass }
  const status = parsed.value.fact_status
  if (!(FACT_STATUSES as readonly unknown[]).includes(status)) return { issues: ['Give "fact_status":"supplied|provisional|wait|unknown" for this fact question.'] }
  const quote = parsed.value.fact_quote
  if (status !== "unknown" && (typeof quote !== "string" || !quote.trim() || !question.answer?.includes(quote)))
    return { issues: ['"fact_quote" must copy the exact words of the lawyer\'s reply that establish this fact treatment.'] }
  if (status === "supplied" && value !== "answered") return { issues: ['A delegated or declined reply supplies no fact; preserve it as unknown unless it expressly permits provisional drafting or waiting.'] }
  return { issues: [], answer_class: value as AnswerClass, fact_status: status as FactStatus }
}
