import { parseJsonObject } from "./checks"
import { parallel } from "./checklist-flow"
import { appendLeanEvent, leanElapsedMs } from "./events"
import { OutputLimitError } from "./native-stage-result"
import { retryPrompt, tagged } from "./prompts"
import { factContext, laterFactInstructions, missingFactEvidence, originalFactHold, sourceResolvedFact } from "./fact-answers"
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
    question.status === "answered" && (["wait", "unknown", "provisional"].includes(factDisposition(question) ?? "") || sourceResolvedFact(question)))
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

async function refreshFactAnswer(directory: string, state: LeanState, deps: LeanDeps, question: Question): Promise<{ status: FactStatus; resolution?: Question["fact_resolution"] } | undefined> {
  const agent = LEAN_STAGE_AGENTS.answer
  const at = { stage: "intake" as const, step: `reconsider fact ${question.id}` }
  try {
    const sessionID = await deps.sessions.createSession({ directory, parentID: state.owner_session_id, agent,
      title: "Check whether new instructions or evidence change this wait", model: state.model, permissions: LEAN_CHILD_PERMISSIONS })
    let issues: string[] = []
    for (let index = 0; index <= LEAN_MAX_RETRIES; index += 1) {
      appendLeanEvent(directory, { ...at, event: index ? "retry" : "model_call", attempt: index, session_id: sessionID })
      const started = performance.now()
      const reply = await deps.sessions.prompt({ directory, sessionID, agent, text: index ? retryPrompt("answer", issues) : refreshFactPrompt(directory, question) })
      const replied = performance.now()
      const checked = checkFactRefresh(directory, question, reply)
      issues = checked.issues
      appendLeanEvent(directory, { ...at, event: "checker", attempt: index, issues, wait_ms: leanElapsedMs(started, replied), check_ms: leanElapsedMs(replied) })
      if (checked.status) return { status: checked.status, ...(checked.resolution ? { resolution: checked.resolution } : {}) }
    }
  } catch (error) {
    // Exhausted model output leaves the existing hold in force under these
    // inputs. Runtime/provider interruptions stay retryable after recovery.
    if (!(error instanceof OutputLimitError)) throw error
  }
}

function refreshFactPrompt(directory: string, question: Question) {
  const data = (value: unknown) => JSON.stringify(value).replaceAll("<", "\\u003c").replaceAll(">", "\\u003e").replaceAll("&", "\\u0026")
  return [
    "Reconsider ONE pending factual uncertainty against changed instructions or supplied evidence. Do not answer for the lawyer, invent evidence, or perform the outside task. Preserve the current treatment unless an actual later instruction changes it or supplied evidence satisfies its stated condition. Unrelated updates never cancel a wait or revoke already authorised provisional drafting. A request to obtain or check evidence is not proof it was obtained or checked.",
    "Use targeted source lookup only for this question, preserving qualifications and contrary evidence. supplied means the exact missing fact is now supplied; it is not independent verification. provisional requires the lawyer's explicit permission to continue with uncertainty marked. Evidence cannot create permission to proceed with a fact still unknown. An original provisional choice remains permission only if the later instructions do not supersede it with a hold. Read all later instructions in receipt order; an earlier permission never overrides a later hold. If nothing establishes a change, keep the current treatment; never infer approval or silently weaken a release restriction.",
    tagged("original_fact_question", data({ question: question.question, context: question.context, answer: question.answer, prior_status: factDisposition(question), prior_resolution: question.fact_resolution })),
    tagged("current_request", data(readLeanText(directory, "request.md") ?? "")),
    tagged("later_instructions", data(laterFactInstructions(directory, question))),
    tagged("source_names", data(matterSourceIDs(directory))),
    'Reply JSON only: {"fact_status":"supplied|provisional|wait|unknown","reason":"the actual effect of the current instructions or evidence","basis":[{"source_id":"request, updated-notes, original-answer or an exact matter source name","quote":"exact relevant words"}]}. A change of treatment requires a checked supporting quotation. Keeping the same treatment needs no new approval and basis may be empty. A new provisional permission or human hold must quote later_instructions under updated-notes. To reestablish express provisional permission in the original typed reply or code-offered choice after an unrelated update, quote that original permission under original-answer and explain why no later instruction supersedes it; never do so while a later hold remains unsatisfied. original-answer supplies no missing fact. The original request predates the answer and a matter source alone supplies no human approval. Later instructions contain only timestamped updates received after the original answer; missing chronology cannot establish an override.',
  ].join("\n\n")
}

function checkFactRefresh(directory: string, question: Question, reply: string): { issues: string[]; status?: FactStatus; resolution?: Question["fact_resolution"] } {
  const parsed = parseJsonObject(reply)
  if (!parsed.ok) return { issues: [`Reply with one JSON object. ${parsed.error}`] }
  const value = parsed.value
  if (!(FACT_STATUSES as readonly unknown[]).includes(value.fact_status) || typeof value.reason !== "string" || !value.reason.trim() || !Array.isArray(value.basis))
    return { issues: ["Give fact_status, the reason preserving or changing the hold, and a basis array."] }
  const allowed = new Set([...matterSourceIDs(directory), "request", "updated-notes", "original-answer"])
  const basis: { source_id: string; quote: string }[] = []
  for (const row of value.basis) {
    if (!row || typeof row !== "object" || typeof row.source_id !== "string" || !allowed.has(row.source_id) || typeof row.quote !== "string" || !row.quote.trim())
      return { issues: ["Every basis needs a known source_id and exact supporting quote."] }
    const originalPermission = !question.answer_kind || questionOptions(question).find(option => option.id === question.option_id)?.fact_action === "provisional"
    if (row.source_id === "original-answer" && (value.fact_status !== "provisional" || !originalPermission))
      return { issues: ["original-answer can establish only express permission in the lawyer's own reply or original code-offered provisional choice, never a missing fact or new override."] }
    const found = row.source_id === "original-answer" ? (question.answer ?? "").includes(row.quote) : row.source_id === "request" || row.source_id === "updated-notes"
      ? (row.source_id === "request" ? readLeanText(directory, "request.md") ?? "" : laterFactInstructions(directory, question)).includes(row.quote)
      : quoteFound(directory, row)
    if (!found) return { issues: [`The supporting quote is not in ${row.source_id}.`] }
    basis.push({ source_id: row.source_id, quote: row.quote })
  }
  const status = value.fact_status as FactStatus
  const prior = factDisposition(question)
  if ((status === "supplied" || status !== prior) && !basis.length)
    return { issues: ["Releasing a wait requires actual quoted evidence or a later human instruction."] }
  if ((status === "provisional" && prior !== "provisional" || status === "wait" && prior === "provisional") && !basis.some(row => row.source_id === "updated-notes" || status === "provisional" && row.source_id === "original-answer"))
    return { issues: ["Provisional drafting requires quoted later human permission, not the original request or a source-derived assumption."] }
  if ((status === "wait" || status === "unknown") && factDisposition(question) === "wait") return { issues: [], status: "wait" }
  return { issues: [], status, ...(basis.length ? { resolution: { reason: value.reason.trim(), basis } } : {}) }
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
