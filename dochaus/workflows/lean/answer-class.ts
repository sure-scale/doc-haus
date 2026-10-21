import { parseJsonObject } from "./change-checks"
import { parallel } from "./checklist-flow"
import { appendLeanEvent, leanElapsedMs } from "./events"
import { retryPrompt, tagged } from "./prompts"
import { LEAN_CHILD_PERMISSIONS, LEAN_MAX_RETRIES, LEAN_STAGE_AGENTS, type LeanDeps } from "./stages"
import { ANSWER_CLASSES, questionOptions, readQuestions, recordAnswerClasses, unclassifiedAnswer, type AnswerClass, type LeanState, type Question } from "./state"

// How the lawyer's own words answer a question, read once when the coordinator
// takes the answer and before any step reads it: one fresh read-only session
// per answer, with the same-session bounded retry of every lean job. The reply
// is one of a closed set code accepts (answered, delegated, declined); a
// delegating answer is then recorded as Decide for me records an adoption
// (state.ts classifiedAnswer). A reply that never passes its check stands as
// the lawyer's own answer, as every typed answer did before; a runtime error or
// a pause records nothing, so the next pass reads the answer again.

/** Reads every answer in the lawyer's own words not yet read and records each class. */
export async function classifyAnswers(directory: string, state: LeanState, deps: LeanDeps) {
  const unread = readQuestions(directory).filter(unclassifiedAnswer)
  if (!unread.length) return
  const read = new Map<string, { words: string; answer_class: AnswerClass }>()
  await parallel(unread, async question => {
    const answerClass = await classifyAnswer(directory, state, deps, question)
    if (answerClass) read.set(question.id, { words: question.answer!, answer_class: answerClass })
  })
  if (deps.pause?.aborted) return
  recordAnswerClasses(directory, read)
}

async function classifyAnswer(directory: string, state: LeanState, deps: LeanDeps, question: Question): Promise<AnswerClass | undefined> {
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
      const checked = checkAnswerClass(reply)
      issues = checked.issues
      appendLeanEvent(directory, { ...at, event: "checker", attempt: index, issues, wait_ms: leanElapsedMs(asked, answered), check_ms: leanElapsedMs(answered) })
      if (checked.answer_class) return checked.answer_class
    }
    return "answered"
  } catch {
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
  ].join("\n\n")
}

/** The class a reply gives, or the issue for a retry. */
export function checkAnswerClass(reply: string): { issues: string[]; answer_class?: AnswerClass } {
  const parsed = parseJsonObject(reply)
  if (!parsed.ok) return { issues: [`Reply with one JSON object: ${parsed.error}.`] }
  const value = parsed.value.class
  return (ANSWER_CLASSES as readonly unknown[]).includes(value)
    ? { issues: [], answer_class: value as AnswerClass }
    : { issues: ['"class" must be "answered", "delegated" or "declined".'] }
}
