import { adviseQuestions } from "./advise"
import { parseJsonObject, record } from "./checks"
import { appendLeanEvent, leanElapsedMs } from "./events"
import { OutputLimitError } from "./native-stage-result"
import { keyedQuestionID } from "./package-scope"
import { readLeanPlan } from "./plan"
import { retryPrompt, tagged } from "./prompts"
import { excerpt } from "./quotes"
import { LEAN_CHILD_PERMISSIONS, LEAN_MAX_RETRIES, LEAN_STAGE_AGENTS, type LeanDeps } from "./stages"
import { answeredQuestions, readLeanJson, readLeanText, shortHash, undecidedQuestion, writeLeanText, type LeanState, type Question } from "./state"

// Whether the lawyer's answers can all be followed together, read once for each
// set of answers and before any step reads them. Each answer, and each adopted
// recommendation, was given on its own question, so nothing else sees them side
// by side. One fresh read-only session reads the whole set, each answer with
// the document and paragraph it was asked about, with the same-session bounded
// retry of every lean job; code accepts a conflict only between two or more
// answered questions it names. Each conflict becomes one open question for the
// lawyer, advised as every asked question is, whose answer settles the point in
// place of the answers it names (`replaces`). The result is kept under the
// answers' content key, so an unchanged set is not read again; a reply that
// never passes its check finds no conflict, and a runtime error or a pause
// records nothing, so the next pass reads the set again.

const RECONCILE_FILE = "answer-reconcile.json"
/** Keys the saved result: a change to the job's prompt or to what it reads
 * reads every set of answers again. */
const RECONCILE_CONTRACT = "answer-reconcile-2"
const CONFLICT_PREFIX = "answer-conflict-"
/** The closing sentence of a conflict question, quoting the questions whose answers it replaces. */
const REPLACES = "This answer settles the point in place of the answers to"
const QUOTED = 160

type AnswerConflict = { questions: string[]; conflict: string }

/** The open questions to ask about answers that cannot all be followed
 * together, each advised; none while fewer than two questions are answered or
 * when the answers agree. The caller records them. */
export async function reconcileAnswers(directory: string, state: LeanState, deps: LeanDeps): Promise<Question[]> {
  const answered = answersToReconcile(directory)
  if (answered.length < 2) return []
  const request = readLeanText(directory, "request.md") ?? ""
  const updates = readLeanText(directory, "updated-notes.md") ?? ""
  const titles = new Map((readLeanPlan(directory)?.documents ?? []).map(row => [row.id, row.title]))
  const input_sha = shortHash(JSON.stringify([RECONCILE_CONTRACT, request, updates, answered.map(question => answerRow(question, titles))]), 40)
  const saved = readLeanJson(directory, RECONCILE_FILE)
  let conflicts = record(saved) && saved.input_sha === input_sha && Array.isArray(saved.conflicts) ? saved.conflicts as AnswerConflict[] : undefined
  if (!conflicts) {
    conflicts = await findConflicts(directory, state, deps, answered, request, updates, titles)
    if (!conflicts || deps.pause?.aborted) return []
    writeLeanText(directory, RECONCILE_FILE, `${JSON.stringify({ input_sha, conflicts }, null, 2)}\n`)
  }
  const asked = conflicts.filter(conflict => !settled(directory, conflict)).map(conflict => conflictQuestion(answered, conflict))
  return asked.length ? adviseQuestions(directory, state, "intake", deps, asked) : []
}

/** The answers the job reads together. An answer that still needs the
 * lawyer's own words decides nothing to reconcile. */
export function answersToReconcile(directory: string) {
  return answeredQuestions(directory).filter(question => !undecidedQuestion(question))
}

/** The conflicts the job finds, none when its retries run out or its reply is
 * cut off at the output limit, or undefined on any other runtime error, a
 * pause or a settings change, which the runner acts on. */
async function findConflicts(directory: string, state: LeanState, deps: LeanDeps, answered: Question[], request: string, updates: string, titles: ReadonlyMap<string, string>): Promise<AnswerConflict[] | undefined> {
  const agent = LEAN_STAGE_AGENTS.reconcile
  const at = { stage: "intake" as const, step: "reconcile answers" }
  const known = new Set(answered.map(question => question.id))
  try {
    const sessionID = await deps.sessions.createSession({ directory, parentID: state.owner_session_id, agent, title: "Lean drafting: reconcile answers", model: state.model, permissions: LEAN_CHILD_PERMISSIONS })
    let issues: string[] = []
    for (let index = 0; index <= LEAN_MAX_RETRIES; index += 1) {
      appendLeanEvent(directory, { ...at, event: index ? "retry" : "model_call", attempt: index, session_id: sessionID })
      const asked = performance.now()
      const reply = await deps.sessions.prompt({ directory, sessionID, agent, text: index ? retryPrompt("reconcile", issues) : reconcilePrompt(answered, request, updates, titles) })
      const checkedAt = performance.now()
      const checked = checkConflicts(reply, known)
      issues = checked.issues
      appendLeanEvent(directory, { ...at, event: "checker", attempt: index, issues, wait_ms: leanElapsedMs(asked, checkedAt), check_ms: leanElapsedMs(checkedAt) })
      if (checked.conflicts) return checked.conflicts
    }
    return []
  } catch (error) {
    // A reply cut off before any answer finds no conflict, like one that never passes its check.
    if (error instanceof OutputLimitError) return []
    // A runtime error, a pause or a settings change: the runner acts on it.
    return undefined
  }
}

/** The job's prompt; `titles` names each planned document, for what every
 * answer was asked about. */
export function reconcilePrompt(answered: Question[], request: string, updates: string, titles: ReadonlyMap<string, string>) {
  return [
    "Read the lawyer's answers to the questions raised while drafting together and say which of them, if any, cannot all be followed. Do not answer, choose between or change any answer, and do not look anything up.",
    "Answers conflict when following one of them means not following another. Answers on different points do not conflict, and an answer that narrows, adds a condition to or fills in the detail of another does not conflict with it. Answers asked about different paragraphs are different points unless following one makes the other impossible. An answer that settles a point in place of earlier answers governs over them on that point.",
    tagged("request", request),
    ...(updates ? [tagged("user_updates", updates)] : []),
    tagged("answers", answered.map(question => answerText(question, titles)).join("\n\n")),
    'Reply with one JSON object only: {"conflicts":[{"questions":["question id","question id"],"conflict":"what cannot be followed together, in one or two sentences"}]}, one entry for each point, naming at least two question ids exactly as given above; or {"conflicts":[]} when the answers can all be followed together.',
  ].join("\n\n")
}

/** The conflicts a reply names, merged by the questions they name, or the
 * issues for a retry. */
export function checkConflicts(reply: string, known: ReadonlySet<string>): { issues: string[]; conflicts?: AnswerConflict[] } {
  const parsed = parseJsonObject(reply)
  if (!parsed.ok) return { issues: [`Reply with one JSON object. ${parsed.error}`] }
  const { conflicts } = parsed.value
  if (!Array.isArray(conflicts)) return { issues: ['"conflicts" must be a list, empty when the answers can all be followed together.'] }
  const issues: string[] = []
  const merged = new Map<string, AnswerConflict>()
  conflicts.forEach((row, index) => {
    const at = `conflicts[${index}]`
    if (!record(row) || !Array.isArray(row.questions)) return void issues.push(`${at} must be {"questions":["question id",...],"conflict":"..."}.`)
    const ids = [...new Set(row.questions.map(id => typeof id === "string" ? id.trim() : id))]
    const unknown = ids.filter(id => typeof id !== "string" || !known.has(id))
    if (unknown.length) issues.push(`${at}.questions names ${unknown.map(id => JSON.stringify(id)).join(", ")}, not the id of an answered question given above.`)
    else if (ids.length < 2) issues.push(`${at}.questions must name at least two different answered questions.`)
    const words = typeof row.conflict === "string" ? row.conflict.trim() : ""
    const text = !words || /[.!?]$/.test(words) ? words : `${words}.`
    if (!text) issues.push(`${at}.conflict must say what cannot be followed together.`)
    if (unknown.length || ids.length < 2 || !text) return
    const questions = (ids as string[]).toSorted()
    const key = JSON.stringify(questions)
    const earlier = merged.get(key)
    merged.set(key, { questions, conflict: earlier ? `${earlier.conflict} ${text}` : text })
  })
  return issues.length ? { issues } : { issues: [], conflicts: [...merged.values()] }
}

/** Whether an answer the lawyer gave to a conflict question already settles
 * `conflict`: it names that answer, or only answers that answer replaces. */
function settled(directory: string, conflict: AnswerConflict) {
  const combined = answeredQuestions(directory).filter(question => question.id.startsWith(CONFLICT_PREFIX))
  return combined.some(question => conflict.questions.includes(question.id) || conflict.questions.every(id => question.replaces?.includes(id)))
}

/** One open question asking the lawyer for one approach on the point the
 * answers conflict on. Its closing sentence quotes the questions whose
 * answers it replaces, so every step reads it beside them; `replaces` carries
 * their ids for settled. It is the document's own question when every answer
 * it names is that document's, so only that document's steps read its answer;
 * otherwise it is run-level, as the answers it replaces are. */
export function conflictQuestion(answered: Question[], conflict: AnswerConflict): Question {
  const rows = conflict.questions.map(id => answered.find(question => question.id === id)!)
  const documents = new Set(rows.map(question => question.document))
  const [document] = documents
  return {
    id: keyedQuestionID(CONFLICT_PREFIX, conflict.questions),
    stage: "intake",
    kind: "judgment",
    status: "open",
    question: `Some of your answers cannot all be followed together: ${conflict.conflict} Which one approach should the drafting follow on this point? ${REPLACES} ${quotedList(rows.map(question => excerpt(question.question, QUOTED)))}.`,
    context: [
      "The answers that conflict:",
      ...rows.map(question => `- "${excerpt(question.question, QUOTED)}"\n  ${givenAs(question)}: ${excerpt(question.answer ?? "", QUOTED)}`),
    ].join("\n"),
    ...(documents.size === 1 && document ? { document } : {}),
    replaces: [...conflict.questions],
  }
}

/** "A", "A" and "B", or "A", "B" and "C". */
function quotedList(texts: string[]) {
  const quoted = texts.map(text => `"${text}"`)
  return quoted.length > 1 ? `${quoted.slice(0, -1).join(", ")} and ${quoted.at(-1)}` : quoted.join("")
}

/** What the job's key reads of one answer: the words it is given. */
function answerRow(question: Question, titles: ReadonlyMap<string, string>) {
  return [question.id, question.question, askedAbout(question, titles), question.context, question.answer, question.answer_kind ?? null, question.lawyer_words ?? null]
}

function answerText(question: Question, titles: ReadonlyMap<string, string>) {
  return [
    `[${question.id}] Q: ${question.question}`,
    `  Asked about: ${askedAbout(question, titles)}`,
    ...(question.context ? [`  Context: ${question.context}`] : []),
    `  ${givenAs(question)}: ${question.answer ?? ""}`,
    ...(question.lawyer_words ? [`  The lawyer's own words: ${question.lawyer_words}`] : []),
  ].join("\n")
}

/** What an answer was asked about, from the question's own fields: the
 * document and paragraph the asking step named, or else the whole work. */
function askedAbout(question: Question, titles: ReadonlyMap<string, string>) {
  const where = [
    ...(question.document ? [titles.get(question.document) ?? `document ${question.document}`] : []),
    ...(question.pid ? [`paragraph [${question.pid}]`] : []),
  ]
  return where.length ? where.join(", ") : "the whole work"
}

/** How an answer was given, as the lawyer reads it. */
function givenAs(question: Question) {
  if (question.answer_kind === "adopted") return "Answer (the lawyer delegated the choice; the recommended option was adopted)"
  if (question.answer_kind === "option") return "Answer (an option the lawyer chose)"
  return "Answer (the lawyer's own words)"
}
