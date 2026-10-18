import { createHash } from "node:crypto"
import { copyFileSync, existsSync, mkdirSync, readFileSync, realpathSync, rmSync } from "node:fs"
import path from "node:path"
import { readDurableJson, writeDurableBytes } from "../../lib/durable-json"
import type { ChecklistFindingClass } from "./checklist-classify"
import type { DocID, StepID } from "./plan"

// Lean drafting run state. One JSON file, written whole (tmp + rename) on every
// change. It records where each stage stands and the content hash of the inputs
// it last completed on, so a resume skips unchanged work. It is not a receipt,
// claim or settlement: nothing here grants authority to any other component.
// Stages are keyed by step (document and stage, see plan.ts) in walk order.

export const REVIEW_STAGES = ["review", "review-independent", "review-wording", "review-client-disadvantage"] as const
export type ReviewStageID = (typeof REVIEW_STAGES)[number] | "review-playbook"

export function isReviewStage(stage: StageID): stage is ReviewStageID {
  return stage === "review-playbook" || (REVIEW_STAGES as readonly string[]).includes(stage)
}

export const LEAN_STAGES = ["discrepancy", "intake", "research", "research-review", "spine", "draft", ...REVIEW_STAGES, "review-playbook", "change-plan", "change-prepare", "changes", "change-review", "apply", "finalize", "checklist-check", "package-review", "gate"] as const
export type StageID = (typeof LEAN_STAGES)[number]

export type LeanModel = { providerID: string; modelID: string }

export type LeanStageState = {
  status: "pending" | "running" | "done" | "failed" | "blocked"
  input_sha: string | null
  outputs: string[]
  session_id: string | null
  /** Native child sessions of parallel preparation groups, for progress display. */
  session_ids?: string[]
  /** Tries of the stage's hardest step: 1 is the first try, each retry adds one. */
  attempts: number
  /** Model calls the stage made in all; the sectioned draft makes one per step and retry. */
  calls: number
  feedback: string[]
  started_at: string | null
  finished_at: string | null
}

export type LeanState = {
  schema: "dochaus.lean.state/2"
  profile: "lean"
  revision: number
  status: "running" | "blocked" | "failed" | "completed"
  stop_reason: string | null
  /** Set while the run is blocked only because the lawyer paused it. Continue
   * then simply resumes: the pause is not a settled gate or checklist verdict. */
  paused_at?: string
  owner_session_id: string
  model: LeanModel | null
  created_at: string
  updated_at: string
  /** Keyed by step in walk order; a planned document's steps appear once the plan does. */
  stages: Partial<Record<StepID, LeanStageState>>
}

export type OpenItemKind = "research_gap" | "quote_miss" | "citation_unsupported" | "placeholder" | "unapplied_amendment" | "question" | "omission" | "template_removal" | "checklist_open" | "checklist_unresolved"

export type OpenItem = {
  id: string
  stage: StageID
  kind: OpenItemKind
  detail: string
  ref: {
    amendment_id?: string
    pid?: string
    source_id?: string
    quote?: string
    placeholder?: string
    citation?: string
    repair_target?: "body" | "memo" | "research"
    question_id?: string
    /** The template part an omission or removal item is about, as its heading reads in the template. */
    template_part?: string
    /** The planned document whose body the item is about. */
    document?: DocID
    /** The current checklist item a checklist_open or checklist_unresolved
     * item discloses; for a refined parent's coverage, the parent item. */
    checklist_item?: string
    /** Why a checklist_unresolved finding remains adverse; absent for a parent's coverage. */
    checklist_class?: ChecklistFindingClass
  }
  status: "open" | "resolved" | "disclosed"
  resolution?: string
  /** What the applier reported for an unapplied amendment; not part of the id. */
  applier_detail?: Record<string, string>
}

export type Question = {
  id: string
  stage: StageID
  question: string
  context: string
  /** The draft paragraph the question affects, when the asking stage named one. */
  pid?: string
  /** The document of that paragraph: the latest step that asked the question. */
  document?: DocID
  status: "open" | "answered"
  answer?: string
  /** Set when the lawyer handed the question back for the drafting to decide:
   * it is then answered with no answer text, as a choice the lawyer has not made. */
  answer_kind?: "decide_for_me"
  answered_at?: string
}

export const LEAN_STATE_SCHEMA = "dochaus.lean.state/2"
const FIRST_STEPS: StepID[] = ["discrepancy", "intake", "research", "research-review", "spine", "d1.draft", "d1.review", "d1.review-independent", "d1.review-wording", "d1.review-client-disadvantage", "d1.change-plan", "d1.change-prepare", "d1.changes", "d1.change-review", "d1.apply", "finalize", "checklist-check", "package-review", "gate"]

/** Runtime error text recorded in the run (stop reasons, events, apply results)
 * names matter files relative to the matter folder, so the lean state holds no
 * machine path and a moved or cloned matter resumes as it is. */
export function matterRelative(directory: string, text: string) {
  return [...new Set([realpathSync(directory), path.resolve(directory)])]
    .toSorted((a, b) => b.length - a.length)
    .reduce((out, root) => out.replaceAll(`${root}${path.sep}`, "").replaceAll(root, "."), text)
}

export function leanDir(directory: string) {
  return path.join(directory, ".dochaus", "lean")
}

export function leanFile(directory: string, name: string) {
  return path.join(leanDir(directory), name)
}

export function readLeanState(directory: string): LeanState | undefined {
  const value = readDurableJson(leanFile(directory, "state.json"))
  if (!record(value) || value.schema !== LEAN_STATE_SCHEMA || value.profile !== "lean" || !record(value.stages)) return
  return value as LeanState
}

export function leanStateExists(directory: string) {
  return existsSync(leanFile(directory, "state.json"))
}

export function writeLeanState(directory: string, state: LeanState): LeanState {
  const next = { ...state, revision: state.revision + 1, updated_at: new Date().toISOString() }
  writeLeanText(directory, "state.json", `${JSON.stringify(next, null, 2)}\n`)
  return next
}

export function updateLeanStage(directory: string, state: LeanState, step: StepID, patch: Partial<LeanStageState>) {
  return writeLeanState(directory, { ...state, stages: { ...state.stages, [step]: { ...(state.stages[step] ?? emptyStage()), ...patch } } })
}

/** Puts the state's step rows in the walk's order, adding pending rows for new
 * steps and dropping steps no longer walked, in one write. Unchanged order
 * writes nothing. */
export function syncLeanSteps(directory: string, state: LeanState, steps: StepID[]) {
  if (Object.keys(state.stages).join("\0") === steps.join("\0")) return state
  return writeLeanState(directory, { ...state, stages: Object.fromEntries(steps.map(step => [step, state.stages[step] ?? emptyStage()])) })
}

export function initLeanState(directory: string, input: { request: string; ownerSessionID: string; model: LeanModel | null }) {
  mkdirSync(leanDir(directory), { recursive: true, mode: 0o700 })
  writeLeanText(directory, "request.md", `${input.request.replace(/\r\n/g, "\n").trim()}\n`)
  const now = new Date().toISOString()
  return writeLeanState(directory, {
    schema: LEAN_STATE_SCHEMA,
    profile: "lean",
    revision: 0,
    status: "running",
    stop_reason: null,
    owner_session_id: input.ownerSessionID,
    model: input.model,
    created_at: now,
    updated_at: now,
    stages: Object.fromEntries(FIRST_STEPS.map(step => [step, emptyStage()])),
  })
}

export function writeLeanText(directory: string, name: string, text: string) {
  writeLeanBytes(directory, name, new TextEncoder().encode(text))
}

export function writeLeanBytes(directory: string, name: string, bytes: Uint8Array) {
  mkdirSync(path.dirname(leanFile(directory, name)), { recursive: true, mode: 0o700 })
  writeDurableBytes(leanFile(directory, name), bytes)
}

export function readLeanText(directory: string, name: string) {
  const file = leanFile(directory, name)
  return existsSync(file) ? readFileSync(file, "utf8") : undefined
}

export function readLeanJson(directory: string, name: string): unknown {
  return readDurableJson(leanFile(directory, name))
}

/** Resume key only: identical inputs mean the stage's completed output still
 * answers them. Inputs are named parts so a renamed input changes the key. */
export function hashStageInputs(step: StepID, version: string, parts: { name: string; bytes: string | Uint8Array }[]) {
  const hash = createHash("sha256").update(`${step}\0${version}\0`)
  parts.toSorted((a, b) => a.name.localeCompare(b.name)).forEach(part => {
    hash.update(`${part.name}\0`)
    hash.update(part.bytes)
    hash.update("\0")
  })
  return hash.digest("hex")
}

/** Plain copy of a completed step's outputs, kept for the lawyer to inspect or
 * restore. Matters are not git repositories, so native step snapshots do not
 * cover them. A step's outputs have distinct file names. */
export function snapshotStage(directory: string, revision: number, step: StepID, outputs: string[]) {
  const target = path.join(leanDir(directory), "snapshots", `${String(revision).padStart(4, "0")}-${step}`)
  rmSync(target, { recursive: true, force: true })
  mkdirSync(target, { recursive: true, mode: 0o700 })
  outputs.filter(name => existsSync(leanFile(directory, name)))
    .forEach(name => copyFileSync(leanFile(directory, name), path.join(target, path.basename(name))))
  return path.relative(leanDir(directory), target)
}

export function readOpenItems(directory: string): OpenItem[] {
  const value = readLeanJson(directory, "open-items.json")
  return record(value) && Array.isArray(value.items) ? value.items.filter(isOpenItem) : []
}

/** The run's open items as they stand for the lawyer, each finding once. A
 * document's draft and review findings are found in its draft body and are
 * inputs of its changes step; once its apply step has checked the final body,
 * that check's items are the document's body findings and the earlier ones are
 * superseded: this list leaves them out, so the gate neither shows nor settles
 * them. open-items.json keeps them for the changes step; what the lawyer is
 * shown and the gate settles is this list. */
export function currentOpenItems(directory: string): OpenItem[] {
  const stages = readLeanState(directory)?.stages ?? {}
  return readOpenItems(directory).filter(item => (item.stage !== "draft" && !isReviewStage(item.stage)) ||
    stages[`${item.ref.document ?? "d1"}.apply`]?.status !== "done")
}

/** A step's open items are recomputed whenever it runs, so they replace that
 * step's previous items. Items from other steps are kept. */
export function replaceStageOpenItems(directory: string, stage: StageID, document: DocID | null, items: OpenItem[]) {
  const kept = readOpenItems(directory).filter(item => item.stage !== stage || (item.ref.document ?? null) !== document)
  const unique = [...new Map(items.map(item => [item.id, item])).values()]
  writeOpenItems(directory, [...kept, ...unique])
}

export function writeOpenItems(directory: string, items: OpenItem[]) {
  writeLeanText(directory, "open-items.json", `${JSON.stringify({ items }, null, 2)}\n`)
}

export function readQuestions(directory: string): Question[] {
  const value = readLeanJson(directory, "questions.json")
  return record(value) && Array.isArray(value.questions) ? value.questions.filter(isQuestion) : []
}

export function writeQuestions(directory: string, questions: Question[]) {
  writeLeanText(directory, "questions.json", `${JSON.stringify({ questions }, null, 2)}\n`)
}

/** Questions a step asks replace its earlier open questions. Answered
 * questions are kept: a human answer is never discarded or re-asked. A question
 * asked again by another step (same id) belongs to the latest asker; when that
 * asker stops asking it, it stays with an earlier asker whose open item still
 * refers to it. Call after the step's open items are replaced. */
export function replaceStageQuestions(directory: string, stage: StageID, document: DocID | null, asked: Question[]) {
  const existing = readQuestions(directory)
  const answered = existing.filter(question => question.status === "answered")
  const answeredIDs = new Set(answered.map(question => question.id))
  const fresh = [...new Map(asked.filter(question => !answeredIDs.has(question.id)).map(question => [question.id, question])).values()]
  const others = keepAskedQuestions(
    existing.filter(question => question.status === "open" && !fresh.some(item => item.id === question.id)),
    readOpenItems(directory),
    question => question.stage !== stage || (question.document ?? null) !== document,
  )
  writeQuestions(directory, [...answered, ...others, ...fresh])
}

/** Keeps the questions `kept` accepts and every answered one. Any other open
 * question stays while an open item still asks it, and passes to that item's
 * step, document and paragraph, so the item keeps a question its answer can reach. */
export function keepAskedQuestions(questions: Question[], items: OpenItem[], kept: (question: Question) => boolean): Question[] {
  return questions.flatMap(question => {
    if (question.status === "answered" || kept(question)) return [question]
    const asker = items.findLast(item => item.kind === "question" && item.ref.question_id === question.id)
    if (!asker) return []
    return [{
      id: question.id, stage: asker.stage, question: question.question, context: question.context,
      ...(asker.ref.pid ? { pid: asker.ref.pid } : {}), ...(asker.ref.document ? { document: asker.ref.document } : {}), status: question.status,
    }]
  })
}

export class LeanQuestionError extends Error {
  constructor(message: string, readonly code: "not_found" | "invalid_answer" | "already_answered" | "not_delegable") {
    super(message)
  }
}

/** Canonical answer text; throws when it is empty or too long. */
export function validateAnswerText(answer: string) {
  const text = answer.replace(/\r\n/g, "\n").trim()
  if (!text || text.length > 20_000) throw new LeanQuestionError("An answer must be non-empty text of at most 20,000 characters.", "invalid_answer")
  return text
}

export function answerQuestion(directory: string, questionID: string, answer: string) {
  const text = validateAnswerText(answer)
  const questions = readQuestions(directory)
  const target = questions.find(question => question.id === questionID)
  if (!target) throw new LeanQuestionError("No such question in this matter's lean drafting run.", "not_found")
  // A real answer replaces an earlier hand-back.
  const { answer_kind: _handedBack, ...asked } = target
  const answered: Question = { ...asked, status: "answered", answer: text, answered_at: new Date().toISOString() }
  writeQuestions(directory, questions.map(question => question.id === questionID ? answered : question))
  return answered
}

/** The lawyer hands an open model-asked question back for the drafting to
 * decide. Keyed questions (package scope, requested files, checklist conflict)
 * are read as settled once answered, so they cannot be handed back. */
export function delegateQuestion(directory: string, questionID: string) {
  const questions = readQuestions(directory)
  const target = questions.find(question => question.id === questionID)
  if (!target) throw new LeanQuestionError("No such question in this matter's lean drafting run.", "not_found")
  if (!questionID.startsWith("q-")) throw new LeanQuestionError("This question needs the lawyer's own answer; it cannot be left to the drafting to decide.", "not_delegable")
  if (target.status === "answered") throw new LeanQuestionError("This question is already answered.", "already_answered")
  const { answer: _answer, ...asked } = target
  const delegated: Question = { ...asked, status: "answered", answer_kind: "decide_for_me", answered_at: new Date().toISOString() }
  writeQuestions(directory, questions.map(question => question.id === questionID ? delegated : question))
  return delegated
}

/** A question the lawyer left to the drafting to decide and has not since answered. */
export function handedBack(question: Question) {
  return question.status === "answered" && question.answer_kind === "decide_for_me"
}

/** Every answered question, ordered by id: questions.json's own order moves
 * when a step's questions replace its earlier ones. */
export function answeredQuestions(directory: string) {
  return readQuestions(directory).filter(question => question.status === "answered").toSorted((a, b) => a.id.localeCompare(b.id))
}

/** Canonical text of every lawyer's answer; an input to every drafting stage.
 * A question handed back has no answer: handedBackQuestionsText lists it. */
export function answeredQuestionsText(directory: string) {
  return answeredQuestions(directory).filter(question => !handedBack(question))
    .map(question => `- [${question.id}] Q: ${question.question}\n  A: ${question.answer}`)
    .join("\n")
}

/** Canonical text of every question the lawyer handed back, by id like the answers. */
export function handedBackQuestionsText(directory: string) {
  return answeredQuestions(directory).filter(handedBack).map(question => `- [${question.id}] Q: ${question.question}`).join("\n")
}

/** The hand-backs as prompts show them beside the lawyer's answers, or "" with
 * none: choices the drafting makes and discloses, never the lawyer's answers. */
export function handedBackBlock(directory: string) {
  const text = handedBackQuestionsText(directory)
  return text ? `The lawyer left these questions to the drafting to decide instead of answering them; do not ask them again. Each is a choice the drafting makes on professional judgment and discloses as one the lawyer has not made.\n${text}` : ""
}

/** The "answers" input of every step's key. A hand-back keys it as an answer
 * does, so it reruns from intake; without one it is the answers text alone. */
export function answersInput(directory: string) {
  const text = handedBackQuestionsText(directory)
  return text ? [answeredQuestionsText(directory), `Left to the drafting:\n${text}`].filter(Boolean).join("\n") : answeredQuestionsText(directory)
}

export function shortHash(text: string, length = 10) {
  return createHash("sha1").update(text).digest("hex").slice(0, length)
}

function emptyStage(): LeanStageState {
  return { status: "pending", input_sha: null, outputs: [], session_id: null, attempts: 0, calls: 0, feedback: [], started_at: null, finished_at: null }
}

function isOpenItem(value: unknown): value is OpenItem {
  return record(value) && typeof value.id === "string" && typeof value.kind === "string" && typeof value.stage === "string" &&
    typeof value.detail === "string" && record(value.ref)
}

function isQuestion(value: unknown): value is Question {
  return record(value) && typeof value.id === "string" && typeof value.question === "string" &&
    (value.status === "open" || value.status === "answered")
}

function record(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value)
}
