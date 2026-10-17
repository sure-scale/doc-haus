import {
  checklistFindingsCount,
  leanRunStopped,
  openQuestions,
  unresolvedChecklistFindings,
  unresolvedFindingsPhrase,
} from "../drafting/model"
import { leanDeliveredDocuments, type LeanDraftingStatus } from "../../api/ingest/lean"
import {
  activeStep,
  clip,
  stageMilestones,
  stepActivityLabel,
  stoppedByTechnicalProblem,
  type ActivityLine,
} from "./activity"
import { runOutcome } from "./outcome"
import type { NarrationPlan } from "./schedule"

// The compact excerpt the narrator reads for one update. It carries only what
// the run already shows: the run's state, the checkpoints being reported, the
// latest activity lines, the head of the stage's own written output, the tail
// of its thinking, checker notes, and open questions. Every section is clipped
// and the whole excerpt is capped, so one update stays cheap however long the
// stage has run.

export const EXCERPT_LIMIT = 6_000
const ACTIVITY_LINES = 14
const NOTES_HEAD = 1_800
const THINKING_TAIL = 1_000
const LIST_ITEMS = 4
const ITEM_CHARS = 300

/** What a stage session has written so far: the head of its latest answer text
 * and the tail of its latest thinking. */
export type StageNotes = { text: string; reasoning: string }

/** Share the bounded excerpt across available workers instead of showing only
 * the first worker's output. These are excerpts, not a completeness claim. */
export function stageSessionNotes(notes: StageNotes[]): StageNotes | undefined {
  if (!notes.length) return
  if (notes.length === 1) return notes[0]
  const textLimit = Math.max(1, Math.floor(NOTES_HEAD / notes.length) - 32)
  const thinkingLimit = Math.max(1, Math.floor(THINKING_TAIL / notes.length) - 32)
  return {
    text: notes.map((note, i) => `Session ${i + 1} excerpt:\n${head(note.text, textLimit)}`).join("\n\n"),
    reasoning: notes.map((note, i) => `Session ${i + 1} excerpt:\n${tail(note.reasoning, thinkingLimit)}`).join("\n\n"),
  }
}

export function buildExcerpt(input: {
  status: LeanDraftingStatus
  plan: NarrationPlan
  /** Activity lines seen since the last update, oldest first. */
  activity: ActivityLine[]
  notes?: StageNotes
  now: number
}) {
  const status = input.status
  const step = activeStep(status)
  const questions = openQuestions(status).slice(0, LIST_ITEMS)
  const openItems = status.open_items.filter((item) => item.status === "open")
  const unresolved = unresolvedChecklistFindings(status)
  const feedback = Object.values(status.stages)
    .flatMap((row) => row.feedback ?? [])
    .slice(-LIST_ITEMS)
  const sections = [
    section("run", [
      `Run state: ${runState(status)}.`,
      step ? `Working on: ${stepActivityLabel(status, step)}.` : "",
      status.plan?.documents.length
        ? `Planned documents: ${status.plan.documents.map((row) => row.title).join("; ")}.`
        : "",
      leanDeliveredDocuments(status).length ? `Delivered documents: ${leanDeliveredDocuments(status).join("; ")}.` : "",
      `Time since the run started: ${minutesSince(status.created_at, input.now)}.`,
    ]),
    section(
      "update_for",
      input.plan.keys.map((key) => describeKey(status, key)),
    ),
    section(
      "recent_activity",
      input.activity.slice(-ACTIVITY_LINES).map((line) => `- ${clip(line.text, 160)}`),
    ),
    section("stage_output_so_far", [input.notes?.text ? head(input.notes.text, NOTES_HEAD) : ""]),
    section("stage_thinking_tail", [input.notes?.reasoning ? tail(input.notes.reasoning, THINKING_TAIL) : ""]),
    section(
      "checker_notes",
      feedback.map((note) => `- ${clip(note, ITEM_CHARS)}`),
    ),
    section(
      "open_questions_for_the_lawyer",
      questions.map(
        (question) =>
          `- ${clip(question.question, ITEM_CHARS)}${question.context ? ` (context: ${clip(question.context, ITEM_CHARS)})` : ""}`,
      ),
    ),
    section(
      "open_items",
      openItems.length
        ? [
            `${openItems.length} open.`,
            ...openItems.slice(0, LIST_ITEMS).map((item) => `- ${clip(item.detail, ITEM_CHARS)}`),
          ]
        : [],
    ),
    section(
      "unresolved_checklist_findings",
      unresolved.length
        ? [
            `${unresolved.length} not satisfied.`,
            ...unresolved.slice(0, LIST_ITEMS).map((item) => `- ${clip(item.detail, ITEM_CHARS)}`),
          ]
        : [],
    ),
    section("stop_reason", [status.stop_reason ? clip(status.stop_reason, 400) : ""]),
  ]
  return head(
    [
      "Write the next progress update for the lawyer from this excerpt of the drafting run. Say only what is new since your last update.",
      ...sections.filter(Boolean),
    ].join("\n\n"),
    EXCERPT_LIMIT,
  )
}

function runState(status: LeanDraftingStatus) {
  if (status.status === "running") return status.in_flight ? "working" : "paused, nothing is working on it"
  if (status.status === "blocked") return "stopped, waiting on the lawyer"
  if (stoppedByTechnicalProblem(status)) return "paused after a technical problem; the work so far is kept"
  if (leanRunStopped(status)) return "stopped by the lawyer; a new drafting request starts a new draft"
  if (status.status === "failed") return "stopped before finishing"
  const findings = checklistFindingsCount(status)
  if (status.status === "completed" && findings !== undefined)
    return `complete as a review draft with ${unresolvedFindingsPhrase(findings)}; the draft does not satisfy ${findings === 1 ? "that requirement" : "those requirements"}`
  return "complete"
}

/** A checkpoint key in words, from the same status that produced it. */
function describeKey(status: LeanDraftingStatus, key: string) {
  if (key.startsWith("plan:")) return "- The drafting plan is ready."
  if (key.startsWith("question:")) {
    const question = status.questions.find((row) => `question:${row.id}` === key)
    return `- A question for the lawyer was raised: ${clip(question?.question ?? "", ITEM_CHARS)}`
  }
  if (key.startsWith("run:")) return `- The run settled: ${runOutcome(status)?.text ?? runState(status)}`
  const step = key.split(":")[0]!
  if (key.includes(":progress:")) return `- Progress during: ${stepActivityLabel(status, step)}.`
  const finished = stageMilestones(status)
    .filter((line) => line.step === step && line.state !== "running")
    .at(-1)
  return `- ${finished?.text ?? stepActivityLabel(status, step)}.`
}

function section(name: string, lines: string[]) {
  const body = lines.filter(Boolean).join("\n")
  return body ? `<${name}>\n${body}\n</${name}>` : ""
}

function minutesSince(from: string, now: number) {
  const minutes = Math.max(0, Math.round((now - Date.parse(from)) / 60_000))
  return minutes === 1 ? "1 minute" : `${minutes} minutes`
}

export function head(text: string, max: number) {
  return text.length <= max ? text : `${text.slice(0, max - 1)}…`
}

export function tail(text: string, max: number) {
  return text.length <= max ? text : `…${text.slice(text.length - max + 1)}`
}
