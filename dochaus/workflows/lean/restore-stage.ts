import { copyFileSync, existsSync, readdirSync, rmSync } from "node:fs"
import path from "node:path"
import { appendLeanEvent } from "./events"
import { isStepID, leanSteps, stepDocument, stepStage, type StepID } from "./plan"
import { queuedAnswerIDs } from "./requests"
import { LEAN_PROMPT_VERSION, stageInputs, stepOutputs } from "./stages"
import {
  hashStageInputs,
  keepAskedQuestions,
  leanDir,
  leanFile,
  readLeanState,
  readOpenItems,
  readQuestions,
  writeLeanState,
  writeOpenItems,
  writeQuestions,
  type LeanStageState,
  type StageID,
} from "./state"

// Development checkpoint: puts a matter's lean run back to just after step N
// from the per-step snapshots the coordinator keeps, so a later step can be
// re-run without redrafting. Steps up to N get their snapshot outputs back and
// are marked done with input hashes recomputed from the restored files; later
// steps, including every later document's, lose their outputs and become
// pending. The run stops as blocked, and Continue (or runLean) re-runs only the
// later steps.
//
//   bun dochaus/workflows/lean/restore-stage.ts <matter-dir> <step, e.g. d1.apply>

export function restoreLeanStage(directory: string, step: StepID) {
  const state = readLeanState(directory)
  if (!state) throw new Error("This matter has no lean drafting run.")
  const steps = leanSteps(directory)
  const upTo = steps.indexOf(step)
  if (upTo < 0) throw new Error(`Unknown step "${step}". Steps: ${steps.join(", ")}.`)
  const kept = steps.slice(0, upTo + 1)
  const snapshots = chooseSnapshots(directory, kept)
  steps.slice(upTo + 1).flatMap(stepOutputs).forEach(name => rmSync(leanFile(directory, name), { force: true }))
  if (step !== "gate") rmSync(leanFile(directory, "documents.json"), { force: true })
  kept.forEach(row => stepOutputs(row).forEach(name => {
    const from = path.join(snapshots[row]!, path.basename(name))
    if (existsSync(from)) copyFileSync(from, leanFile(directory, name))
  }))
  const keeps = (stage: StageID, document: string | undefined) => kept.some(row => stepStage(row) === stage && (stepDocument(row) ?? undefined) === document)
  writeOpenItems(directory, readOpenItems(directory).filter(item => keeps(item.stage, item.ref.document)))
  // A lawyer's answer is never discarded, including one still waiting in the
  // request queue; other open questions from later steps are, unless a kept
  // step's open item still asks them. Queued requests and repairs.json are left
  // as they are: they are the lawyer's, and the next run drains or consumes them as usual.
  const queued = queuedAnswerIDs(directory)
  writeQuestions(directory, keepAskedQuestions(readQuestions(directory), readOpenItems(directory),
    question => keeps(question.stage, question.document) || queued.has(question.id)))
  const finished = new Date().toISOString()
  // The restored plan fixes the steps from here on.
  const stages = Object.fromEntries(leanSteps(directory).map((row): [StepID, LeanStageState] => [row, kept.includes(row)
    ? {
      ...pending(), ...state.stages[row], status: "done", outputs: stepOutputs(row), finished_at: state.stages[row]?.finished_at ?? finished,
      input_sha: hashStageInputs(row, LEAN_PROMPT_VERSION, stageInputs(directory, row)),
    }
    : pending()]))
  const reason = `restored to just after step ${step}; Continue re-runs the later steps`
  appendLeanEvent(directory, { stage: stepStage(step), document: stepDocument(step) ?? undefined, event: "restored", reason,
    outputs: kept.map(row => path.basename(snapshots[row]!)) })
  return writeLeanState(directory, { ...state, status: "blocked", stop_reason: reason, stages })
}

function pending(): LeanStageState {
  return { status: "pending", input_sha: null, outputs: [], session_id: null, attempts: 0, calls: 0, feedback: [], started_at: null, finished_at: null }
}

/** The latest snapshot of the target step, and for each earlier step its latest
 * snapshot taken before that one, so the restored files belong together. */
function chooseSnapshots(directory: string, steps: StepID[]) {
  const root = path.join(leanDir(directory), "snapshots")
  const all = existsSync(root)
    ? readdirSync(root).flatMap(name => {
      const match = /^(\d+)-(.+)$/.exec(name)
      return match ? [{ name, revision: Number(match[1]), step: match[2] }] : []
    })
    : []
  const target = steps.at(-1)!
  const latest = (step: StepID, before: number) =>
    all.filter(row => row.step === step && row.revision < before).toSorted((a, b) => b.revision - a.revision)[0]
  const anchor = latest(target, Number.POSITIVE_INFINITY)
  if (!anchor) throw new Error(`No snapshot of step "${target}" exists; run the step once first.`)
  return Object.fromEntries(steps.map(step => {
    const chosen = step === target ? anchor : latest(step, anchor.revision)
    if (!chosen) throw new Error(`No snapshot of step "${step}" precedes the ${target} snapshot ${anchor.name}.`)
    return [step, path.join(root, chosen.name)]
  })) as Partial<Record<StepID, string>>
}

if (import.meta.main) {
  const [directory, step] = process.argv.slice(2)
  if (!directory || !step || !isStepID(step)) {
    console.error("usage: bun dochaus/workflows/lean/restore-stage.ts <matter-dir> <step: d1.draft|d1.review|d1.changes|d1.apply|...|finalize|gate>")
    process.exit(2)
  }
  const state = restoreLeanStage(path.resolve(directory), step)
  console.log(`restored: status=${state.status} revision=${state.revision} ` +
    Object.entries(state.stages).map(([row, value]) => `${row}=${value?.status}`).join(" "))
}
