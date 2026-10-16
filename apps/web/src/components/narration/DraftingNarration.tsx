import UserText from "../chat/UserText"
import { activeStep, isWorking, stepActivityLabel } from "./activity"
import { openQuestions } from "../drafting/model"
import { runOutcome } from "./outcome"
import ResultCard from "./ResultCard"
import { narrationTimeline, type TimelineBlock } from "./timeline"
import Thinking from "./Thinking"
import TimelineEntry from "./TimelineEntry"
import type { DraftingNarration as Narration } from "./use-drafting-narration"

// The drafting run, narrated in the matter chat: one row per stage with its
// activity folded beneath, the narrator's plain-English updates between them, a
// live pill while a stage works, and the result once the run settles. The run
// header above the chat carries the controls; this is the running commentary.
export default function DraftingNarration({
  narration,
  matterId,
  showRequest,
  onViewDocument,
}: {
  narration: Narration
  matterId: string
  /** Show the run's recorded request; off where the request is already a turn in this chat. */
  showRequest: boolean
  onViewDocument: (name: string) => void
}) {
  const status = narration.status
  if (!status || !narration.visible) return null
  const working = isWorking(status)
  const blocks = narrationTimeline(narration.lines, narration.notes, working, status.plan?.documents)
  // Fold only consecutive successful mechanical stages. Narrator prose,
  // thoughts, flagged findings and interrupted activity retain their original
  // positions outside the disclosure, including within an otherwise done run.
  const timeline = blocks.reduce<Array<TimelineBlock | TimelineBlock[]>>((rows, block) => {
    const fold = status.status === "completed" && block.kind === "stage" &&
      block.line.state === "done" && !block.line.flag &&
      block.lines.every((line) => line.kind === "tool" && line.state === "done" && !line.flag) &&
      block.rollUp.every((part) => !part.flagged)
    const last = rows.at(-1)
    if (fold && Array.isArray(last)) {
      last.push(block)
      return rows
    }
    rows.push(fold ? [block] : block)
    return rows
  }, [])
  const outcome = runOutcome(status)
  const step = activeStep(status)
  const since = Date.parse((step && status.stages[step]?.started_at) || status.updated_at)
  // What the run is doing this moment: the stage session's newest unfinished
  // read, search or thought, else the stage itself.
  const current = narration.lines
    .filter((line) => line.kind !== "milestone" && line.state === "running" && line.step === step)
    .sort((a, b) => a.at - b.at)
    .at(-1)
  return (
    <>
      {showRequest && status.request && (
        <div className="msg-user-wrap">
          <div className="msg user">
            <UserText text={status.request} />
          </div>
        </div>
      )}
      <div className="msg assistant narration">
        <div className="msg-agent">Drafting</div>
        <ol className="narration-timeline">
          {timeline.map((entry) => Array.isArray(entry) ? (
            <li key={`activity:${blockKey(entry[0]!)}`} className="narration-completed-activity">
              <details>
                <summary>Drafting activity <span>{entry.length} {entry.length === 1 ? "step" : "steps"}</span></summary>
                <ol>
                  {entry.map((block) => <TimelineEntry key={blockKey(block)} block={block} />)}
                </ol>
              </details>
            </li>
          ) : <TimelineEntry key={blockKey(entry)} block={entry} />)}
        </ol>
        {!working && narration.narrating && <div className="narration-muted">Summarising progress…</div>}
        {status.status === "completed" && (
          <ResultCard status={status} matterId={matterId} onViewDocument={onViewDocument} />
        )}
        {outcome && outcome.tone !== "done" && openQuestions(status).length === 0 && (
          <p className={`narration-outcome-note is-${outcome.tone}`} role="status">
            {outcome.text}
          </p>
        )}
        {narration.failed && (
          <div className="narration-muted">
            A progress summary could not be written just now. The steps above are current.
          </div>
        )}
        {working && (
          <Thinking
            label={current?.text ?? (step ? stepActivityLabel(status, step) : "Working")}
            since={since}
            summarising={narration.narrating}
          />
        )}
      </div>
    </>
  )
}

function blockKey(block: TimelineBlock) {
  if (block.kind === "milestone" || block.kind === "stage") return block.line.id
  return `${block.kind}:${block.id}`
}
