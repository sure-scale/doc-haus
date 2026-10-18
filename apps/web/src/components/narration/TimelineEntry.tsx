import { useEffect, useLayoutEffect, useRef, useState } from "react"
import Markdown from "../markdown/Markdown"
import type { ActivityLine } from "./activity"
import { elapsed, type RollUpPart, type StepRow, type TimelineBlock } from "./timeline"

// One entry of the narration timeline: a narrator update (clamped to three
// lines with More), a document heading, a stage row with its activity folded
// beneath it, a run-level line, or a loose group of activity steps.

export default function TimelineEntry({ block }: { block: TimelineBlock }) {
  if (block.kind === "note") return <Note text={block.text} past={block.past} />
  if (block.kind === "heading") return <li className="narration-heading">{block.text}</li>
  if (block.kind === "milestone")
    return (
      <li className={`narration-stage is-${block.line.state}`}>
        <div className="narration-stage-row">
          <StateIcon state={block.line.state} />
          <span className="narration-stage-text">{block.line.text}</span>
        </div>
      </li>
    )
  if (block.kind === "stage") return <Stage block={block} />
  return (
    <li className="narration-steps">
      <Activity id={block.id} lines={block.lines} shown={block.shown} />
    </li>
  )
}

function Stage({ block }: { block: Extract<TimelineBlock, { kind: "stage" }> }) {
  const line = block.line
  return (
    <li className={`narration-stage is-${line.state}`}>
      <div className="narration-stage-row">
        <StateIcon state={line.state} />
        <span className="narration-stage-text">{line.text}</span>
        {line.state === "running" ? <Duration from={line.at} /> : line.until !== undefined && <Duration from={line.at} until={line.until} />}
      </div>
      {block.lines.length > 0 && (
        <div className="narration-stage-body">
          <Activity id={line.id} lines={block.lines} shown={block.shown} rollUp={block.rollUp} />
        </div>
      )}
    </li>
  )
}

/** A stage's own clock: fixed once it has finished, ticking while it runs. */
function Duration({ from, until }: { from: number; until?: number }) {
  const [now, setNow] = useState(Date.now())
  useEffect(() => {
    if (until !== undefined) return
    const timer = setInterval(() => setNow(Date.now()), 1_000)
    return () => clearInterval(timer)
  }, [until])
  if (!Number.isFinite(from)) return null
  return <span className="narration-duration">{elapsed((until ?? now) - from)}</span>
}

function StateIcon({ state }: { state: ActivityLine["state"] }) {
  if (state === "running") return <span className="narration-icon is-running" aria-hidden />
  return (
    <svg className={`narration-icon is-${state}`} width="16" height="16" viewBox="0 0 16 16" aria-hidden>
      {state === "done" && (
        <>
          <circle cx="8" cy="8" r="7" fill="currentColor" />
          <path d="m5 8.2 2 2 4-4.2" fill="none" stroke="var(--text-on-accent)" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" />
        </>
      )}
      {state === "paused" && (
        <>
          <circle cx="8" cy="8" r="6.25" fill="none" stroke="currentColor" strokeWidth="1.5" />
          <path d="M6.5 5.6v4.8M9.5 5.6v4.8" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" />
        </>
      )}
      {state === "error" && (
        <>
          <circle cx="8" cy="8" r="6.25" fill="none" stroke="currentColor" strokeWidth="1.5" />
          <path d="m5.9 5.9 4.2 4.2m0-4.2-4.2 4.2" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" />
        </>
      )}
    </svg>
  )
}

function Note({ text, past }: { text: string; past?: boolean }) {
  const body = useRef<HTMLDivElement>(null)
  const [long, setLong] = useState(false)
  const [open, setOpen] = useState(false)
  useLayoutEffect(() => {
    const el = body.current
    if (!el || open) return
    const measure = () => setLong(el.scrollHeight > el.clientHeight + 1)
    measure()
    const observer = new ResizeObserver(measure)
    observer.observe(el)
    return () => observer.disconnect()
  }, [text, open])
  return (
    <li className={`narration-note${past ? " is-past" : ""}`}>
      {past && <div className="narration-note-past">Earlier, while the run was paused. It has since resumed.</div>}
      <div ref={body} className={`narration-note-body${open ? "" : " is-clamped"}${long && !open ? " is-long" : ""}`}>
        <Markdown>{text}</Markdown>
      </div>
      {(long || open) && (
        <button type="button" className="narration-more" aria-expanded={open} onClick={() => setOpen(!open)}>
          {open ? "Less" : "More"}
        </button>
      )}
    </li>
  )
}

/** A stage's activity: the roll-up line (or "Show N steps") opens the full list;
 * while the stage runs its newest rows stay in view beneath. */
function Activity({
  id,
  lines,
  shown,
  rollUp = [],
}: {
  id: string
  lines: ActivityLine[]
  shown: StepRow[]
  rollUp?: RollUpPart[]
}) {
  const [open, setOpen] = useState(false)
  const listId = `narration-steps-${id}`
  const count = `${lines.length} step${lines.length === 1 ? "" : "s"}`
  return (
    <>
      {lines.length > shown.length && (
        <button
          type="button"
          className={`narration-toggle${rollUp.length ? " is-rollup" : ""}`}
          aria-expanded={open}
          aria-controls={listId}
          onClick={() => setOpen(!open)}
        >
          {rollUp.length ? <RollUp parts={rollUp} /> : open ? `Hide ${count}` : `Show ${count}`}
        </button>
      )}
      <div id={listId} className={`narration-collapse${open ? " is-open" : ""}`} inert={!open} aria-hidden={!open}>
        <ol>
          {lines.map((line) => (
            <Line key={line.id} row={{ ...line, count: 1, flagged: line.flag ? 1 : 0 }} />
          ))}
        </ol>
      </div>
      {!open && shown.length > 0 && (
        <ol>
          {shown.map((row) => (
            <Line key={row.id} row={row} />
          ))}
        </ol>
      )}
    </>
  )
}

function RollUp({ parts }: { parts: RollUpPart[] }) {
  return (
    <span className="narration-rollup">
      {parts.map((part, index) => (
        <span key={index}>
          {index > 0 && " · "}
          {part.text}
          {part.text && part.flagged && ", "}
          {part.flagged && <span className="narration-flagged">{part.flagged}</span>}
        </span>
      ))}
    </span>
  )
}

function Line({ row }: { row: StepRow }) {
  const state = row.flag && row.state === "done" ? "flagged" : row.state
  return (
    <li className={`narration-line ${row.kind} ${state}`} title={row.detail}>
      <span className={`narration-dot small ${state}`} aria-hidden />
      <span className="narration-line-text">
        {row.text}
        {row.count > 1 && row.target && <span className="narration-suffix"> — {row.target}</span>}
      </span>
    </li>
  )
}
