import type { LeanDraftingStatus } from "../../api/ingest/lean"
import type { NarratorNote } from "../../api/opencode/narrator"
import { cleanTitle } from "../drafting/model"
import { TOOL_GROUPS, type ActivityLine } from "./activity"
import { checkpointKeys } from "./schedule"

// The narration's reading order: one row per stage run, each holding the
// activity its session produced, with narrator updates between them by time. A
// finished stage folds its activity behind a one-line roll-up; the stage that is
// running keeps its last few rows open so the lawyer sees what is happening now.
// A run drafting several documents gets a heading wherever the document changes.

export const VISIBLE_STEPS = 3

/** One visible row of activity: a single line, or a run of one tool's lines. */
export type StepRow = ActivityLine & { count: number; flagged: number }

/** One part of a stage's roll-up, with an amber tail for what needs a look. */
export type RollUpPart = { text: string; flagged?: string }

export type TimelineBlock =
  | { kind: "note"; id: string; text: string; past?: boolean }
  | { kind: "heading"; id: string; text: string }
  | { kind: "milestone"; line: ActivityLine }
  | { kind: "stage"; line: ActivityLine; lines: ActivityLine[]; shown: StepRow[]; rollUp: RollUpPart[] }
  | { kind: "steps"; id: string; lines: ActivityLine[]; shown: StepRow[] }

type Planned = Array<{ id: string; title: string }>

export function narrationTimeline(
  lines: ActivityLine[],
  notes: NarratorNote[],
  live: boolean,
  planned: Planned = [],
  status?: LeanDraftingStatus,
): TimelineBlock[] {
  const current = status ? new Set(checkpointKeys(status)) : undefined
  const stages = lines.filter((line) => line.kind === "milestone" && line.step).sort((a, b) => a.at - b.at)
  const attached = new Map(stages.map((stage) => [stage.id, [] as ActivityLine[]]))
  const loose = lines
    .filter((line) => line.kind !== "milestone")
    .sort((a, b) => a.at - b.at)
    .filter((line) => {
      const stage = owningStage(line, stages)
      if (!stage) return true
      attached.get(stage.id)!.push(line)
      return false
    })
  const entries = [
    ...stages.map((line) => ({ at: line.at, block: stageBlock(line, attached.get(line.id)!, live) })),
    ...lines
      .filter((line) => line.kind === "milestone" && !line.step)
      .map((line) => ({ at: line.at, block: { kind: "milestone" as const, line } })),
    ...uniqueNotes(notes).map((note) => ({
      at: note.at,
      block: { kind: "note" as const, id: note.id, text: note.text, ...(current && supersededPause(note, current) && { past: true }) },
    })),
    ...loose.map((line) => ({ at: line.at, line })),
  ].sort((a, b) => a.at - b.at)
  const blocks = entries.reduce<TimelineBlock[]>((acc, entry) => {
    if ("block" in entry) return [...acc, entry.block]
    const last = acc.at(-1)
    if (last?.kind === "steps") return [...acc.slice(0, -1), { ...last, lines: [...last.lines, entry.line] }]
    return [...acc, { kind: "steps", id: entry.line.id, lines: [entry.line], shown: [] }]
  }, [])
  const tail = blocks.at(-1)
  const shown =
    live && tail?.kind === "steps"
      ? [...blocks.slice(0, -1), { ...tail, shown: stepRows(tail.lines).slice(-VISIBLE_STEPS) }]
      : blocks
  return planned.length < 2 ? shown : withHeadings(shown, planned)
}

/** The stage run a line belongs to: its own step's latest run that had started
 * by then (or that step's first run, should the clocks disagree), else the latest
 * stage that had started. */
function owningStage(line: ActivityLine, stages: ActivityLine[]) {
  const own = stages.filter((stage) => stage.step === line.step)
  const started = (rows: ActivityLine[]) => rows.filter((stage) => stage.at <= line.at).at(-1)
  return own.length ? (started(own) ?? own[0]) : started(stages)
}

function stageBlock(line: ActivityLine, lines: ActivityLine[], live: boolean): TimelineBlock {
  const running = live && line.state === "running"
  return {
    kind: "stage",
    line,
    lines,
    shown: running ? stepRows(lines).slice(-VISIBLE_STEPS) : [],
    rollUp: running ? [] : rollUp(lines),
  }
}

/** A stage's activity in one line: "6 reads of source material · 12 searches of matter documents · Checked 23
 * quotes, 2 not found word for word". */
export function rollUp(lines: ActivityLine[]): RollUpPart[] {
  const of = (group: string) => lines.filter((line) => line.group === group)
  const reads = of("read").length
  const searches = of("search").length
  const webSearches = of("web-search").length
  const webFetches = of("web-fetch").length
  const sourceFetches = of("source-fetch").length
  const retrievalIssues = lines.filter(line => ["web-search", "web-fetch", "source-fetch"].includes(line.group ?? "") && line.flag).length
  const quotes = of("quote")
  const missed = quotes.filter((line) => line.flag).length
  const failed = lines.filter((line) => line.state === "error").length
  const other = lines.filter((line) => line.kind === "tool" && !line.group && line.state !== "error").length
  return [
    reads ? { text: `${count(reads, "read")} of source material` } : undefined,
    searches ? { text: `${count(searches, "search", "searches")} of matter documents` } : undefined,
    webSearches ? { text: count(webSearches, "web search", "web searches") } : undefined,
    webFetches ? { text: count(webFetches, "web page retrieval call") } : undefined,
    sourceFetches ? { text: count(sourceFetches, "source passage retrieval call") } : undefined,
    retrievalIssues ? { text: "", flagged: `${count(retrievalIssues, "retrieval limitation")}` } : undefined,
    quotes.length
      ? {
          text: `Checked ${count(quotes.length, "quote")}`,
          flagged: missed ? `${missed} not found word for word` : undefined,
        }
      : undefined,
    other ? { text: count(other, "other step") } : undefined,
    failed ? { text: "", flagged: `${count(failed, "step")} did not complete` } : undefined,
  ].filter((part) => part !== undefined)
}

function count(n: number, one: string, many = `${one}s`) {
  return `${n} ${n === 1 ? one : many}`
}

const PAUSE_KEY = /^question:|^run:(?!completed:)|:(blocked|failed):/

/** An update written because the run stopped (a question for the lawyer, a
 * block, a pause) that no longer describes the run: none of the stop
 * checkpoints it covered still hold, so the run has since resumed. It stays in
 * the history, marked as past, instead of reading as the run's present state. */
export function supersededPause(note: NarratorNote, current: ReadonlySet<string>) {
  const stops = note.keys.filter((key) => PAUSE_KEY.test(key))
  return stops.length > 0 && !stops.some((key) => current.has(key))
}

/** A narrator update the lawyer has already read, word for word, is not shown twice. */
function uniqueNotes(notes: NarratorNote[]) {
  const seen = new Set<string>()
  return notes.filter((note) => {
    const key = note.text.toLowerCase().replace(/\s+/g, " ").trim()
    if (seen.has(key)) return false
    seen.add(key)
    return true
  })
}

/** A heading before each stage row whose document differs from the stage
 * before it; the run-level steps read as "Finishing". */
function withHeadings(blocks: TimelineBlock[], planned: Planned) {
  return blocks.reduce<{ blocks: TimelineBlock[]; document?: string }>(
    (acc, block) => {
      if (block.kind !== "stage") return { ...acc, blocks: [...acc.blocks, block] }
      const step = block.line.step!
      const document = step.includes(".") ? step.slice(0, step.indexOf(".")) : "run"
      if (document === acc.document) return { ...acc, blocks: [...acc.blocks, block] }
      const heading: TimelineBlock = {
        kind: "heading",
        id: `heading:${block.line.id}`,
        text: headingText(document, planned),
      }
      return { blocks: [...acc.blocks, heading, block], document }
    },
    { blocks: [] },
  ).blocks
}

function headingText(document: string, planned: Planned) {
  if (document === "run") return "Finishing"
  const index = planned.findIndex((row) => row.id === document)
  const title = index === -1 ? "" : cleanTitle(planned[index]!.title)
  if (!title) return `Document ${index === -1 ? planned.length + 1 : index + 1}`
  return `${title[0]!.toUpperCase()}${title.slice(1)}`
}

/** Consecutive lines of one tool group fold into a single row that counts them
 * and carries the latest line's state and target. */
export function stepRows(lines: ActivityLine[]): StepRow[] {
  return lines.reduce<StepRow[]>((rows, line) => {
    const last = rows.at(-1)
    const flagged = line.flag ? 1 : 0
    if (!last || !line.group || last.group !== line.group) return [...rows, { ...line, count: 1, flagged }]
    const count = last.count + 1
    const total = last.flagged + flagged
    return [
      ...rows.slice(0, -1),
      {
        ...line,
        id: last.id,
        at: last.at,
        count,
        flagged: total,
        flag: total > 0,
        text: groupText(line.group, line.state, count, total),
        detail: undefined,
      },
    ]
  }, [])
}

function groupText(group: string, state: ActivityLine["state"], count: number, flagged: number) {
  const words = TOOL_GROUPS[group]
  if (!words) return `${group} · ${count}`
  const head = `${state === "running" ? words.running : words.done} · ${count}`
  return flagged && words.flagged ? `${head}, ${flagged} ${words.flagged}` : head
}

/** Elapsed time as "45s", "4m 12s" or "1h 3m". */
export function elapsed(ms: number) {
  const seconds = Math.max(0, Math.floor(ms / 1000))
  if (seconds < 60) return `${seconds}s`
  const minutes = Math.floor(seconds / 60)
  if (minutes < 60) return `${minutes}m ${seconds % 60}s`
  return `${Math.floor(minutes / 60)}h ${minutes % 60}m`
}

/** How long a completed run took, from its start to its last finished stage,
 * as "7m", "1h 5m" or "under a minute". */
export function draftedIn(status: LeanDraftingStatus) {
  if (status.status !== "completed") return undefined
  const finished = Math.max(...Object.values(status.stages).map((row) => Date.parse(row.finished_at ?? "") || 0), 0)
  const ms = (finished || Date.parse(status.updated_at)) - Date.parse(status.created_at)
  if (!Number.isFinite(ms)) return undefined
  const minutes = Math.round(ms / 60_000)
  if (minutes < 1) return "under a minute"
  if (minutes < 60) return `${minutes}m`
  return `${Math.floor(minutes / 60)}h ${minutes % 60}m`
}
