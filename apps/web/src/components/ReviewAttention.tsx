import { useEffect, useId, useLayoutEffect, useMemo, useRef, useState } from "react"
import type { GridColumnComparison } from "../api/ingest/grid"
import { legend, type AttentionFilter, type MarkerState, type PrimaryAction, type ReviewAttention, type ReviewProgress } from "../api/review-attention"
import { MarkerGlyph } from "./ReviewMarker"
import { SplitButton } from "./SplitButton"
import { Tooltip } from "./Tooltip"

// The review table's one toolbar row: a quiet progress line on the left; search,
// the Needs attention filter, the quiet actions and the one primary action on
// the right. Filtering and search only change what is shown.

function useNarrow(query = "(max-width: 700px)") {
  const [narrow, setNarrow] = useState(() => typeof window !== "undefined" && Boolean(window.matchMedia?.(query).matches))
  useEffect(() => {
    const media = window.matchMedia?.(query)
    if (!media) return
    const change = () => setNarrow(media.matches)
    change()
    media.addEventListener("change", change)
    return () => media.removeEventListener("change", change)
  }, [query])
  return narrow
}

const plural = (count: number, word: string) => `${count} ${word}${count === 1 ? "" : "s"}`

// "3 documents · 4 questions · 8 of 12 answered · 3 reviewed". When narrow the
// scope moves into the tooltip, and stays in the accessible text.
function ProgressLine({ progress }: { progress: ReviewProgress }) {
  const narrow = useNarrow()
  const scope = `${plural(progress.documents, "document")} · ${plural(progress.questions, "question")}`
  const total = progress.documents * progress.questions
  const done = total ? `${progress.answered} of ${total} answered · ${progress.reviewed} reviewed` : ""
  if (!narrow || !done) return <span className="review-summary">{done ? `${scope} · ${done}` : scope}</span>
  return (
    <Tooltip label={scope} describes={false}>
      <span className="review-summary" tabIndex={0}><span className="sr-only">{scope} · </span>{done}</span>
    </Tooltip>
  )
}

// A disclosure: each state with its count, its shape and meaning, selectable as
// the table's filter. Escape returns focus to the trigger; arrows move between states.
function NeedsAttention({ model, checking, filter, onFilter }: { model: ReviewAttention; checking: boolean; filter?: AttentionFilter; onFilter: (filter?: AttentionFilter) => void }) {
  const [open, setOpen] = useState(false)
  const root = useRef<HTMLDivElement>(null)
  const trigger = useRef<HTMLButtonElement>(null)
  const panel = useRef<HTMLDivElement>(null)
  const id = useId()
  const [shift, setShift] = useState(0)
  const notes = useMemo(legend, [])
  const { attention, progress } = model

  // Right-anchored; nudged back inside the viewport when the trigger wraps to the left edge.
  useLayoutEffect(() => {
    if (!open || !panel.current) return
    const node = panel.current
    node.style.right = "0px"
    const rect = node.getBoundingClientRect()
    const zoom = parseFloat(getComputedStyle(document.body).zoom) || 1
    const wide = document.documentElement.clientWidth
    const move = rect.left < 8 ? 8 - rect.left : rect.right > wide - 8 ? wide - 8 - rect.right : 0
    node.style.right = ""
    setShift(move / zoom)
  }, [open])

  useEffect(() => {
    if (!open) return
    const away = (event: PointerEvent) => { if (!root.current?.contains(event.target as Node)) setOpen(false) }
    document.addEventListener("pointerdown", away)
    return () => document.removeEventListener("pointerdown", away)
  }, [open])

  const choose = (state: AttentionFilter) => {
    onFilter(filter === state ? undefined : state)
    setOpen(false)
    trigger.current?.focus()
  }
  const cells = progress.documents * progress.questions
  const label = attention ? <>Needs attention · <strong>{attention}</strong></>
    : checking ? "Checking sources…"
    : cells && progress.reviewed === cells ? "All reviewed"
    : "Nothing flagged"

  return (
    <div
      ref={root}
      className="review-attention"
      onKeyDown={(event) => {
        if (event.key === "Escape" && open) {
          event.preventDefault()
          event.stopPropagation()
          setOpen(false)
          trigger.current?.focus()
        } else if ((event.key === "ArrowDown" || event.key === "ArrowUp") && open) {
          const buttons = [...panel.current!.querySelectorAll<HTMLButtonElement>("button:not(:disabled)")]
          const at = buttons.indexOf(document.activeElement as HTMLButtonElement)
          if (at < 0 && event.key === "ArrowUp") return
          event.preventDefault()
          buttons[(at + (event.key === "ArrowDown" ? 1 : buttons.length - 1)) % buttons.length]?.focus()
        }
      }}
      onBlur={(event) => {
        const next = event.relatedTarget as Node | null
        if (open && next && !root.current?.contains(next)) setOpen(false)
      }}
    >
      <button
        ref={trigger}
        type="button"
        className={`review-attention-trigger${attention ? " has-attention" : ""}`}
        aria-expanded={open}
        aria-controls={id}
        onClick={() => setOpen((value) => !value)}
      >
        {label}
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><path d="m6 9 6 6 6-6" /></svg>
      </button>
      <div id={id} ref={panel} className="review-attention-menu" role="group" aria-label="Show only" hidden={!open} style={shift ? { right: -shift } : undefined}>
        <p className="review-attention-heading" aria-hidden="true">Show only</p>
        {notes.map((note) => {
          const count = model.counts[note.state]
          const marker = note.state !== "unreviewed"
          return (
            <button
              key={note.state}
              type="button"
              className={marker ? undefined : "is-unreviewed"}
              aria-pressed={filter === note.state}
              aria-describedby={marker ? `${id}-${note.state}` : undefined}
              disabled={!count && filter !== note.state}
              onClick={() => choose(note.state)}
            >
              {marker ? <MarkerGlyph state={note.state as MarkerState} /> : <span className="review-glyph" aria-hidden="true" />}
              <span className="review-attention-word">
                {note.word}
                {note.state === "differs" && <span className="review-ai-tag">AI</span>}
              </span>
              <span className="review-attention-count">{count}</span>
              {marker && <small id={`${id}-${note.state}`} aria-hidden="true">{note.reason}</small>}
            </button>
          )
        })}
      </div>
    </div>
  )
}

// A column header's comparison, in a word, as a disclosure: the AI summary and
// each differing document, which opens that cell. Escape returns focus to the
// trigger; arrows move between the documents.
export function CompareSummary({ comparison, onOpen }: { comparison: GridColumnComparison; onOpen: (docName: string) => void }) {
  const [open, setOpen] = useState(false)
  const root = useRef<HTMLDivElement>(null)
  const trigger = useRef<HTMLButtonElement>(null)
  const panel = useRef<HTMLDivElement>(null)
  const id = useId()
  const differing = Object.entries(comparison.cells).filter(([, cell]) => cell.label === "differs")
  const total = Object.keys(comparison.cells).length
  const label = differing.length ? `${differing.length} of ${total} ${differing.length === 1 ? "differs" : "differ"}` : "Consistent"
  const [place, setPlace] = useState<{ top: number; left: number }>()

  // Fixed, so the table's scroll area cannot clip it: under the trigger, kept
  // inside the viewport. Scrolling or resizing closes it.
  useLayoutEffect(() => {
    if (!open || !panel.current || !trigger.current) return
    const zoom = parseFloat(getComputedStyle(document.body).zoom) || 1
    const anchor = trigger.current.getBoundingClientRect()
    const width = panel.current.getBoundingClientRect().width
    const wide = document.documentElement.clientWidth
    setPlace({ top: (anchor.bottom + 6) / zoom, left: Math.max(16, Math.min(anchor.left - 4, wide - 16 - width)) / zoom })
  }, [open])

  useEffect(() => {
    if (!open) return
    const away = (event: PointerEvent) => { if (!root.current?.contains(event.target as Node)) setOpen(false) }
    const moved = (event: Event) => { if (!panel.current?.contains(event.target as Node)) setOpen(false) }
    const resized = () => setOpen(false)
    document.addEventListener("pointerdown", away)
    window.addEventListener("scroll", moved, true)
    window.addEventListener("resize", resized)
    return () => {
      document.removeEventListener("pointerdown", away)
      window.removeEventListener("scroll", moved, true)
      window.removeEventListener("resize", resized)
    }
  }, [open])

  return (
    <div
      ref={root}
      className="review-compare"
      onKeyDown={(event) => {
        if (event.key === "Escape" && open) {
          event.preventDefault()
          event.stopPropagation()
          setOpen(false)
          trigger.current?.focus()
        } else if ((event.key === "ArrowDown" || event.key === "ArrowUp") && open && differing.length) {
          const buttons = [...panel.current!.querySelectorAll<HTMLButtonElement>("button")]
          const at = buttons.indexOf(document.activeElement as HTMLButtonElement)
          if (at < 0 && event.key === "ArrowUp") return
          event.preventDefault()
          buttons[(at + (event.key === "ArrowDown" ? 1 : buttons.length - 1)) % buttons.length]?.focus()
        }
      }}
      onBlur={(event) => {
        const next = event.relatedTarget as Node | null
        if (open && next && !root.current?.contains(next)) setOpen(false)
      }}
    >
      <button ref={trigger} type="button" className={`review-compare-trigger${differing.length ? " has-differs" : ""}`} aria-label={`${label}, AI comparison`} aria-expanded={open} aria-controls={id} onClick={() => setOpen((value) => !value)}>
        {differing.length > 0 && <MarkerGlyph state="differs" />}
        {label}
        <span className="review-ai-tag">AI</span>
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><path d="m6 9 6 6 6-6" /></svg>
      </button>
      <div id={id} ref={panel} className="review-compare-menu" role="group" aria-label="AI comparison" hidden={!open} style={place}>
        <p className="review-compare-summary">{comparison.summary}<span className="review-ai-tag">AI</span></p>
        {differing.map(([docName, cell], index) => (
          <button
            key={docName}
            type="button"
            aria-describedby={`${id}-${index}`}
            onClick={() => {
              setOpen(false)
              trigger.current?.focus()
              onOpen(docName)
            }}
          >
            <MarkerGlyph state="differs" />
            <span>{docName}</span>
            <small id={`${id}-${index}`} aria-hidden="true">{cell.reason}</small>
          </button>
        ))}
      </div>
    </div>
  )
}

export function ReviewToolbar({ model, checking, running, comparing, primary, fillCount, fillBlanks, startDisabled, fillDisabled, canExport, searchable, filter, onFilter, query, onQuery, onStarter, onFill, onReviewNext, onExport, notice }: {
  model: ReviewAttention
  checking: boolean
  running: number
  // Columns whose answers are being compared after a Fill.
  comparing: number
  primary?: PrimaryAction
  fillCount: number
  // Whether any target is blank; otherwise every target is an answer to refresh.
  fillBlanks: boolean
  startDisabled: boolean
  fillDisabled: boolean
  canExport: boolean
  searchable: boolean
  filter?: AttentionFilter
  onFilter: (filter?: AttentionFilter) => void
  query: string
  onQuery: (query: string) => void
  onStarter: () => void
  onFill: () => void
  onReviewNext: () => void
  onExport: (format: "xlsx" | "csv" | "docx") => void
  // A brief word on a step just finished, such as a Review next walk.
  notice?: string
}) {
  const exportButton = (main: boolean) => (
    <SplitButton
      key={main ? "primary" : "export"}
      className={`review-export${main ? " review-primary" : ""}`}
      primary={main}
      label="Export Excel"
      onClick={() => onExport("xlsx")}
      toggleLabel="More export formats"
      items={[
        { label: "Export CSV", onSelect: () => onExport("csv") },
        { label: "Export Word", onSelect: () => onExport("docx") },
      ]}
    />
  )
  // One slot, so a primary button keeps its node (and focus) as its action changes.
  const main = primary?.kind === "export" ? exportButton(true)
    : primary?.kind === "start" ? <button key="primary" type="button" className="primary review-primary" disabled={startDisabled} onClick={onStarter}>Use a starter</button>
    : primary?.kind === "fill" ? <button key="primary" type="button" className="primary review-primary" disabled={fillDisabled} onClick={onFill}>{fillBlanks ? "Fill" : "Refresh"} {plural(fillCount, "cell")}</button>
    : primary?.kind === "reviewNext" ? <button key="primary" type="button" className="primary review-primary" onClick={onReviewNext}>Review next</button>
    : null
  return (
    <div className="row review-toolbar">
      <div className="review-status">
        <ProgressLine progress={model.progress} />
        {running > 0 && <span className="review-running" role="status">Extracting {running}</span>}
        {comparing > 0 && <span className="review-running" role="status">Comparing {plural(comparing, "question")}</span>}
        {/* Always present, so screen readers announce the text when it appears. */}
        <span className="review-notice" role="status">{notice}</span>
      </div>
      <div className="review-toolbar-actions">
        {searchable && (
          <label className="matter-search">
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" aria-hidden="true"><circle cx="10.5" cy="10.5" r="6.5" /><path d="m16 16 4.5 4.5" /></svg>
            <span className="sr-only">Search the table</span>
            <input
              type="search"
              className="page-filter"
              placeholder="Search"
              value={query}
              onChange={(event) => onQuery(event.target.value)}
              onKeyDown={(event) => {
                // The first Escape clears the search; the next one closes the table.
                if (event.key !== "Escape" || !query) return
                event.preventDefault()
                event.stopPropagation()
                onQuery("")
              }}
            />
          </label>
        )}
        <NeedsAttention model={model} checking={checking} filter={filter} onFilter={onFilter} />
        {primary?.kind !== "start" && <button key="starter" type="button" className="ghost" disabled={startDisabled} onClick={onStarter}>Use a starter</button>}
        {/* The next steps wrap as one group, so the primary never sits alone. */}
        <div className="review-toolbar-steps">
          {primary?.kind !== "reviewNext" && model.queue.length > 0 && <button key="next" type="button" className="ghost" onClick={onReviewNext}>Review next</button>}
          {primary?.kind !== "export" && canExport && exportButton(false)}
          {main}
        </div>
      </div>
    </div>
  )
}

// One removable line while a filter or search narrows the table.
export function FilterSummary({ shown, total, filter, onClear }: { shown: number; total: number; filter?: AttentionFilter; onClear: () => void }) {
  const word = filter && legend().find((note) => note.state === filter)?.word
  return (
    <p className="review-filter-summary">
      <span role="status">Showing {shown} of {plural(total, "cell")}{word ? ` · ${word}` : ""}</span>
      {" · "}
      <button type="button" className="linklike" onClick={onClear}>Clear</button>
    </p>
  )
}
