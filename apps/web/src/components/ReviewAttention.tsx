import { useEffect, useId, useLayoutEffect, useMemo, useRef, useState } from "react"
import type { GridColumnComparison } from "../api/ingest/grid"
import { legend, type AttentionFilter, type MarkerState, type PrimaryAction, type ReviewAttention } from "../api/review-attention"
import { MarkerGlyph } from "./ReviewMarker"
import { SplitButton } from "./SplitButton"
import { Tooltip } from "./Tooltip"

// The review table's one toolbar row: progress as a bar and one short line on
// the left; search, the filter, export and the one primary action on the right.
// Filtering and search only change what is shown. Each count shows once.

const plural = (count: number, word: string) => `${count} ${word}${count === 1 ? "" : "s"}`
const listed = (words: string[]) => words.length < 2 ? words.join("") : `${words.slice(0, -1).join(", ")} or ${words.at(-1)}`

// A toolbar disclosure: one Escape closes it and returns focus to its trigger,
// as do a press outside and focus leaving; arrows move between its controls.
// Right-anchored, nudged back inside the viewport when the trigger wraps left.
function useDisclosure(items: string) {
  const [open, setOpen] = useState(false)
  const root = useRef<HTMLDivElement>(null)
  const trigger = useRef<HTMLButtonElement>(null)
  const panel = useRef<HTMLDivElement>(null)
  const [shift, setShift] = useState(0)

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
    // Captured on the document, so a tooltip shown inside cannot take this Escape first.
    const escape = (event: KeyboardEvent) => {
      if (event.key !== "Escape") return
      event.preventDefault()
      event.stopPropagation()
      setOpen(false)
      trigger.current?.focus()
    }
    document.addEventListener("pointerdown", away)
    document.addEventListener("keydown", escape, true)
    return () => {
      document.removeEventListener("pointerdown", away)
      document.removeEventListener("keydown", escape, true)
    }
  }, [open])

  const rootProps = {
    ref: root,
    onKeyDown: (event: React.KeyboardEvent) => {
      if (!open || (event.key !== "ArrowDown" && event.key !== "ArrowUp")) return
      const controls = [...panel.current!.querySelectorAll<HTMLElement>(items)]
      const at = controls.indexOf(document.activeElement as HTMLElement)
      if (at < 0 && event.key === "ArrowUp") return
      event.preventDefault()
      controls[(at + (event.key === "ArrowDown" ? 1 : controls.length - 1)) % controls.length]?.focus()
    },
    onBlur: (event: React.FocusEvent) => {
      const next = event.relatedTarget as Node | null
      if (open && next && !root.current?.contains(next)) setOpen(false)
    },
  }
  return { open, setOpen, trigger, panel, rootProps, panelStyle: shift ? { right: -shift } : undefined }
}

const chevron = <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><path d="m6 9 6 6 6-6" /></svg>

type Part = "reviewed" | "answered" | "attention" | "empty"
const PARTS: Array<{ part: Part; word: string; spoken: (count: number) => string }> = [
  { part: "reviewed", word: "Reviewed", spoken: () => "reviewed" },
  { part: "answered", word: "Answered, not reviewed", spoken: () => "answered but not reviewed" },
  { part: "attention", word: "Needs attention", spoken: (count) => count === 1 ? "needs attention" : "need attention" },
  { part: "empty", word: "Not filled", spoken: () => "not filled" },
]

// Every current cell in one part: a flagged cell needs attention whether or not
// it was reviewed, since its marker says the answer may no longer hold.
function progressParts(model: ReviewAttention) {
  const count: Record<Part, number> = { reviewed: 0, answered: 0, attention: 0, empty: 0 }
  for (const entry of model.cells.values()) count[entry.marker ? "attention" : !entry.cell ? "empty" : entry.cell.status === "reviewed" ? "reviewed" : "answered"]++
  return PARTS.map((part) => ({ ...part, count: count[part.part] })).filter((part) => part.count > 0)
}

// A thin segmented bar and one short line. The bar's name gives every count in
// words; hover or focus shows its key, with the table's documents and questions.
function ProgressBar({ model }: { model: ReviewAttention }) {
  const { documents, questions } = model.progress
  const scope = `${plural(documents, "document")} · ${plural(questions, "question")}`
  const total = model.cells.size
  if (!total) return <span className="review-summary">{scope}</span>
  const parts = progressParts(model)
  const reviewed = parts.find((part) => part.part === "reviewed")?.count ?? 0
  const name = `Review progress, ${plural(total, "cell")} in ${plural(documents, "document")} and ${plural(questions, "question")}: ${parts.map((part) => `${part.count} ${part.spoken(part.count)}`).join(", ")}`
  const key = (
    <span className="review-progress-key">
      <span className="review-progress-scope">{scope}</span>
      {parts.map((part) => (
        <span key={part.part} className="review-progress-key-row">
          <span className={`review-progress-swatch is-${part.part}`} aria-hidden="true" />
          <span>{part.word}</span>
          <span className="review-progress-key-count">{part.count}</span>
        </span>
      ))}
    </span>
  )
  return (
    <span className="review-progress">
      <Tooltip label={key} describes={false}>
        <span className="review-progress-bar" role="img" aria-label={name} tabIndex={0}>
          {parts.map((part) => <span key={part.part} className={`review-progress-part is-${part.part}`} style={{ flexGrow: part.count }} />)}
        </span>
      </Tooltip>
      <span className="review-summary" aria-hidden="true">{reviewed === total ? "All reviewed" : `${reviewed} of ${total} reviewed`}</span>
    </span>
  )
}

// A magnifier that opens the search field. Escape clears a search, then closes
// the field; leaving an empty field closes it; a search keeps it open.
function SearchToggle({ query, onQuery }: { query: string; onQuery: (query: string) => void }) {
  const [open, setOpen] = useState(false)
  const toggle = useRef<HTMLButtonElement>(null)
  const input = useRef<HTMLInputElement>(null)
  const refocus = useRef<"input" | "toggle" | undefined>(undefined)
  const shown = open || !!query
  useLayoutEffect(() => {
    if (refocus.current === "input") input.current?.focus()
    if (refocus.current === "toggle") toggle.current?.focus()
    refocus.current = undefined
  }, [shown])
  if (!shown) return (
    <Tooltip label="Search the table" describes={false}>
      <button ref={toggle} type="button" className="review-icon-button review-search-toggle" aria-label="Search the table" onClick={() => { refocus.current = "input"; setOpen(true) }}>
        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" aria-hidden="true"><circle cx="10.5" cy="10.5" r="6.5" /><path d="m16 16 4.5 4.5" /></svg>
      </button>
    </Tooltip>
  )
  return (
    <label className="matter-search">
      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" aria-hidden="true"><circle cx="10.5" cy="10.5" r="6.5" /><path d="m16 16 4.5 4.5" /></svg>
      <span className="sr-only">Search the table</span>
      <input
        ref={input}
        type="search"
        className="page-filter"
        placeholder="Search the table"
        value={query}
        onChange={(event) => onQuery(event.target.value)}
        onFocus={() => setOpen(true)}
        onBlur={() => setOpen(false)}
        onKeyDown={(event) => {
          // The first Escape clears the search, the next closes the field; only
          // then does Escape reach the table.
          if (event.key !== "Escape") return
          event.preventDefault()
          event.stopPropagation()
          if (query) onQuery("")
          else {
            refocus.current = "toggle"
            setOpen(false)
          }
        }}
      />
    </label>
  )
}

// The filter: a checkbox per state, with its marker and count, so the list is
// also the legend. Any checked state shows its cells. A state no cell has is
// left out unless it is checked; each state's meaning is in its tooltip.
function FilterMenu({ model, checking, filters, onFilters }: { model: ReviewAttention; checking: boolean; filters: AttentionFilter[]; onFilters: (filters: AttentionFilter[]) => void }) {
  const { open, setOpen, trigger, panel, rootProps, panelStyle } = useDisclosure("input:not(:disabled), button:not(:disabled)")
  const id = useId()
  const notes = useMemo(legend, [])
  const rows = notes.filter((note) => model.counts[note.state] > 0 || filters.includes(note.state))
  const flagged = rows.filter((note) => note.state !== "unreviewed")
  const cells = model.progress.documents * model.progress.questions
  const quiet = checking ? "Checking sources…" : flagged.length ? undefined : cells && model.progress.reviewed === cells ? "All reviewed" : "Nothing flagged"
  const toggle = (state: AttentionFilter, on: boolean) => onFilters(notes.map((note) => note.state).filter((item) => item === state ? on : filters.includes(item)))

  return (
    <div className="review-filter" {...rootProps}>
      <button ref={trigger} type="button" className={`review-filter-trigger${filters.length ? " is-active" : ""}`} aria-expanded={open} aria-controls={id} onClick={() => setOpen((value) => !value)}>
        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><path d="M4 6h16M7 12h10M10 18h4" /></svg>
        Filter
        {filters.length > 0 && <> · <strong>{filters.length}</strong><span className="sr-only"> selected</span></>}
        {chevron}
      </button>
      {open && (
        <div id={id} ref={panel} className="review-filter-menu" role="group" aria-label="Show only" style={panelStyle}>
          <p className="review-filter-heading" aria-hidden="true">Show only</p>
          {quiet && <p className="review-filter-quiet">{quiet}</p>}
          {rows.map((note) => {
            const marker = note.state !== "unreviewed"
            return (
              <Tooltip key={note.state} label={note.reason} describes={false}>
                <label className={`review-filter-option${marker ? "" : " is-unreviewed"}`}>
                  <input type="checkbox" checked={filters.includes(note.state)} aria-label={marker ? `${note.word}, ${plural(model.counts[note.state], "cell")}` : note.word} aria-describedby={`${id}-${note.state}`} onChange={(event) => toggle(note.state, event.target.checked)} />
                  {marker ? <MarkerGlyph state={note.state as MarkerState} /> : <span className="review-glyph" aria-hidden="true" />}
                  <span className="review-filter-word">
                    {note.word}
                    {note.state === "differs" && <span className="review-ai-tag">AI</span>}
                  </span>
                  {/* Not reviewed goes without a count: the progress line already says how many are reviewed. */}
                  {marker && <span className="review-filter-count">{model.counts[note.state]}</span>}
                  <span id={`${id}-${note.state}`} hidden>{note.reason}</span>
                </label>
              </Tooltip>
            )
          })}
          <div className="review-filter-foot">
            <button type="button" className="linklike" disabled={!filters.length} onClick={() => { onFilters([]); trigger.current?.focus() }}>Clear all</button>
          </div>
        </div>
      )}
    </div>
  )
}

// Export as one icon: a menu of formats.
function ExportMenu({ onExport }: { onExport: (format: "xlsx" | "csv" | "docx") => void }) {
  const { open, setOpen, trigger, panel, rootProps, panelStyle } = useDisclosure("button")
  const id = useId()
  const choose = (format: "xlsx" | "csv" | "docx") => {
    setOpen(false)
    trigger.current?.focus()
    onExport(format)
  }
  return (
    <div className="review-export-menu" {...rootProps}>
      <Tooltip label="Export" describes={false}>
        <button ref={trigger} type="button" className="review-icon-button" aria-label="Export" aria-expanded={open} aria-controls={id} onClick={() => setOpen((value) => !value)}>
          <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><path d="M12 4v11M7.5 10.5 12 15l4.5-4.5M5 19.5h14" /></svg>
        </button>
      </Tooltip>
      {open && (
        <div id={id} ref={panel} className="viewer-more-menu review-export-list" role="group" aria-label="Export as" style={panelStyle}>
          <button type="button" onClick={() => choose("xlsx")}>Excel</button>
          <button type="button" onClick={() => choose("csv")}>CSV</button>
          <button type="button" onClick={() => choose("docx")}>Word</button>
        </div>
      )}
    </div>
  )
}

// What a column's AI comparison is and is not, in one line, for its header.
export const COMPARE_HELP = "One AI call compares the cited answers across documents. It does not check legal correctness."

// A column header's comparison, in a few words, as a disclosure: the AI summary
// and each differing document, which opens that cell. Escape returns focus to
// the trigger; arrows move between the documents.
export function CompareSummary({ comparison, onOpen }: { comparison: GridColumnComparison; onOpen: (docName: string) => void }) {
  const [open, setOpen] = useState(false)
  const root = useRef<HTMLDivElement>(null)
  const trigger = useRef<HTMLButtonElement>(null)
  const panel = useRef<HTMLDivElement>(null)
  const id = useId()
  const differing = Object.entries(comparison.cells).filter(([, cell]) => cell.label === "differs")
  const total = Object.keys(comparison.cells).length
  const result = differing.length
    ? `${differing.length} of ${total} ${differing.length === 1 ? "differs" : "differ"}`
    : total === 2 ? "both agree" : `all ${total} agree`
  const [place, setPlace] = useState<{ top: number; left: number }>()

  // Fixed, so the table's scroll area cannot clip it: under the trigger, kept
  // inside the viewport. It follows the trigger while the table scrolls and
  // closes once the trigger leaves the table's view, or on resize.
  const placePanel = () => {
    if (!panel.current || !trigger.current) return
    const zoom = parseFloat(getComputedStyle(document.body).zoom) || 1
    const anchor = trigger.current.getBoundingClientRect()
    const view = trigger.current.closest(".review-scroll")?.getBoundingClientRect()
    if (view && (anchor.right < view.left || anchor.left > view.right || anchor.bottom < view.top || anchor.top > view.bottom)) return setOpen(false)
    const width = panel.current.getBoundingClientRect().width
    const wide = document.documentElement.clientWidth
    setPlace({ top: (anchor.bottom + 6) / zoom, left: Math.max(16, Math.min(anchor.left - 4, wide - 16 - width)) / zoom })
  }
  useLayoutEffect(() => { if (open) placePanel() }, [open])

  useEffect(() => {
    if (!open) return
    const away = (event: PointerEvent) => { if (!root.current?.contains(event.target as Node)) setOpen(false) }
    const moved = (event: Event) => { if (!panel.current?.contains(event.target as Node)) placePanel() }
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
      {/* A native title, not Tooltip: Tooltip takes Escape while shown, and Escape closes this panel. */}
      <button ref={trigger} type="button" className={`review-compare-trigger${differing.length ? " has-differs" : ""}`} title={COMPARE_HELP} aria-expanded={open} aria-controls={id} onClick={() => setOpen((value) => !value)}>
        {/* Each phrase stays whole, the marker with the first and the chevron with the last, so a
            narrow header wraps between them. */}
        <span className="review-compare-text">
          <span>{differing.length > 0 && <MarkerGlyph state="differs" />}AI comparison:</span>{" "}
          <span>{result}<svg className="review-compare-chevron" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><path d="m6 9 6 6 6-6" /></svg></span>
        </span>
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

export function ReviewToolbar({ model, checking, running, comparing, primary, fillCount, fillBlanks, startDisabled, fillDisabled, canExport, searchable, filters, onFilters, query, onQuery, onStarter, onFill, onReviewNext, onExport, notice }: {
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
  filters: AttentionFilter[]
  onFilters: (filters: AttentionFilter[]) => void
  query: string
  onQuery: (query: string) => void
  onStarter: () => void
  onFill: () => void
  onReviewNext: () => void
  onExport: (format: "xlsx" | "csv" | "docx") => void
  // A brief word on a step just finished, such as a Review next walk.
  notice?: string
}) {
  // One slot, so a primary button keeps its node (and focus) as its action changes.
  // With every cell reviewed, Export is the next step in words; before then it is an icon.
  const main = primary?.kind === "export" ? (
    <SplitButton
      key="primary"
      className="review-export review-primary"
      primary
      label="Export Excel"
      onClick={() => onExport("xlsx")}
      toggleLabel="More export formats"
      items={[
        { label: "Export CSV", onSelect: () => onExport("csv") },
        { label: "Export Word", onSelect: () => onExport("docx") },
      ]}
    />
  )
    : primary?.kind === "start" ? <button key="primary" type="button" className="primary review-primary" disabled={startDisabled} onClick={onStarter}>Add questions</button>
    : primary?.kind === "fill" ? <button key="primary" type="button" className="primary review-primary" disabled={fillDisabled} onClick={onFill}>{running > 0 ? "Filling…" : `${fillBlanks ? "Fill" : "Refresh"} ${plural(fillCount, "cell")}`}</button>
    : primary?.kind === "reviewNext" ? <button key="primary" type="button" className="primary review-primary" onClick={onReviewNext}>Review next</button>
    : null
  const busy = running > 0 || comparing > 0 || !!notice
  return (
    <div className={`row review-toolbar${query ? " is-searching" : ""}`}>
      <div className={`review-status${busy ? " is-busy" : ""}`}>
        <ProgressBar model={model} />
        {running > 0 && <span className="review-running" role="status"><span>Extracting {running}</span></span>}
        {comparing > 0 && <span className="review-running" role="status"><span>Comparing {plural(comparing, "question")}</span></span>}
        {/* Always present, so screen readers announce the text when it appears. */}
        <span className="review-notice" role="status">{notice}</span>
      </div>
      <div className="review-toolbar-actions">
        {searchable && <SearchToggle query={query} onQuery={onQuery} />}
        {searchable && <FilterMenu model={model} checking={checking} filters={filters} onFilters={onFilters} />}
        {primary?.kind !== "export" && canExport && <ExportMenu onExport={onExport} />}
        {/* The next steps wrap as one group, so the primary never sits alone. */}
        <div className="review-toolbar-steps">
          {primary?.kind !== "reviewNext" && model.queue.length > 0 && <button key="next" type="button" className="ghost" onClick={onReviewNext}>Review next</button>}
          {main}
        </div>
      </div>
    </div>
  )
}

// One removable line while a filter or search narrows the table: what is
// shown, and by which states and words.
export function FilterSummary({ shown, total, filters, query, onClear }: { shown: number; total: number; filters: AttentionFilter[]; query: string; onClear: () => void }) {
  const words = legend().filter((note) => filters.includes(note.state)).map((note) => note.word)
  const text = query.trim()
  return (
    <p className="review-filter-summary">
      <span role="status">Showing {shown} of {plural(total, "cell")}{words.length ? ` · ${listed(words)}` : ""}{text ? ` · “${text}”` : ""}</span>{" "}
      <button type="button" className="linklike" onClick={onClear}>Clear</button>
    </p>
  )
}
