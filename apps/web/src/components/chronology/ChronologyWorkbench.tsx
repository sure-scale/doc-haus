import { useEffect, useMemo, useRef, useState } from "react"
import {
  batchChronologyAdditions,
  chronologyAction,
  chronologyCounters,
  ChronologyConflict,
  chronologyNeedsAttention,
  chronologyPreflight,
  decideChronologyProposal,
  getChronology,
  previewChronologyAction,
  runChronology,
  stopChronology,
  type Chronology,
  type ChronologyAction,
  type ChronologyEvent,
  type ChronologyEvidence,
  type ChronologyPreflight,
  type ChronologyPrivilege,
} from "../../api/ingest/chronology"
import { type Document } from "../../api/ingest/documents"
import DocumentViewer from "../DocumentViewer"
import { type CitationFocus } from "../citation-target"
import ChronologyDialog from "./Dialog"
import EventDetail from "./EventDetail"
import EventForm from "./EventForm"
import ProposalReview from "./ProposalReview"
import ChronologyTimeline from "./Timeline"
import SourceIcon from "./SourceIcon"
import {
  chronologyChatContext,
  chronologyPeople,
  dateLabel,
  emptyFilters,
  eventEvidence,
  eventQualifications,
  matchesChronology,
  type ChronologyFilters,
} from "./model"
import "./chronology.css"
export default function ChronologyWorkbench({
  matterId,
  title,
  client,
  documents,
  chronology,
  initialEvent,
  initialProposal,
  onState,
  onAsk,
  onClose,
}: {
  matterId: string
  title: string
  client?: string
  documents: Document[]
  chronology?: Chronology
  initialEvent?: string
  initialProposal?: string
  onState: (state: Chronology) => void
  onAsk: (context: string) => void
  onClose: () => void
}) {
  const [state, setState] = useState(chronology)
  const [selectedId, setSelectedId] = useState<string | undefined>(initialEvent)
  const [search, setSearch] = useState("")
  const [sort, setSort] = useState("newest")
  const [filters, setFilters] = useState<ChronologyFilters>(emptyFilters)
  const [filterOpen, setFilterOpen] = useState(false)
  const [error, setError] = useState("")
  const [busy, setBusy] = useState(false)
  const busyRef = useRef(false)
  const [sourcesOpen, setSourcesOpen] = useState(false)
  const [updatesOpen, setUpdatesOpen] = useState(!!initialProposal)
  const [editing, setEditing] = useState<{ mode: "add" | "edit" | "merge" | "split"; event?: ChronologyEvent }>()
  const [sourceDate, setSourceDate] = useState(false)
  const [generation, setGeneration] = useState<"initial" | "refresh" | "resume">()
  const [preflight, setPreflight] = useState<ChronologyPreflight>()
  const [loadingPreflight, setLoadingPreflight] = useState(false)
  const [privilege, setPrivilege] = useState<ChronologyPrivilege>()
  const [viewing, setViewing] = useState<{ name: string; citation?: CitationFocus }>()
  const root = useRef<HTMLElement>(null)
  const searchControl = useRef<HTMLInputElement>(null)
  const rowFocus = useRef<HTMLButtonElement | null>(null)
  const tableScroll = useRef<Array<{ element: HTMLElement; top: number; left: number }>>([])
  const selected = state?.events.find((event) => event.id === selectedId)
  function receive(next: Chronology) {
    setState((current) => (current && current.revision > next.revision ? current : next))
    onState(next)
  }
  useEffect(() => {
    if (chronology) setState((current) => (current && current.revision > chronology.revision ? current : chronology))
  }, [chronology])
  useEffect(() => {
    let live = true
    if (!chronology)
      getChronology(matterId).then(
        (next) => {
          if (live) receive(next)
        },
        (failure) => {
          if (live) setError(failure instanceof Error ? failure.message : "Could not load the chronology.")
        },
      )
    return () => {
      live = false
    }
  }, [matterId])
  useEffect(() => {
    if (initialEvent) setSelectedId(initialEvent)
  }, [initialEvent])
  useEffect(() => {
    if (initialProposal) setUpdatesOpen(true)
  }, [initialProposal])
  useEffect(() => {
    if (state?.run?.status !== "running") return
    let live = true
    const timer = window.setInterval(() => {
      if (!busyRef.current)
        getChronology(matterId).then(
          (next) => {
            if (live) receive(next)
          },
          () => {},
        )
    }, 1800)
    return () => {
      live = false
      window.clearInterval(timer)
    }
  }, [matterId, state?.run?.status])
  async function perform(operation: () => Promise<Chronology>): Promise<boolean> {
    if (busyRef.current) return false
    busyRef.current = true
    setBusy(true)
    setError("")
    try {
      receive(await operation())
      return true
    } catch (failure) {
      if (failure instanceof ChronologyConflict && failure.current) receive(failure.current)
      else await getChronology(matterId).then(receive, () => {})
      setError(
        failure instanceof Error ? failure.message : "Could not save. Reload the latest state before another decision.",
      )
      return false
    } finally {
      busyRef.current = false
      setBusy(false)
    }
  }
  const action = (value: ChronologyAction) =>
    state ? perform(() => chronologyAction(matterId, state.revision, value)) : Promise.resolve(false)
  async function preview(value: ChronologyAction) {
    if (!state || busyRef.current) return
    busyRef.current = true
    setBusy(true)
    setError("")
    try {
      return (await previewChronologyAction(matterId, state.revision, value)).preview
    } catch (failure) {
      setError(failure instanceof Error ? failure.message : "Could not preview these events.")
      return undefined
    } finally {
      busyRef.current = false
      setBusy(false)
    }
  }
  function select(event: ChronologyEvent, control?: HTMLButtonElement) {
    if (!selectedId) {
      const containers: HTMLElement[] = []
      for (let element = root.current; element; element = element.parentElement) containers.push(element)
      const table = root.current?.querySelector<HTMLElement>(".chronology-table-scroll")
      if (table) containers.push(table)
      tableScroll.current = containers.map((element) => ({ element, top: element.scrollTop, left: element.scrollLeft }))
    }
    rowFocus.current =
      control ??
      [...(root.current?.querySelectorAll<HTMLButtonElement>(".chronology-event-title") ?? [])].find(
        (button) => button.closest<HTMLElement>("[data-event-id]")?.dataset.eventId === event.id,
      ) ??
      null
    setSelectedId(event.id)
    if (!event.opened) void action({ type: "open", eventId: event.id })
  }
  function closeDetail() {
    setSelectedId(undefined)
    requestAnimationFrame(() => {
      for (const { element, top, left } of tableScroll.current) element.scrollTo({ top, left, behavior: "instant" })
      const target = rowFocus.current
      if (target?.isConnected && target.getClientRects().length) target.focus({ preventScroll: true })
      else {
        searchControl.current?.focus({ preventScroll: true })
        if (!tableScroll.current.length)
          searchControl.current?.scrollIntoView({ block: "nearest", behavior: "instant" })
      }
      tableScroll.current = []
    })
  }
  function inspect(evidence: ChronologyEvidence) {
    const safeName = !/[\\/\u0000]/.test(evidence.documentName)
    const matches = documents.filter((document) => document.name === evidence.documentName)
    const supplied = state?.coverage.some(
      (source) =>
        source.suppliedChronology &&
        source.sourceId === evidence.sourceId &&
        source.documentName === evidence.documentName &&
        source.status !== "removed" &&
        source.uploaded,
    )
    if (!safeName || (matches.length !== 1 && !supplied)) {
      setError("This supporting document is no longer available in this matter. The event and review are preserved.")
      return
    }
    setViewing({
      name: evidence.documentName,
      citation:
        evidence.verified && evidence.freshness === "current"
          ? {
              section: evidence.section,
              excerpt: evidence.excerpt,
              charStart: evidence.charStart,
              charEnd: evidence.charEnd,
            }
          : undefined,
    })
  }
  async function prepare(mode: "initial" | "refresh" | "resume") {
    setGeneration(mode)
    setPreflight(undefined)
    setPrivilege(undefined)
    setLoadingPreflight(true)
    setError("")
    try {
      setPreflight(await chronologyPreflight(matterId))
    } catch (failure) {
      setError(failure instanceof Error ? failure.message : "Could not inspect the sources.")
    } finally {
      setLoadingPreflight(false)
    }
  }
  const counters = state && chronologyCounters(state)
  const people = useMemo(() => chronologyPeople(state?.events ?? []), [state?.events])
  const events = (state?.events ?? [])
    .filter((event) => !event.supersededBy?.length)
    .filter((event) => matchesChronology(event, search, filters, people.canonical))
    .sort((a, b) => {
      if (!a.date.start) return b.date.start ? 1 : a.createdAt - b.createdAt
      if (!b.date.start) return -1
      const order = a.date.start.localeCompare(b.date.start) || a.createdAt - b.createdAt
      return sort === "newest" ? -order : order
    })
  const activeFilterCount = Object.entries(filters).filter(([key, value]) =>
    key === "hidden" ? value !== "exclude" : !!value,
  ).length
  const running = state?.run?.status === "running"
  const complete = state?.coverage.filter((source) => source.status === "complete").length ?? 0
  return (
    <section
      ref={root}
      className={`chronology-workbench${selected ? " has-selection" : ""}${viewing ? " has-source-viewer" : ""}`}
      aria-label={`Fact chronology — ${title}`}
    >
      <div className="chronology-head">
        <div>
          <h1>Fact chronology</h1>
          {!client && <p className="chronology-muted">Client framing not recorded; a neutral account is used.</p>}
        </div>
        <div className="chronology-inline">
          <button
            type="button"
            className="primary"
            disabled={!state || busy}
            onClick={() => setEditing({ mode: "add" })}
          >
            + Add event
          </button>
          <details className="chronology-workbench-menu">
            <summary aria-label="Chronology actions" title="Chronology actions">
              ···
            </summary>
            <div>
              {state?.created && (
                <button
                  type="button"
                  disabled={busy || running}
                  onClick={(click) => {
                    click.currentTarget.closest("details")?.removeAttribute("open")
                    void prepare("refresh")
                  }}
                >
                  Refresh from sources
                </button>
              )}
              <button type="button" onClick={onClose}>
                Return to conversation
              </button>
            </div>
          </details>
        </div>
      </div>
      {error && (
        <div className="chronology-error" role="alert">
          <span>{error}</span>
          <button
            type="button"
            onClick={() => void getChronology(matterId).then(receive, (failure) => setError(failure.message))}
          >
            Reload latest
          </button>
        </div>
      )}
      {!state && <p role="status">Loading saved chronology…</p>}
      {state && (
        <>
          <div className="chronology-counters" aria-label="Chronology status">
            <span>{state.events.filter((event) => !event.hidden && !event.supersededBy?.length).length} events</span>
            <button
              type="button"
              title="Saved events not yet opened; opening does not mark reviewed"
              onClick={() => {
                setFilters({ ...emptyFilters, review: "new" })
                setUpdatesOpen(false)
              }}
            >
              {counters!.new} New
            </button>
            <button
              type="button"
              title="Events awaiting individual lawyer review"
              onClick={() => {
                setFilters({ ...emptyFilters, review: "unreviewed" })
                setUpdatesOpen(false)
              }}
            >
              {counters!.unreviewed} Unreviewed
            </button>
            <button
              type="button"
              title="Distinct events with unresolved conflicts or stale/unverified evidence"
              onClick={() => {
                setFilters({ ...emptyFilters, review: "attention" })
                setUpdatesOpen(false)
              }}
            >
              {counters!.attention} Needs attention
            </button>
            <button
              type="button"
              aria-expanded={updatesOpen}
              onClick={() => {
                setUpdatesOpen((value) => !value)
                setSelectedId(undefined)
              }}
            >
              {counters!.updates} Updates
            </button>
            <button
              type="button"
              aria-expanded={sourcesOpen}
              aria-controls="chronology-sources"
              onClick={() => setSourcesOpen((value) => !value)}
            >
              Sources · {complete}/{state.coverage.length} processed
              {counters!.sources ? ` · ${counters!.sources} incomplete` : ""}
            </button>
          </div>
          {!state.created && !state.events.length && (
            <div className="chronology-invitation">
              <h2>Build a source-linked account of what happened</h2>
              <p>
                {state.suggestion.reason ||
                  "Bring the events across this matter into one timeline, with source passages, uncertainty and a review queue."}
              </p>
              <p>
                Generation covers the authorised whole matter. Usable results are saved directly as an unreviewed draft,
                with incomplete coverage shown. You review individual events after generation.
              </p>
              <div className="chronology-inline">
                <button
                  type="button"
                  className="primary"
                  disabled={busy || running}
                  onClick={() => void prepare("initial")}
                >
                  Generate chronology
                </button>
                <button
                  type="button"
                  disabled={busy}
                  onClick={async () => {
                    if (await action({ type: "invitation", choice: "hide" })) onClose()
                  }}
                >
                  Hide invitation
                </button>
                <button
                  type="button"
                  disabled={busy}
                  onClick={async () => {
                    if (await action({ type: "invitation", choice: "later" })) onClose()
                  }}
                >
                  Remind me after relevant changes
                </button>
              </div>
            </div>
          )}
          {state.run && (sourcesOpen || state.run.status !== "complete" || state.run.errors.length > 0) && (
            <div className={`chronology-run is-${state.run.status}`} role="status">
              <div>
                <strong>
                  {running
                    ? "Building chronology"
                    : state.run.status === "complete"
                      ? "Source processing finished"
                      : state.run.status === "stopped"
                        ? "Generation stopped"
                        : "Generation interrupted"}
                </strong>
                <span>
                  {" "}
                  · {state.coverage.reduce((sum, source) => sum + source.processedSections.length, 0)}/
                  {state.coverage.reduce((sum, source) => sum + source.sections, 0)} source sections completed ·{" "}
                  {state.run.sourceKeys.length} sources enumerated
                </span>
                <p>
                  {state.run.mode === "initial"
                    ? "Completed draft entries are saved as unreviewed."
                    : "Completed changes are staged for your review. Saved entries are preserved."}{" "}
                  {state.run.model && `Model: ${state.run.model}.`}
                </p>
                {state.run.usage && (
                  <p>
                    Recorded usage: {state.run.usage.input.toLocaleString()} input /{" "}
                    {state.run.usage.output.toLocaleString()} output tokens
                    {state.run.usage.cost !== undefined ? ` · ${state.run.usage.cost.toFixed(4)} recorded cost` : ""}.
                  </p>
                )}
                {state.run.errors.map((issue, index) => (
                  <p key={index} className="chronology-warning">
                    {issue}
                  </p>
                ))}
              </div>
              <div className="chronology-inline">
                {running ? (
                  <button
                    type="button"
                    disabled={busy}
                    onClick={() => void perform(() => stopChronology(matterId, state.revision))}
                  >
                    Stop; keep completed work
                  </button>
                ) : (
                  (state.run.status !== "complete" || counters!.sources > 0) && (
                    <button type="button" disabled={busy} onClick={() => void prepare("resume")}>
                      Resume / retry unfinished sources
                    </button>
                  )
                )}
              </div>
            </div>
          )}
          {sourcesOpen && (
            <section id="chronology-sources" className="chronology-sources">
              <div className="chronology-section-head">
                <h2>Sources & coverage</h2>
                <button type="button" onClick={() => setSourceDate(true)}>
                  Supply / correct source date
                </button>
              </div>
              <p>
                Uploaded, text extracted and chronology processed are separate states. Every eligible source is
                included; the table's display filters do not restrict extraction. Empty results do not prove complete
                processing.
              </p>
              {state.coverage.length === 0 ? (
                <p>No chronology run has enumerated the source set yet. Generate to review the current source list.</p>
              ) : (
                <div className="chronology-source-scroll">
                  <table>
                    <thead>
                      <tr>
                        <th>Source</th>
                        <th>Uploaded / text</th>
                        <th>Chronology coverage</th>
                        <th>Contribution & limitations</th>
                      </tr>
                    </thead>
                    <tbody>
                      {state.coverage.map((source) => (
                        <tr key={source.sourceId}>
                          <td>
                            <strong>{source.documentName}</strong>
                            {source.suppliedChronology && <p>Supplied chronology; provenance may be unknown</p>}
                          </td>
                          <td>
                            {source.uploaded ? "Uploaded" : "Unavailable"} ·{" "}
                            {source.textExtracted ? "Text extracted" : "No readable text"}
                          </td>
                          <td>
                            {source.status}
                            <p>
                              {source.processedSections.length}/{source.sections} sections read
                            </p>
                            {source.processedSections.length > 0 && (
                              <details>
                                <summary>Sections actually read</summary>
                                {source.processedSections.map((section) => section + 1).join(", ")}
                              </details>
                            )}
                          </td>
                          <td>
                            {source.events} sourced events
                            {source.issues.map((issue, index) => (
                              <p key={index} className="chronology-warning">
                                {issue}
                              </p>
                            ))}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
              <p className="chronology-muted">
                Archives and unreadable material may contain sources that could not be processed. A quiet period
                describes the supplied record; it does not prove nothing occurred or that evidence is missing.
              </p>
              {state.gaps.length > 0 && (
                <div className="chronology-notice">
                  <h3>Evidence questions to check</h3>
                  <p className="chronology-muted">
                    These source-level questions need checking against the whole record. A reference absent from one
                    section may be supplied elsewhere in the matter.
                  </p>
                  {state.gaps.map((gap, index) => (
                    <p key={index}>{gap}</p>
                  ))}
                </div>
              )}
            </section>
          )}
          {updatesOpen ? (
            <ProposalReview
              chronology={state}
              busy={busy}
              initialProposal={initialProposal}
              onDecision={(proposal, decision) =>
                perform(() => decideChronologyProposal(matterId, state.revision, proposal.id, decision))
              }
              onBatch={(ids) => perform(() => batchChronologyAdditions(matterId, state.revision, ids))}
              onDiscard={() => action({ type: "discard-refresh" })}
              onClose={() => setUpdatesOpen(false)}
              onEvidence={inspect}
              onAsk={(event, proposal) => onAsk(chronologyChatContext(matterId, event, proposal))}
            />
          ) : (
            <div className="chronology-content">
              <div className="chronology-main">
                {(state.created || state.events.length > 0) && (
                  <ChronologyTimeline events={events} selected={selectedId} onSelect={select} />
                )}
                <div className="chronology-search-row">
                  <label className="chronology-search">
                    <span className="chronology-sr-only">Search events and quotations</span>
                    <input
                      ref={searchControl}
                      type="search"
                      placeholder="Search events and supporting quotations…"
                      value={search}
                      onChange={(change) => setSearch(change.target.value)}
                    />
                  </label>
                  <label className="chronology-sort">
                    <span className="chronology-sr-only">Sort events</span>
                    <select value={sort} onChange={(change) => setSort(change.target.value)}>
                      <option value="oldest">Date (oldest first)</option>
                      <option value="newest">Date (newest first)</option>
                    </select>
                  </label>
                  <button
                    type="button"
                    aria-expanded={filterOpen}
                    aria-controls="chronology-filters"
                    onClick={() => setFilterOpen((value) => !value)}
                  >
                    Filter{activeFilterCount ? ` (${activeFilterCount})` : ""}
                  </button>
                  {(search || activeFilterCount > 0) && (
                    <button
                      type="button"
                      onClick={() => {
                        setSearch("")
                        setFilters(emptyFilters)
                      }}
                    >
                      Clear
                    </button>
                  )}
                </div>
                {filterOpen && (
                  <div id="chronology-filters" className="chronology-filters">
                    <label>
                      Document
                      <select
                        value={filters.document}
                        onChange={(change) => setFilters({ ...filters, document: change.target.value })}
                      >
                        <option value="">All documents</option>
                        {[...new Set(state.events.flatMap(eventEvidence).map((evidence) => evidence.documentName))]
                          .sort()
                          .map((name) => (
                            <option key={name}>{name}</option>
                          ))}
                      </select>
                    </label>
                    <label>
                      Person / organisation
                      <select
                        value={filters.person}
                        onChange={(change) => setFilters({ ...filters, person: change.target.value })}
                      >
                        <option value="">All people</option>
                        {people.options.map((person) => (
                          <option key={person}>{person}</option>
                        ))}
                      </select>
                    </label>
                    <label>
                      From date
                      <input
                        type="date"
                        value={filters.from}
                        onChange={(change) => setFilters({ ...filters, from: change.target.value })}
                      />
                    </label>
                    <label>
                      To date
                      <input
                        type="date"
                        value={filters.to}
                        onChange={(change) => setFilters({ ...filters, to: change.target.value })}
                      />
                    </label>
                    <label>
                      Review
                      <select
                        value={filters.review}
                        onChange={(change) => setFilters({ ...filters, review: change.target.value })}
                      >
                        <option value="">All review states</option>
                        <option value="new">New, not opened</option>
                        <option value="unreviewed">Unreviewed</option>
                        <option value="reviewed">Reviewed</option>
                        <option value="attention">Needs attention</option>
                      </select>
                    </label>
                    <label>
                      Significance
                      <select
                        value={filters.significance}
                        onChange={(change) => setFilters({ ...filters, significance: change.target.value })}
                      >
                        <option value="">All significance</option>
                        <option value="key">Key (AI assessment)</option>
                        <option value="relevant">Relevant</option>
                        <option value="background">Background</option>
                      </select>
                    </label>
                    <label>
                      Visibility
                      <select
                        value={filters.hidden}
                        onChange={(change) =>
                          setFilters({ ...filters, hidden: change.target.value as ChronologyFilters["hidden"] })
                        }
                      >
                        <option value="exclude">Visible events</option>
                        <option value="include">Include hidden</option>
                        <option value="only">Hidden only</option>
                      </select>
                    </label>
                    <label className="chronology-check">
                      <input
                        type="checkbox"
                        checked={filters.key}
                        onChange={(change) => setFilters({ ...filters, key: change.target.checked })}
                      />
                      Lawyer-starred key events
                    </label>
                    <label className="chronology-check">
                      <input
                        type="checkbox"
                        checked={filters.conflict}
                        onChange={(change) => setFilters({ ...filters, conflict: change.target.checked })}
                      />
                      Conflicts / disputes
                    </label>
                    <label className="chronology-check">
                      <input
                        type="checkbox"
                        checked={filters.undated}
                        onChange={(change) => setFilters({ ...filters, undated: change.target.checked })}
                      />
                      Undated
                    </label>
                  </div>
                )}
                <div className="chronology-table-scroll" tabIndex={0} role="region" aria-label="Chronology events">
                  <table className="chronology-event-table">
                    <thead>
                      <tr>
                        <th>Date</th>
                        <th>Event</th>
                        <th>Sources</th>
                        <th>Review</th>
                      </tr>
                    </thead>
                    <tbody>
                      {events.map((event) => (
                        <tr
                          key={event.id}
                          data-event-id={event.id}
                          className={`${selectedId === event.id ? "is-selected" : ""}${event.hidden ? " is-hidden" : ""}`}
                        >
                          <td>
                            {dateLabel(event.date)}
                            {event.date.rangeKind === "duration" && (
                              <span className="chronology-table-qualification">Known period</span>
                            )}
                          </td>
                          <td>
                            <button
                              type="button"
                              className="chronology-event-title"
                              aria-expanded={selectedId === event.id}
                              onClick={(click) => select(event, click.currentTarget)}
                            >
                              {event.key && <span aria-label="Key event">★ </span>}
                              {event.title}
                            </button>
                            <div className="chronology-table-qualifications">
                              {eventQualifications(event).map((text) => (
                                <span key={text} className={text === "Conflict / dispute" ? "is-conflict" : undefined}>
                                  {text}
                                </span>
                              ))}
                            </div>
                          </td>
                          <td>
                            {[...new Set(eventEvidence(event).map((evidence) => evidence.documentName))]
                              .slice(0, 2)
                              .map((name) => (
                                <span key={name} className="chronology-source-name">
                                  <SourceIcon
                                    name={name}
                                    kind={
                                      eventEvidence(event).find((evidence) => evidence.documentName === name)
                                        ?.sourceKind
                                    }
                                  />
                                  <span>{name}</span>
                                </span>
                              ))}
                            {new Set(eventEvidence(event).map((evidence) => evidence.documentName)).size > 2 && (
                              <span className="chronology-muted">
                                +{new Set(eventEvidence(event).map((evidence) => evidence.documentName)).size - 2}{" "}
                                sources
                              </span>
                            )}
                            {eventEvidence(event).length === 0 && (
                              <span className="chronology-warning">Source not supplied</span>
                            )}
                          </td>
                          <td>
                            <span className={`chronology-review-state ${event.reviewed ? "is-reviewed" : ""}`}>
                              {event.reviewed ? "✓ Reviewed" : "● Unreviewed"}
                            </span>
                            {!event.opened && <span className="chronology-new">New</span>}
                            {chronologyNeedsAttention(event) && (
                              <span className="chronology-table-qualification is-alert">Needs attention</span>
                            )}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                  {events.length === 0 && (
                    <p className="chronology-empty">
                      {state.events.length > 0
                        ? "No events match these display filters."
                        : "No saved events yet. Generation or a manual addition will appear here."}
                    </p>
                  )}
                </div>
                <p className="chronology-muted chronology-table-foot">
                  Showing {events.length} of {state.events.filter((event) => !event.supersededBy?.length).length} events
                  · Display filters do not limit whole-matter extraction.
                </p>
              </div>
              {selected && (
                <EventDetail
                  key={selected.id}
                  event={selected}
                  chronology={state}
                  busy={busy}
                  onAction={action}
                  onEdit={(mode) => setEditing({ mode, event: selected })}
                  onClose={closeDetail}
                  onEvidence={inspect}
                  onNavigate={(eventId) => {
                    const event = state.events.find((item) => item.id === eventId)
                    if (event) select(event)
                    else setError("This earlier event record is unavailable. Its recorded history is preserved.")
                  }}
                  onAsk={() =>
                    onAsk(
                      chronologyChatContext(
                        matterId,
                        selected,
                        state.proposals.find(
                          (proposal) =>
                            proposal.targetIds.includes(selected.id) &&
                            ["pending", "deferred"].includes(proposal.status),
                        ),
                      ),
                    )
                  }
                />
              )}
            </div>
          )}
        </>
      )}
      {editing && (
        <EventForm
          {...editing}
          error={error}
          events={state?.events.filter((event) => !event.supersededBy?.length) ?? []}
          busy={busy}
          onClose={() => setEditing(undefined)}
          onSave={action}
          onPreview={preview}
        />
      )}
      {sourceDate && (
        <SourceDateDialog
          documents={documents}
          chronology={state!}
          busy={busy}
          error={error}
          onClose={() => setSourceDate(false)}
          onSave={action}
        />
      )}
      {generation && (
        <ChronologyDialog
          title={
            generation === "refresh"
              ? "Refresh chronology"
              : generation === "resume"
                ? "Resume chronology"
                : "Generate chronology"
          }
          onClose={() => setGeneration(undefined)}
        >
          {error && (
            <p role="alert" className="chronology-warning">
              {error}
            </p>
          )}
          {loadingPreflight ? (
            <p role="status">Inspecting the authorised source set…</p>
          ) : preflight ? (
            <>
              <p>{preflight.workExplanation}</p>
              <p>
                <strong>Scope:</strong> every eligible source in this matter. Display filters do not restrict
                processing.
              </p>
              <p>
                <strong>Model:</strong>{" "}
                {preflight.model
                  ? `${preflight.model.providerID} / ${preflight.model.modelID}`
                  : "No model configured — complete AI setup first."}
              </p>
              <p>
                <strong>Expected work:</strong> {preflight.sources.length} sources · approximately{" "}
                {preflight.estimatedSections} bounded source sections. This section count estimates the work, not a
                token or price quote. See native session usage where available; no fixed price is assumed.
              </p>
              <p>
                <strong>Client framing:</strong> {preflight.framing}
                {!preflight.framingKnown ? " — unknown framing; the account stays neutral." : ""}
              </p>
              <details open>
                <summary>Included sources ({preflight.sources.length})</summary>
                <ul className="chronology-included-sources">
                  {preflight.sources.map((source) => (
                    <li key={source.sourceId}>
                      <strong>{source.documentName}</strong> · {source.extraction}
                      {source.kind === "supplied-chronology" ? " · supplied chronology, original preserved" : ""}
                      {source.detail && <p>{source.detail}</p>}
                    </li>
                  ))}
                </ul>
              </details>
              <fieldset>
                <legend>Fresh privilege choice for this run</legend>
                <p>
                  No choice is carried over from an earlier run. A model's flag is not a lawyer's privilege
                  determination.
                </p>
                {(
                  [
                    { value: "all-cleared", label: "All cleared — all included sources are cleared for this use" },
                    { value: "mixed", label: "Mixed — keep privilege uncertainty visible for every event" },
                    { value: "stop", label: "Stop — screen the sources before extraction" },
                  ] as const
                ).map((choice) => (
                  <label className="chronology-check" key={choice.value}>
                    <input
                      type="radio"
                      name="chronology-privilege"
                      value={choice.value}
                      checked={privilege === choice.value}
                      onChange={() => setPrivilege(choice.value)}
                    />
                    {choice.label}
                  </label>
                ))}
              </fieldset>
              <p className="chronology-notice">
                {generation === "initial"
                  ? "Usable first results are saved directly as an unreviewed draft. You can stop and resume without discarding completed work."
                  : "AI changes are staged as proposals. Your saved events and lawyer decisions are preserved until you review the changes."}
              </p>
              <div className="chronology-form-actions">
                <button type="button" onClick={() => setGeneration(undefined)}>
                  Cancel
                </button>
                <button
                  type="button"
                  className="primary"
                  disabled={busy || !privilege || (privilege !== "stop" && !preflight.model)}
                  onClick={async () => {
                    if (
                      privilege &&
                      state &&
                      (await perform(() => runChronology(matterId, state.revision, privilege, generation)))
                    )
                      setGeneration(undefined)
                  }}
                >
                  {privilege === "stop"
                    ? "Stop without extracting"
                    : busy
                      ? "Starting…"
                      : generation === "refresh"
                        ? "Generate refresh proposals"
                        : generation === "resume"
                          ? "Resume unfinished work"
                          : "Generate unreviewed draft"}
                </button>
              </div>
            </>
          ) : (
            <p>Source inspection failed. Close this dialog, reload, and try again.</p>
          )}
        </ChronologyDialog>
      )}
      {viewing && (
        <DocumentViewer
          matterId={matterId}
          name={viewing.name}
          citation={viewing.citation}
          onClose={() => setViewing(undefined)}
        />
      )}
    </section>
  )
}
function SourceDateDialog({
  documents,
  chronology,
  busy,
  error,
  onClose,
  onSave,
}: {
  documents: Document[]
  chronology: Chronology
  busy: boolean
  error: string
  onClose: () => void
  onSave: (action: ChronologyAction) => Promise<boolean>
}) {
  const names = [
    ...new Set([
      ...documents.filter((document) => document.role !== "deliverable").map((document) => document.name),
      ...chronology.coverage
        .filter((source) => source.suppliedChronology && source.status !== "removed" && source.uploaded)
        .map((source) => source.documentName),
    ]),
  ]
  const [documentName, setDocumentName] = useState(names[0] ?? "")
  const [date, setDate] = useState(chronology.sourceDates[documentName]?.date ?? "")
  return (
    <ChronologyDialog title="Supply / correct a source date" onClose={onClose}>
      <form
        onSubmit={async (event) => {
          event.preventDefault()
          if (await onSave({ type: "source-date", documentName, date })) onClose()
        }}
      >
        {error && (
          <p role="alert" className="chronology-warning">
            {error}
          </p>
        )}
        <p>
          This records a lawyer-supplied date without changing the source file. Only affected supported relative dates
          become proposals for your confirmation; the saved event dates are not silently changed.
        </p>
        <label>
          Source
          <select
            value={documentName}
            onChange={(event) => {
              setDocumentName(event.target.value)
              setDate(chronology.sourceDates[event.target.value]?.date ?? "")
            }}
          >
            {names.map((name) => (
              <option key={name}>{name}</option>
            ))}
          </select>
        </label>
        <label>
          Lawyer-supplied source date
          <input type="date" required value={date} onChange={(event) => setDate(event.target.value)} />
        </label>
        {chronology.sourceDates[documentName] && (
          <div>
            <p>
              Previously supplied by {chronology.sourceDates[documentName].actor} ·{" "}
              {chronology.sourceDates[documentName].freshness ?? "version not confirmed"}.
            </p>
            {chronology.sourceDates[documentName].issue && (
              <p className="chronology-warning">{chronology.sourceDates[documentName].issue}</p>
            )}
          </div>
        )}
        <div className="chronology-form-actions">
          <button type="button" onClick={onClose}>
            Cancel
          </button>
          <button type="submit" className="primary" disabled={busy || !documentName || !date}>
            Record date & propose affected changes
          </button>
        </div>
      </form>
    </ChronologyDialog>
  )
}
