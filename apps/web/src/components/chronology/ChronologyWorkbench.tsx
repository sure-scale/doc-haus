import { useEffect, useMemo, useRef, useState } from "react"
import useExit from "../useExit"
import backdropClose from "../backdropClose"
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
import JurisdictionSelect from "../JurisdictionSelect"
import ChronologyFiltersPanel from "./ChronologyFiltersPanel"
import SourceCoveragePanel from "./SourceCoveragePanel"
import DateInput, { useChronologyDateFormat } from "./DateInput"
import { type CitationFocus } from "../citation-target"
import ChronologyDialog from "./Dialog"
import ChronologySurface from "./ChronologySurface"
import MatterActionsMenu, { type MatterMenuAction } from "../matter/MatterActionsMenu"
import EventDetail from "./EventDetail"
import EventForm from "./EventForm"
import ProposalReview from "./ProposalReview"
import ChronologyTimeline from "./Timeline"
import SourceIcon from "./SourceIcon"
import {
  chronologyChatReference,
  type ChronologyChatReference,
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
  onActionsChange,
}: {
  matterId: string
  title: string
  client?: string
  documents: Document[]
  chronology?: Chronology
  initialEvent?: string
  initialProposal?: string
  onState: (state: Chronology) => void
  onAsk: (reference: ChronologyChatReference) => void
  onClose: () => void
  onActionsChange?: (actions: MatterMenuAction[]) => void
}) {
  const [state, setState] = useState(chronology)
  const [selectedId, setSelectedId] = useState<string | undefined>(initialEvent)
  const [search, setSearch] = useState("")
  const [sort, setSort] = useState<{
    key: "date" | "event" | "sources" | "review"
    direction: "ascending" | "descending"
  }>({ key: "date", direction: "descending" })
  const [filters, setFilters] = useState<ChronologyFilters>(emptyFilters)
  const [filterOpen, setFilterOpen] = useState(false)
  const [filterReset, setFilterReset] = useState(0)
  const [error, setError] = useState("")
  const [busy, setBusy] = useState(false)
  const busyRef = useRef(false)
  const [sourcesOpen, setSourcesOpen] = useState(false)
  const [updatesOpen, setUpdatesOpen] = useState(!!initialProposal)
  const [editing, setEditing] = useState<{ mode: "add" | "edit" | "merge" | "split"; event?: ChronologyEvent }>()
  const dateFormat = useChronologyDateFormat()
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
  const [detail, detailClosing] = useExit(selected)
  const [compactDetail, setCompactDetail] = useState(() => window.matchMedia("(max-width: 800px)").matches)
  useEffect(() => {
    const query = window.matchMedia("(max-width: 800px)")
    const update = () => setCompactDetail(query.matches)
    query.addEventListener("change", update)
    return () => query.removeEventListener("change", update)
  }, [])

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
    if (initialProposal) showUpdates()
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
      if (compactDetail)
        for (const { element, top, left } of tableScroll.current) element.scrollTo({ top, left, behavior: "instant" })
      const target = rowFocus.current
      if (target?.isConnected && target.getClientRects().length) target.focus({ preventScroll: true })
      else {
        searchControl.current?.focus({ preventScroll: true })
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
      let order = 0
      if (sort.key === "date") {
        if (!a.date.start) return b.date.start ? 1 : a.createdAt - b.createdAt
        if (!b.date.start) return -1
        order = a.date.start.localeCompare(b.date.start)
      } else {
        const value = (event: ChronologyEvent) =>
          sort.key === "event"
            ? event.title
            : sort.key === "sources"
              ? [...new Set(eventEvidence(event).map((evidence) => evidence.documentName))].sort().join(", ")
              : event.reviewed
                ? "Reviewed"
                : "Unreviewed"
        order = value(a).localeCompare(value(b), undefined, { sensitivity: "base", numeric: true })
      }
      return (sort.direction === "descending" ? -order : order) || a.createdAt - b.createdAt
    })
  const activeFilterCount = Object.entries(filters).filter(([key, value]) =>
    key === "hidden" ? value !== "exclude" : !!value,
  ).length
  const running = state?.run?.status === "running"
  const currentEvents = state?.events.filter((event) => !event.hidden && !event.supersededBy?.length) ?? []
  const menuActions: MatterMenuAction[] = [
    ...(running
      ? [
          {
            id: "stop",
            label: "Stop building; keep saved work",
            disabled: busy,
            onSelect: () => {
              if (state) void perform(() => stopChronology(matterId, state.revision))
            },
          },
        ]
      : [
          {
            id: "update",
            label: state?.created ? "Check for source changes" : "Build chronology",
            disabled: !state || busy,
            onSelect: () => void prepare(state?.created ? "refresh" : "initial"),
          },
        ]),
    ...(!running && state?.run && (state.run.status !== "complete" || counters!.sources > 0)
      ? [
          {
            id: "continue",
            label: "Continue unfinished sources",
            disabled: busy,
            onSelect: () => void prepare("resume"),
          },
        ]
      : []),
    { id: "scope", label: "About source processing", disabled: !state, onSelect: showSources },
  ]
  useEffect(() => {
    onActionsChange?.(menuActions)
  }, [onActionsChange, state, busy, matterId])
  function showSources() {
    setSelectedId(undefined)
    setUpdatesOpen(false)
    setSourcesOpen(true)
  }
  function showUpdates() {
    setSelectedId(undefined)
    setSourcesOpen(false)
    setUpdatesOpen(true)
  }
  function resetDisplay() {
    setSearch("")
    setFilters(emptyFilters)
    setFilterReset((value) => value + 1)
  }
  function reviewFilter(review: string) {
    setFilterReset((value) => value + 1)
    setSearch("")
    setFilters({ ...emptyFilters, review: filters.review === review ? "" : review })
  }
  function sortBy(key: typeof sort.key) {
    setSort({ key, direction: sort.key === key && sort.direction === "ascending" ? "descending" : "ascending" })
  }
  const errorAlert = error && (
    <div className="chronology-error" role="alert">
      <span>{error}</span>
      <button
        type="button"
        onClick={() =>
          void getChronology(matterId).then(
            (next) => {
              receive(next)
              setError("")
            },
            (failure) => setError(failure.message),
          )
        }
      >
        Reload latest
      </button>
    </div>
  )
  return (
    <section
      ref={root}
      className={`chronology-workbench${detail ? " has-selection" : ""}${viewing ? " has-source-viewer" : ""}`}
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
          {!onActionsChange && <MatterActionsMenu actions={menuActions} />}
        </div>
      </div>
      {!sourcesOpen && !updatesOpen && errorAlert}
      {!state && <p role="status">Loading saved chronology…</p>}
      {state && (
        <>
          <div className="chronology-counters">
            <div className="chronology-review-filters" role="group" aria-label="Filter saved events">
              <button
                type="button"
                onClick={resetDisplay}
                aria-pressed={!search && activeFilterCount === 0}
                title="Show every visible saved event and clear display filters"
              >
                All events <span>{currentEvents.length}</span>
              </button>
              <span className="chronology-counter-divider" aria-hidden="true" />
              <button
                type="button"
                onClick={() => reviewFilter("new")}
                aria-pressed={filters.review === "new"}
                title="Show events not yet opened. Opening does not mark an event reviewed."
              >
                Not opened <span>{currentEvents.filter((event) => !event.opened).length}</span>
              </button>
              <button
                type="button"
                onClick={() => reviewFilter("unreviewed")}
                aria-pressed={filters.review === "unreviewed"}
                title="Show events awaiting lawyer review"
              >
                Unreviewed <span>{currentEvents.filter((event) => !event.reviewed).length}</span>
              </button>
              <button
                type="button"
                onClick={() => reviewFilter("attention")}
                aria-pressed={filters.review === "attention"}
                title="Show conflicts, uncertain dates or supporting evidence that needs checking"
              >
                Needs attention <span>{currentEvents.filter(chronologyNeedsAttention).length}</span>
              </button>
            </div>
            <div className="chronology-summary-actions">
              <button
                type="button"
                aria-haspopup="dialog"
                onClick={showUpdates}
                title="Review suggested changes before applying them to saved events"
              >
                Updates <span>{counters!.updates}</span>
              </button>
              <button
                type="button"
                aria-haspopup="dialog"
                onClick={showSources}
                title="See which sources and sections were processed, limitations, and outstanding work"
              >
                Sources <span>{state.coverage.length}</span>
                {counters!.sources > 0 && (
                  <span className="chronology-summary-attention">{counters!.sources} to check</span>
                )}
              </button>
            </div>
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
          {state.run && (state.run.status !== "complete" || state.run.errors.length > 0) && (
            <div className={`chronology-run is-${state.run.status}`} role="status">
              {running && <span className="chronology-building-spinner" aria-hidden="true" />}
              <div>
                <strong>
                  {running
                    ? "Building chronology…"
                    : state.run.status === "stopped"
                      ? "Building stopped"
                      : state.run.status === "complete"
                        ? "Source processing needs attention"
                        : "Building interrupted"}
                </strong>
                <p>
                  {state.run.mode === "initial"
                    ? "Completed draft entries are saved as unreviewed."
                    : "Completed changes are saved as suggestions. Your existing events stay unchanged."}
                </p>
                {state.run.errors.map((issue, index) => (
                  <p key={index} className="chronology-warning">
                    {issue}
                  </p>
                ))}
              </div>
              <div className="chronology-inline">
                <button type="button" onClick={showSources}>
                  View source coverage
                </button>
                {running ? (
                  <button
                    type="button"
                    disabled={busy}
                    onClick={() => void perform(() => stopChronology(matterId, state.revision))}
                  >
                    Stop; keep saved work
                  </button>
                ) : (
                  state.run.status !== "complete" && (
                    <button type="button" disabled={busy} onClick={() => void prepare("resume")}>
                      Continue unfinished sources
                    </button>
                  )
                )}
              </div>
            </div>
          )}
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
                <button
                  type="button"
                  aria-expanded={filterOpen}
                  aria-controls="chronology-filters"
                  onClick={() => setFilterOpen((value) => !value)}
                >
                  <svg
                    width="17"
                    height="17"
                    viewBox="0 0 24 24"
                    fill="none"
                    stroke="currentColor"
                    strokeWidth="1.5"
                    aria-hidden="true"
                  >
                    <path d="M4 7h16M7 12h10M10 17h4" />
                  </svg>
                  Filter{activeFilterCount ? ` (${activeFilterCount})` : ""}
                </button>
                {(search || activeFilterCount > 0) && (
                  <button type="button" onClick={resetDisplay}>
                    Clear
                  </button>
                )}
              </div>
              {filterOpen && (
                <ChronologyFiltersPanel
                  key={filterReset}
                  filters={filters}
                  onChange={setFilters}
                  documentNames={[
                    ...new Set(state.events.flatMap(eventEvidence).map((evidence) => evidence.documentName)),
                  ].sort()}
                  people={people.options}
                  dateFormat={dateFormat}
                />
              )}
              <div className="chronology-table-scroll" tabIndex={0} role="region" aria-label="Chronology events">
                <table className="chronology-event-table">
                  <thead>
                    <tr>
                      {(
                        [
                          ["date", "Date"],
                          ["event", "Event"],
                          ["sources", "Sources"],
                          ["review", "Review"],
                        ] as const
                      ).map(([key, label]) => (
                        <th key={key} aria-sort={sort.key === key ? sort.direction : "none"}>
                          <button
                            type="button"
                            className="chronology-sort-heading"
                            aria-label={`Sort by ${label}`}
                            onClick={() => sortBy(key)}
                          >
                            {label}
                            <span aria-hidden="true">
                              {sort.key === key ? (sort.direction === "ascending" ? "↑" : "↓") : "↕"}
                            </span>
                          </button>
                        </th>
                      ))}
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
                                    eventEvidence(event).find((evidence) => evidence.documentName === name)?.sourceKind
                                  }
                                />
                                <span>{name}</span>
                              </span>
                            ))}
                          {new Set(eventEvidence(event).map((evidence) => evidence.documentName)).size > 2 && (
                            <span className="chronology-muted">
                              +{new Set(eventEvidence(event).map((evidence) => evidence.documentName)).size - 2} sources
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
                Showing {events.length} of {state.events.filter((event) => !event.supersededBy?.length).length} events ·
                Display filters do not limit whole-matter extraction.
              </p>
            </div>
            {detail && (
              <div
                className={`viewer-overlay chronology-detail-overlay${detailClosing ? " is-closing" : ""}${compactDetail ? " is-compact" : ""}`}
                inert={detailClosing}
                {...backdropClose(() => closeDetail())}
              >
                <div className="chronology-detail-sheet">
                  <EventDetail
                    key={detail.id}
                    event={detail}
                    closing={detailClosing}
                    chronology={state}
                    busy={busy}
                    onAction={action}
                    onEdit={(mode) => setEditing({ mode, event: detail })}
                    onClose={closeDetail}
                    onEvidence={inspect}
                    onNavigate={(eventId) => {
                      const event = state.events.find((item) => item.id === eventId)
                      if (event) select(event)
                      else setError("This earlier event record is unavailable. Its recorded history is preserved.")
                    }}
                    onAsk={() =>
                      onAsk(
                        chronologyChatReference(
                          matterId,
                          detail,
                          state.proposals.find(
                            (proposal) =>
                              proposal.targetIds.includes(detail.id) &&
                              ["pending", "deferred"].includes(proposal.status),
                          ),
                        ),
                      )
                    }
                  />
                </div>
              </div>
            )}
          </div>
          {sourcesOpen && (
            <ChronologySurface title="Sources and coverage" onClose={() => setSourcesOpen(false)}>
              {errorAlert}
              <SourceCoveragePanel chronology={state} onSourceDate={() => setSourceDate(true)} />
            </ChronologySurface>
          )}
          {updatesOpen && (
            <ChronologySurface title="Review chronology updates" onClose={() => setUpdatesOpen(false)}>
              {errorAlert}
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
                onAsk={(event, proposal) => onAsk(chronologyChatReference(matterId, event, proposal))}
              />
            </ChronologySurface>
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
                {generation === "initial"
                  ? "This first build covers the authorised matter."
                  : "Only new, changed, or unfinished source sections normally need processing. Unchanged completed sections are reused."}{" "}
                Every run starts with a fresh privilege choice.
              </p>
              <p>
                <strong>Scope:</strong> every eligible source in this matter. Display filters do not restrict
                processing.
              </p>
              {!preflight.model && (
                <p className="chronology-warning">Connect an AI provider in setup before building the chronology.</p>
              )}
              {preflight.model && (
                <p>
                  <strong>Configured model:</strong> {preflight.model.providerID} · {preflight.model.modelID}
                </p>
              )}
              <p>
                <strong>Source set:</strong> {preflight.sources.length} sources · approximately{" "}
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
                  : generation === "resume" && state?.run?.mode === "initial"
                    ? "This resumes the first draft. New usable events are saved unreviewed; proposed changes to saved entries wait for review. Completed work is kept."
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
                        ? "Check sources and suggest updates"
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
  const format = useChronologyDateFormat()
  const [dateValid, setDateValid] = useState(true)
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
        <JurisdictionSelect
          jurisdictions={names.map((name) => ({ code: name, name }))}
          selected={documentName ? [documentName] : []}
          onChange={(codes) => {
            const name = codes[0] ?? ""
            setDocumentName(name)
            setDate(chronology.sourceDates[name]?.date ?? "")
          }}
          multiple={false}
          label="Source"
          searchLabel="Search source documents"
          emptyLabel="Choose source"
          showCode={false}
        />
        <DateInput
          key={documentName}
          label="Lawyer-supplied source date"
          required
          value={date}
          format={format}
          onChange={(value) => setDate(value ?? "")}
          onValidityChange={setDateValid}
        />
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
          <button type="submit" className="primary" disabled={busy || !documentName || !date || !dateValid}>
            Record date & propose affected changes
          </button>
        </div>
      </form>
    </ChronologyDialog>
  )
}
