import { useContext, useEffect, useId, useLayoutEffect, useMemo, useRef, useState } from "react"
import { createPortal } from "react-dom"
import { UNSAFE_DataRouterContext, useBlocker } from "react-router-dom"
import { computeCell, describeRecheck, needsFill, recheckOutcome } from "../api/review-recheck"
import { citesRow, type ReviewRow } from "../api/review-answer"
import { compareColumn, comparisonInputs, MAX_COMPARE_DOCUMENTS } from "../api/review-compare"
import { matches, primaryAction, reviewAttention, type AttentionFilter, type QueueEntry } from "../api/review-attention"
import { cellKey, checkCellSources, getGrid, GridConflictError, hashQuestion, recheckPassages, saveGrid, type Grid, type GridCellData, type GridColumn, type SourceCheck } from "../api/ingest/grid"
import { humanizeError } from "../api/ingest/errors"
import { getPreferences } from "../api/ingest/preferences"
import { getMatter } from "../api/ingest/matters"
import type { Document } from "../api/ingest/documents"
import { createSession, getMessages, sendPrompt } from "../api/opencode/sessions"
import { matterClient, type Client } from "../api/opencode/client"
import ModalHeader from "./ModalHeader"
import { reviewCsv, reviewDocx, reviewExport, reviewXlsx } from "./review-export"
import { formatDateTime } from "../lib/format"
import backdropClose from "./backdropClose"
import { QualityNote } from "./DocumentUpload"
import useExit, { Closing, Exit } from "./useExit"
import DocumentViewer from "./DocumentViewer"
import { citationDocument } from "./citation-target"
import { WordWarning } from "./WordDocuments"
import useReviewGridPresence from "./useReviewGridPresence"
import ReviewStarter from "./ReviewStarter"
import { normaliseStarterQuestion, REVIEW_STARTERS } from "../review-starters"
import FillPreview from "./FillPreview"
import { CompareSummary, FilterSummary, ReviewToolbar } from "./ReviewAttention"
import { MARKER_LOOK, MarkerGlyph, ReviewMarker } from "./ReviewMarker"

// Tabular review: rows are the matter's documents, columns are natural-language
// questions, each cell is an extracted answer with a source citation. The grid is
// the durable diligence work product — columns and cells persist server-side
// (see api/ingest getGrid/saveGrid). Cells compute once and are cached; only
// empty or stale cells recompute, and `reviewed` cells are frozen. A cell left
// without a citation gets exactly one re-check of its own document as part of
// the same computation (api/review-recheck).

const POOL = 5 // concurrent cell extractions; the engine handles one prompt each
const COMPARE_POOL = 2 // concurrent column comparisons after a Fill
const COMPARE_TIMEOUT = 120_000

// The walk's next step in the detail: its primary, else whichever move is left.
const walkControl = (modal: HTMLElement | null) =>
  modal?.querySelector<HTMLButtonElement>(".review-walk-next:not(:disabled)") ?? modal?.querySelector<HTMLButtonElement>(".review-walk-bar button:not(:disabled)")

// A source check answers for the citation it was asked about, so a cell whose
// citation has since been recomputed shows no verdict until it is checked.
type Citation = NonNullable<GridCellData["citation"]>
const sourceSignature = (c: Citation) => JSON.stringify([c.documentName, c.excerpt, c.charStart, c.charEnd])
// A cell saved before answers were tied to their row can cite another
// document's passage. It is never sent to be checked: that passage cannot
// support this row's answer, whatever its own document says.
type SourceVerdict = SourceCheck | "other_document"
const DEGRADED: SourceVerdict[] = ["not_found", "missing", "other_document"]
const SOURCE_BADGE: Record<SourceVerdict, string> = { verified: "Verified", reanchored: "Re-anchored", not_found: "Source changed", missing: "Source missing", other_document: "Other document" }
const SOURCE_DETAIL: Record<SourceVerdict, string> = {
  verified: "The excerpt matches the document at the recorded position.",
  reanchored: "The excerpt is still in the document, but at a different position from the one recorded. The document may have changed since this answer.",
  not_found: "The document no longer contains this excerpt. It has changed since this answer; check the answer against the current text.",
  missing: "This document is no longer in the matter, so the excerpt cannot be checked.",
  other_document: "The cited passage is from another document, so it does not support this answer for this document. Refresh the cell to answer from this document; a reviewed cell must be unlocked first.",
}
// Cell keys are `<row document>::<column id>`; column ids never contain "::".
function citesOwnRow(key: string, citation: Citation, documents: Document[]) {
  const name = key.slice(0, key.lastIndexOf("::"))
  return citesRow(citation, { name, docPath: documents.find((doc) => doc.name === name)?.doc_path })
}

// The chat question for one cell: its question, answer and recorded citation,
// prefilled for the lawyer to edit and send. It carries what the run recorded;
// the assistant does not re-derive the answer from scratch.
function cellQuestion(docName: string, question: string, cell: GridCellData, check?: SourceVerdict) {
  const c = cell.citation
  return [
    `About this review table answer for ${docName}:`,
    `Question: ${question}`,
    `Answer: ${cell.answer}`,
    c ? `Cited source: ${c.documentName} § ${c.section}${c.structuralPath ? ` (${c.structuralPath})` : ""}` : "Cited source: none recorded",
    ...(c ? [`Excerpt: "${c.excerpt}"`] : []),
    ...(c && check ? [`Source check: ${SOURCE_BADGE[check]}. ${SOURCE_DETAIL[check]}`] : []),
    "",
    c ? "Explain how the cited source supports this answer, and anything it leaves open." : "Explain why no source supports this answer, and where the document might address the question.",
  ].join("\n")
}

// A column's stored comparison counts only while it answers this question about
// these answers; otherwise it is out of date and shows no Differs markers.
function freshComparison(column: GridColumn | undefined, input: ReturnType<typeof comparisonInputs>) {
  const comparison = column?.comparison
  return comparison && typeof input === "object" && comparison.questionHash === input.questionHash && comparison.inputs === input.signature ? comparison : undefined
}
const without = (set: ReadonlySet<string>, id: string) => {
  const next = new Set(set)
  next.delete(id)
  return next
}
const reviewRows = (documents: Document[]): ReviewRow[] => documents.map((doc) => ({ name: doc.name, docPath: doc.doc_path }))

// The production app uses a data router. Keep this separate so component-only
// harnesses without a router can still mount the review table.
function GridNavigationGuard({ canLeave, bypass }: { canLeave: () => boolean; bypass: React.RefObject<boolean> }) {
  const blocker = useBlocker(({ currentLocation, nextLocation }) => {
    if (bypass.current) return false
    if (currentLocation.pathname === nextLocation.pathname && new URLSearchParams(nextLocation.search).get("view") === "review") return false
    return !canLeave()
  })
  useEffect(() => { if (blocker.state === "blocked") blocker.reset() }, [blocker])
  return null
}

// `documentsRevision` changes whenever the matter's documents may have changed
// (an upload, a new version, a Word save), so cited excerpts are checked again.
// `initialStarter` is a starter chosen before the table opened; it opens at its
// questions once the table can be edited, and `onStarterSeated` clears it.
export default function ReviewGrid({ matterId, title, directory, documents, documentsRevision, initialStarter, onStarterSeated, onClose, onAsk }: { matterId: string; title: string; directory: string; documents: Document[]; documentsRevision: number; initialStarter?: string; onStarterSeated?: () => void; onClose: () => void; onAsk?: (prompt: string) => void }) {
  const closing = useContext(Closing)
  const dialog = useRef<HTMLDialogElement>(null)
  const detailDialog = useRef<HTMLDialogElement>(null)
  const client = useMemo<Client>(() => matterClient(directory), [directory])
  // A cell that cites a source whose extracted text is degraded says so (#239).
  const degradedCitation = (name: string) => {
    const quality = documents.find((doc) => doc.name === name)?.quality
    return quality?.status === "degraded" ? (
      <span className="review-cite-quality" title={quality.signals.map((signal) => signal.detail).join("; ")}> · cited text may be incomplete or garbled</span>
    ) : null
  }
  const [grid, setGrid] = useState<Grid>({ columns: [], cells: {} })
  const gridRef = useRef<Grid>(grid)
  const documentsRef = useRef(documents)
  documentsRef.current = documents
  const [error, setError] = useState<string>()
  const [loaded, setLoaded] = useState(false)
  const [saveProblem, setSaveProblem] = useState<string>()
  const [reloading, setReloading] = useState(false)
  const reloadInFlight = useRef(false)
  const bypassNavigation = useRef(false)
  const preflight = useRef(false)
  const dataRouter = useContext(UNSAFE_DataRouterContext)
  const saveState = useRef<{ version: number; pending: boolean; saving: boolean; blocked: boolean; epoch: number; queue: Array<{ grid: Grid; resolve?: (saved: boolean) => void }> }>({ version: 0, pending: false, saving: false, blocked: false, epoch: 0, queue: [] })
  const otherWindow = useReviewGridPresence(directory, matterId)
  const canEdit = loaded && !saveProblem
  const [running, setRunning] = useState(0) // cells currently extracting
  // Columns being compared, queued ones included, and those whose last
  // comparison failed. Neither is stored; in-flight comparisons stop on unmount.
  const [comparing, setComparing] = useState<ReadonlySet<string>>(new Set())
  const [compareFailed, setCompareFailed] = useState<ReadonlySet<string>>(new Set())
  const compareAbort = useRef<AbortController>(undefined)
  // Columns with a comparison call in flight, so one column never runs two at once.
  const comparingIds = useRef(new Set<string>())
  useEffect(() => {
    const controller = new AbortController()
    compareAbort.current = controller
    return () => controller.abort()
  }, [])
  const [starterOpen, setStarterOpen] = useState<{ starter?: string; step?: "choose" | "questions" }>()
  // An arriving starter opens once, at its questions: the lawyer already chose
  // it, and Back returns to the chooser. A table that already has questions
  // appends them. An unknown starter is dropped.
  const seatedStarter = useRef<string>(undefined)
  useEffect(() => {
    if (!initialStarter || !canEdit || seatedStarter.current === initialStarter) return
    seatedStarter.current = initialStarter
    onStarterSeated?.()
    if (!REVIEW_STARTERS.some((starter) => starter.id === initialStarter)) return
    setError(undefined)
    setStarterOpen({ starter: initialStarter, step: "questions" })
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [initialStarter, canEdit])
  const [fillPreview, setFillPreview] = useState<Array<{ docName: string; column: { id: string; question: string } }>>()
  const [checkingFill, setCheckingFill] = useState(false)
  // Filter and search change only what is shown. Rows and columns with a match
  // stay shown until they change, so a cell reviewed in place does not jump away.
  const [filter, setFilter] = useState<AttentionFilter>()
  const [query, setQuery] = useState("")
  const shownView = useRef({ key: "", rows: new Set<string>(), columns: new Set<string>() })
  // The detail overlay identifies its cell by key and reads it live from `grid`,
  // so the comment thread updates in place as comments are added or removed.
  const [detail, setDetail] = useState<{ docName: string; columnId: string; question: string }>()
  const [draft, setDraft] = useState("") // comment being composed in the overlay
  const [commentSaving, setCommentSaving] = useState(false)
  // Each dialog closes, and returns focus, as its exit starts.
  const [detailShown, detailClosing] = useExit(detail)
  const detailKey = detailShown ? cellKey(detailShown.docName, detailShown.columnId) : undefined
  const detailCell = detailKey ? grid.cells[detailKey] : undefined
  // Review next walks the queue as it stood when the walk started, inside the
  // same detail; `at` is the open cell's place in it. Opening a cell from the
  // table leaves the walk. `walkDone` counts finished walks for their status.
  const [walk, setWalk] = useState<{ entries: QueueEntry[]; at: number }>()
  const [walkDone, setWalkDone] = useState(0)
  const detailIds = useId()
  const inWalk = !!walk && walk.entries[walk.at]?.key === detailKey
  const walkAt = inWalk ? walk.at : -1
  // A comment being written belongs to this cell, so the walk waits for it.
  const composing = !!draft.trim() || commentSaving
  // The cell whose cited document is open over its detail. Closing the viewer
  // returns to that detail, with focus back on the citation that opened it.
  const [viewing, setViewing] = useState<string>()
  const viewerOpen = !!detail && viewing === cellKey(detail.docName, detail.columnId)
  const sourceLink = useRef<HTMLButtonElement>(null)
  // Source checks of every cited cell, keyed by cell, each with the citation
  // it was made for; a checked cell without a verdict could not be checked.
  // `sourceRound` re-checks after the document viewer changes a document (an
  // accepted or rejected redline).
  const [sourceChecks, setSourceChecks] = useState<{ key: string; checks: Record<string, { signature: string; check?: SourceCheck }>; failed?: boolean }>()
  const [sourceRound, setSourceRound] = useState(0)
  const sources = useMemo(() => Object.entries(grid.cells).flatMap(([key, cell]) => cell.citation && citesOwnRow(key, cell.citation, documents)
    ? [{ key, signature: sourceSignature(cell.citation), documentName: cell.citation.documentName, excerpt: cell.citation.excerpt, charStart: cell.citation.charStart, charEnd: cell.citation.charEnd }]
    : []), [grid.cells, documents])
  // An edit invalidates the verdict in the same render, before the next effect
  // starts. Late replies remain tied to the document round they actually checked.
  const sourcesKey = JSON.stringify([matterId, sourceRound, documentsRevision, sources])
  // Each round re-checks every cited cell, so while one runs no cell has a verdict.
  const checkingSources = sources.length > 0 && sourceChecks?.key !== sourcesKey
  const sourceEntry = (key: string, cell: GridCellData | undefined): { signature: string; check?: SourceVerdict } | undefined => {
    if (cell?.citation && !citesOwnRow(key, cell.citation, documents)) return { signature: sourceSignature(cell.citation), check: "other_document" }
    const entry = sourceChecks?.key === sourcesKey ? sourceChecks.checks[key] : undefined
    return cell?.citation && entry?.signature === sourceSignature(cell.citation) ? entry : undefined
  }
  const sourceCheck = (key: string, cell: GridCellData | undefined) => sourceEntry(key, cell)?.check

  useLayoutEffect(() => {
    if (closing) return
    const opener = document.activeElement instanceof HTMLElement ? document.activeElement : null
    const modal = dialog.current
    modal?.showModal()
    return () => {
      modal?.close()
      if (opener?.isConnected) opener.focus({ preventScroll: true })
    }
  }, [closing])

  // Opens and closes with the detail, not with each cell it shows, so the walk
  // moves between cells without reopening the modal or taking focus.
  const detailOpen = !!detail
  useLayoutEffect(() => {
    if (!detailOpen) return
    const opener = document.activeElement instanceof HTMLElement ? document.activeElement : null
    const modal = detailDialog.current
    modal?.showModal()
    ;(walkControl(modal) ?? modal?.querySelector<HTMLTextAreaElement>("textarea"))?.focus({ preventScroll: true })
    return () => {
      modal?.close()
      if (opener?.isConnected) opener.focus({ preventScroll: true })
    }
  }, [detailOpen])

  // A move that removes or disables the focused walk control hands focus to
  // the next step, so it never drops out of the detail.
  useLayoutEffect(() => {
    const modal = detailDialog.current
    if (walkAt < 0 || !modal) return
    const active = document.activeElement
    if (active && modal.contains(active) && !active.matches(":disabled")) return
    walkControl(modal)?.focus({ preventScroll: true })
  }, [walkAt, canEdit])

  // A finished walk hands focus to the table's next step and says it is done.
  useLayoutEffect(() => {
    if (!walkDone) return
    const main = dialog.current?.querySelector<HTMLElement>(".review-toolbar .review-primary")
    ;(main?.matches("button") ? main : main?.querySelector<HTMLElement>("button"))?.focus({ preventScroll: true })
  }, [walkDone])
  useEffect(() => {
    if (!walkDone) return
    const timer = window.setTimeout(() => setWalkDone(0), 6000)
    return () => window.clearTimeout(timer)
  }, [walkDone])

  // Escape closes only the open cell detail. Consumed here, on document (before
  // ModalHeader's window listeners) and wherever focus is, so neither those
  // listeners nor the dialog's native cancel also close the table on that press.
  useEffect(() => {
    if (!detail) return
    const onKey = (e: KeyboardEvent) => {
      if (e.key !== "Escape" || e.defaultPrevented) return
      e.preventDefault()
      setDetail(undefined)
    }
    document.addEventListener("keydown", onKey)
    return () => document.removeEventListener("keydown", onKey)
  }, [detail])

  useEffect(() => {
    let cancelled = false
    setLoaded(false)
    getGrid(matterId).then((g) => {
      if (cancelled) return
      gridRef.current = g
      setGrid(g)
      saveState.current = { version: g.version ?? 0, pending: false, saving: false, blocked: false, epoch: saveState.current.epoch + 1, queue: [] }
      setLoaded(true)
    }).catch((e) => !cancelled && setError(humanizeError(e)))
    return () => { cancelled = true; saveState.current.epoch++ }
  }, [matterId])

  // A failed or in-flight save is not safely disposable just because the tab is closed.
  useEffect(() => {
    const warn = (event: BeforeUnloadEvent) => {
      if (!saveState.current.pending && !saveState.current.saving && !preflight.current && !running && !commentSaving && !draft.trim()) return
      event.preventDefault()
      event.returnValue = ""
    }
    window.addEventListener("beforeunload", warn)
    return () => window.removeEventListener("beforeunload", warn)
  }, [running, commentSaving, draft])

  // Check each cited cell's excerpt against its document's current text
  // whenever the set of citations changes. Read-only: a verdict never edits a cell.
  useEffect(() => {
    if (!sources.length) return
    let cancelled = false
    checkCellSources(matterId, sources.map(({ signature: _, ...cell }) => cell)).then((checks) => {
      if (cancelled) return
      setSourceChecks({ key: sourcesKey, checks: Object.fromEntries(sources.map(({ key, signature }) => [key, { signature, check: checks[key] }])) })
    }, () => !cancelled && setSourceChecks({ key: sourcesKey, checks: {}, failed: true }))
    return () => { cancelled = true }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [matterId, sourcesKey, sourceRound])

  // Serialize whole-grid saves. A failed response can be ambiguous (the server
  // might have written it), so stop and keep the local copy for explicit recovery.
  async function flush() {
    const state = saveState.current
    if (state.saving || state.blocked || !state.pending) return
    state.saving = true
    const epoch = state.epoch
    while (state.queue.length && !state.blocked && state.epoch === epoch) {
      const queued = state.queue.shift()!
      state.pending = state.queue.length > 0
      const snapshot = { ...queued.grid, version: state.version }
      try {
        const saved = await saveGrid(matterId, snapshot)
        if (state.epoch !== epoch) {
          queued.resolve?.(false)
          break
        }
        state.version = saved.version ?? state.version + 1
        gridRef.current = { ...gridRef.current, version: state.version }
        setGrid(gridRef.current)
        queued.resolve?.(true)
      } catch (e) {
        if (state.epoch !== epoch) {
          queued.resolve?.(false)
          break
        }
        state.queue.unshift(queued)
        state.pending = true
        state.blocked = true
        queued.resolve?.(false)
        state.queue.forEach((item) => item.resolve?.(false))
        setSaveProblem(e instanceof GridConflictError
          ? "This table conflicts with the saved table. Your changes are still shown here and have not been saved. Extractions in progress may need to be run again after reloading."
          : "The table could not be saved. Your changes are still shown here; the save may or may not have reached the server. Extractions in progress may need to be run again after reloading.")
      }
    }
    state.saving = false
  }

  // One local writer for concurrent cell completions and user changes.
  function commit(next: Grid): Promise<boolean> {
    if (!loaded || saveState.current.blocked) return Promise.resolve(false)
    gridRef.current = next
    setGrid(next)
    saveState.current.pending = true
    // Keep each intentional transition, including an unlock before a later
    // edit, so the server can distinguish it from replacing a reviewed cell.
    return new Promise((resolve) => {
      saveState.current.queue.push({ grid: next, resolve })
      void flush()
    })
  }

  async function reloadGrid() {
    if (reloadInFlight.current) return
    if (!window.confirm("Reload the saved table? Your unsaved changes in this window will be discarded. You can download a copy first.")) return
    reloadInFlight.current = true
    setReloading(true)
    const state = saveState.current
    const epoch = state.epoch
    let completed = false
    try {
      const latest = await getGrid(matterId)
      if (saveState.current !== state || state.epoch !== epoch) return
      state.epoch++
      completed = true
      state.version = latest.version ?? 0
      state.pending = false
      // Discarded changes resolve as unsaved, so nothing waits on them.
      state.queue.forEach((item) => item.resolve?.(false))
      state.queue = []
      state.blocked = false
      state.saving = false
      gridRef.current = latest
      setGrid(latest)
      setDetail(undefined)
      setDraft("")
      setSaveProblem(undefined)
      setError(undefined)
    } catch (e) {
      if (saveState.current === state && state.epoch === epoch) setError(humanizeError(e, "Could not reload the review table."))
    } finally {
      reloadInFlight.current = false
      if (saveState.current === state && state.epoch === epoch + (completed ? 1 : 0)) setReloading(false)
    }
  }

  function downloadPending() {
    const url = URL.createObjectURL(new Blob([JSON.stringify(gridRef.current, null, 2)], { type: "application/json" }))
    const link = document.createElement("a")
    link.href = url
    link.download = "review-table-unsaved.json"
    link.click()
    URL.revokeObjectURL(url)
  }

  function canLeaveGrid() {
    const state = saveState.current
    return !(state.pending || state.saving || preflight.current || running || commentSaving || draft.trim()) || window.confirm("There are unsaved changes or work in progress. Leave the review table anyway?")
  }

  function closeGrid() {
    if (canLeaveGrid()) allowNavigation(onClose)
  }

  function allowNavigation(action: () => void) {
    bypassNavigation.current = true
    try { action() } finally { queueMicrotask(() => { bypassNavigation.current = false }) }
  }

  function commitCell(key: string, data: GridCellData) {
    commit({ ...gridRef.current, cells: { ...gridRef.current.cells, [key]: data } })
  }

  function addColumn() {
    const question = window.prompt("New column question (e.g. What is the governing law?)")?.trim()
    if (!question) return
    commit({ ...gridRef.current, columns: [...gridRef.current.columns, { id: crypto.randomUUID().slice(0, 6), question }] })
  }

  function renameColumn(id: string) {
    const col = gridRef.current.columns.find((c) => c.id === id)
    const question = window.prompt("Edit column question", col?.question)?.trim()
    if (!question) return
    commit({ ...gridRef.current, columns: gridRef.current.columns.map((c) => (c.id === id ? { ...c, question } : c)) })
  }

  function deleteColumn(id: string) {
    const cells = { ...gridRef.current.cells }
    for (const doc of documents) delete cells[cellKey(doc.name, id)]
    commit({ ...gridRef.current, columns: gridRef.current.columns.filter((c) => c.id !== id), cells })
    dialog.current?.querySelector<HTMLButtonElement>(".review-add-button")?.focus()
  }

  function setReviewed(key: string, reviewed: boolean) {
    const cell = gridRef.current.cells[key]
    if (!cell || (cell.status === "reviewed") === reviewed) return
    commitCell(key, { ...cell, status: reviewed ? "reviewed" : "filled" })
  }

  function showDetail(docName: string, columnId: string) {
    const column = gridRef.current.columns.find((col) => col.id === columnId)
    if (!column) return
    setDraft("")
    setViewing(undefined)
    setDetail({ docName, columnId, question: column.question })
  }

  function openDetail(docName: string, columnId: string) {
    setWalk(undefined)
    showDetail(docName, columnId)
  }

  async function addComment() {
    const text = draft.trim()
    if (!detail || !text || commentSaving) return
    const key = cellKey(detail.docName, detail.columnId)
    const submitted = draft
    const at = Date.now()
    const epoch = saveState.current.epoch
    setCommentSaving(true)
    try {
      // Read at creation, rather than labeling stored comments with a name
      // that may have changed since they were written. No name (or unavailable
      // settings) leaves an honest, unattributed comment.
      const author = await getPreferences().then((prefs) => prefs.attorney?.trim() || undefined, () => undefined)
      const cell = gridRef.current.cells[key]
      if (!cell || saveState.current.blocked || saveState.current.epoch !== epoch) return
      commitCell(key, {
        ...cell,
        comments: [...(cell.comments ?? []), { id: crypto.randomUUID().slice(0, 6), text, at, ...(author ? { author } : {}) }],
      })
      setDraft((current) => current === submitted ? "" : current)
    } finally {
      setCommentSaving(false)
    }
  }

  function deleteComment(id: string) {
    if (!detail) return
    const key = cellKey(detail.docName, detail.columnId)
    const cell = gridRef.current.cells[key]
    commitCell(key, { ...cell, comments: (cell.comments ?? []).filter((c) => c.id !== id) })
    detailDialog.current?.querySelector<HTMLTextAreaElement>("textarea")?.focus({ preventScroll: true })
  }

  // One extract turn in a scoped, throwaway session titled so it stays out of
  // the sidebar's conversation list (which filters "doc.haus"-titled sessions).
  // A tool-using turn emits several assistant messages (one per model step): the
  // search-document call lands on an earlier one, the final answer text on the
  // last. Return every assistant message so the citation is not dropped.
  async function extractTurn(title: string, prompt: string, permissions?: Parameters<typeof createSession>[3]) {
    const session = await createSession(client, title, undefined, permissions)
    await sendPrompt(client, session.id, "extract", prompt)
    return (await getMessages(client, session.id)).filter((m) => m.info.role === "assistant")
  }

  // Extract one cell. A cell left without a citation gets its one re-check: the
  // deterministic pass over this document only, then, if it found candidate
  // passages, one extract call that may use no tools and sees only those.
  function extract(docName: string, question: string): Promise<GridCellData> {
    const docPath = documentsRef.current.find((doc) => doc.name === docName)?.doc_path
    return computeCell({ document: docName, docPath, question, questionHash: hashQuestion(question) }, {
      extract: (prompt) => extractTurn(`doc.haus grid: ${docName}`, prompt),
      passages: (request) => recheckPassages(matterId, { document: docName, question, ...request }),
      recheck: (prompt) => extractTurn(`doc.haus grid re-check: ${docName}`, prompt, [{ action: "*", resource: "*", effect: "deny" }]),
    })
  }

  // Run a set of (document, column) targets through a small concurrency pool so a
  // large grid does not open hundreds of sessions at once.
  async function fill(targets: Array<{ docName: string; column: { id: string; question: string } }>) {
    if (targets.length === 0 || !loaded || saveState.current.blocked) return
    setError(undefined)
    setRunning((n) => n + targets.length)
    const fillEpoch = saveState.current.epoch
    let next = 0
    const worker = async () => {
      while (next < targets.length) {
        const t = targets[next++]
        const key = cellKey(t.docName, t.column.id)
        const original = gridRef.current.cells[key]
        const initialColumn = gridRef.current.columns.find((col) => col.id === t.column.id)
        // Comments can be added during extraction. Compare all other cell data,
        // and carry the newest comments forward if the result is still applicable.
        const withoutComments = (cell: GridCellData | undefined) => {
          if (!cell) return undefined
          const { comments: _, ...rest } = cell
          return JSON.stringify(rest)
        }
        try {
          if (saveState.current.epoch !== fillEpoch || saveState.current.blocked || !initialColumn || initialColumn.question !== t.column.question || original?.status === "reviewed") continue
          const data = await extract(t.docName, t.column.question)
          const current = gridRef.current.cells[key]
          const column = gridRef.current.columns.find((col) => col.id === t.column.id)
          if (saveState.current.epoch !== fillEpoch || saveState.current.blocked || !column || column.question !== t.column.question || current?.status === "reviewed" || withoutComments(current) !== withoutComments(original)) continue
          commitCell(key, { ...data, comments: current?.comments })
        } catch (error) {
          setError(humanizeError(error, "Extraction failed. Try the unfinished cells again."))
        } finally {
          setRunning((n) => n - 1)
        }
      }
    }
    await Promise.all(Array.from({ length: Math.min(POOL, targets.length) }, worker))
  }

  // Compare one column's cited answers. The result is written only if the table
  // still holds the answers that were sent; otherwise the comparison stays out
  // of date. It is stored on the column, never in a cell.
  async function compareOne(columnId: string, epoch: number) {
    const unmounted = compareAbort.current?.signal
    if (!unmounted || unmounted.aborted || saveState.current.epoch !== epoch || saveState.current.blocked) return
    const sent = gridRef.current.columns.find((col) => col.id === columnId)
    const input = comparisonInputs(gridRef.current, columnId, reviewRows(documentsRef.current))
    if (typeof input !== "object" || freshComparison(sent, input)) return
    setCompareFailed((failed) => without(failed, columnId))
    try {
      const result = await compareColumn(client, input, AbortSignal.any([AbortSignal.timeout(COMPARE_TIMEOUT), unmounted]))
      if (unmounted.aborted) return
      if (!result) {
        setCompareFailed((failed) => new Set(failed).add(columnId))
        return
      }
      const current = gridRef.current
      const column = current.columns.find((col) => col.id === columnId)
      const now = comparisonInputs(current, columnId, reviewRows(documentsRef.current))
      if (saveState.current.epoch !== epoch || saveState.current.blocked || !column || column.question !== input.question || typeof now !== "object" || now.signature !== input.signature) return
      await commit({ ...current, columns: current.columns.map((col) => (col.id === columnId ? { ...col, comparison: result } : col)) })
    } catch {
      if (!unmounted.aborted) setCompareFailed((failed) => new Set(failed).add(columnId))
    }
  }

  // Each column with two or more cited answers and no up-to-date comparison,
  // through a pool of two; a column above the document cap is not compared.
  async function compareColumns(columnIds: string[], epoch: number) {
    const ids = columnIds.filter((id) => {
      if (comparingIds.current.has(id)) return false
      const input = comparisonInputs(gridRef.current, id, reviewRows(documentsRef.current))
      return typeof input === "object" && !freshComparison(gridRef.current.columns.find((col) => col.id === id), input)
    })
    if (!ids.length) return
    ids.forEach((id) => comparingIds.current.add(id))
    setComparing((current) => new Set([...current, ...ids]))
    let next = 0
    const worker = async () => {
      while (next < ids.length) {
        const id = ids[next++]
        try {
          await compareOne(id, epoch)
        } finally {
          comparingIds.current.delete(id)
          setComparing((current) => without(current, id))
        }
      }
    }
    await Promise.all(Array.from({ length: Math.min(COMPARE_POOL, ids.length) }, worker))
  }

  function fillTargets() {
    const current = gridRef.current
    return documentsRef.current.flatMap((doc) =>
      current.columns
        .filter((col) => {
          // Compute blank, stale and never re-checked negative cells; never touch a
          // reviewed (locked) cell or one already re-checked for this question.
          return needsFill(current.cells[cellKey(doc.name, col.id)], hashQuestion(col.question))
        })
        .map((column) => ({ docName: doc.name, column })),
    )
  }

  function sameDocuments(current: Document[]) {
    const names = (docs: Document[]) => docs.map((doc) => doc.name).sort()
    return JSON.stringify(names(current)) === JSON.stringify(names(documentsRef.current))
  }

  async function confirmFill() {
    if (!fillPreview || checkingFill) return
    const state = saveState.current
    if (state.pending || state.saving || state.blocked) {
      setError("Wait for the table to save, or resolve the save problem, before filling.")
      setFillPreview(undefined)
      return
    }
    setCheckingFill(true)
    preflight.current = true
    const epoch = state.epoch
    const version = state.version
    try {
      const [saved, matter] = await Promise.all([getGrid(matterId), getMatter(matterId)])
      if (state.epoch !== epoch) return
      if (!sameDocuments(matter.documents)) {
        setError("The matter's documents changed. Close the review table and reload the matter, then review the fill again.")
        setFillPreview(undefined)
        return
      }
      const current = fillTargets()
      // A comparison saved from this window during the check moves the version
      // within this range; only a version outside it was saved elsewhere.
      const savedVersion = saved.version ?? 0
      if (savedVersion < version || savedVersion > state.version) {
        state.blocked = true
        setSaveProblem("This table changed elsewhere. Reload the saved table before filling; download a copy first if needed.")
        setFillPreview(undefined)
        return
      }
      if (state.blocked || state.pending || state.saving || JSON.stringify(current) !== JSON.stringify(fillPreview)) {
        setError("The fill targets changed. Review the fill again before starting.")
        setFillPreview(undefined)
        return
      }
      setFillPreview(undefined)
      // Only a reviewed Fill compares; a single cell's Fill or Refresh does not.
      void fill(current).then(() => compareColumns([...new Set(current.map((target) => target.column.id))], epoch))
    } catch (e) {
      setError(humanizeError(e, "Could not check the saved table. Try again before filling."))
      setFillPreview(undefined)
    } finally {
      preflight.current = false
      setCheckingFill(false)
    }
  }

  async function addStarterQuestions(questions: string[]) {
    const state = saveState.current
    const epoch = state.epoch
    preflight.current = true
    try {
      const matter = await getMatter(matterId)
      if (state.epoch !== epoch) return false
      if (!sameDocuments(matter.documents)) {
        setError("The matter's documents changed. Close the review table and reload the matter before adding the starter.")
        return false
      }
    } catch (e) {
      setError(humanizeError(e, "Could not check the matter's documents. Try again."))
      return false
    } finally {
      preflight.current = false
    }
    const existing = new Set(gridRef.current.columns.map((col) => normaliseStarterQuestion(col.question)))
    const unique = questions.map((question) => question.trim()).filter((question) => {
      const key = normaliseStarterQuestion(question)
      if (!key || existing.has(key)) return false
      existing.add(key)
      return true
    })
    if (!unique.length || unique.length !== questions.length || !documentsRef.current.length) {
      setError("Questions or documents changed. Review the starter and try again.")
      return false
    }
    return commit({ ...gridRef.current, columns: [...gridRef.current.columns, ...unique.map((question) => ({ id: crypto.randomUUID(), question }))] })
  }

  // Excel, CSV and Word read one model of the grid (./review-export) and save
  // through the same browser download.
  function exportAs(format: "xlsx" | "csv" | "docx") {
    const model = reviewExport(title, grid, documents.map((doc) => doc.name))
    const slug = title.trim().toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, "") || "matter"
    const save = (data: BlobPart, type: string) => {
      const url = URL.createObjectURL(new Blob([data], { type }))
      const a = document.createElement("a")
      a.href = url
      a.download = `${slug}-tabular-review.${format}`
      a.click()
      URL.revokeObjectURL(url)
    }
    if (format === "xlsx") save(reviewXlsx(model), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    if (format === "csv") save(reviewCsv(model), "text/csv;charset=utf-8")
    if (format === "docx")
      reviewDocx(model)
        .then((bytes) => save(bytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document"))
        .catch((e) => setError(humanizeError(e, "Could not build the Word export. Try again.")))
  }

  const compareInputs = useMemo(() => new Map(grid.columns.map((col) => [col.id, comparisonInputs(grid, col.id, reviewRows(documents))])), [grid, documents])
  const comparisonOf = (column: GridColumn) => freshComparison(column, compareInputs.get(column.id))
  const differs = (docName: string, column: GridColumn) => {
    const compared = comparisonOf(column)?.cells[docName]
    return compared?.label === "differs" ? compared.reason : undefined
  }
  const attention = reviewAttention({ rows: documents.map((doc) => doc.name), columns: grid.columns, cells: grid.cells, source: sourceCheck, checkingSources, differs })
  // The same cells fillTargets() would compute, counted from this render.
  const fillCount = documents.reduce((count, doc) => count + grid.columns.filter((col) => needsFill(grid.cells[cellKey(doc.name, col.id)], hashQuestion(col.question))).length, 0)
  const fillBlanks = documents.some((doc) => grid.columns.some((col) => !grid.cells[cellKey(doc.name, col.id)]))
  const primary = primaryAction(attention.progress, fillCount)
  const narrowed = !!filter || !!query.trim()
  const viewKey = JSON.stringify([filter, query.trim().toLowerCase()])
  if (shownView.current.key !== viewKey) shownView.current = { key: viewKey, rows: new Set(), columns: new Set() }
  let matching = 0
  for (const entry of attention.cells.values()) {
    if (!matches(entry, filter, query)) continue
    matching++
    shownView.current.rows.add(entry.docName)
    shownView.current.columns.add(entry.columnId)
  }
  const shownDocuments = narrowed ? documents.filter((doc) => shownView.current.rows.has(doc.name)) : documents
  // With no match the questions stay as headings above "No matching cells".
  const shownColumns = narrowed && shownDocuments.length ? grid.columns.filter((col) => shownView.current.columns.has(col.id)) : grid.columns
  const clearView = () => {
    setFilter(undefined)
    setQuery("")
    dialog.current?.querySelector<HTMLInputElement>(".review-toolbar input[type=search]")?.focus()
  }
  // Review next walks the queued cells the view shows, else the whole queue.
  const reviewNext = () => {
    const shown = attention.queue.filter((entry) => matches(attention.cells.get(entry.key)!, filter, query))
    const entries = shown.length ? shown : attention.queue
    if (!entries.length) return
    setWalkDone(0)
    setWalk({ entries, at: 0 })
    showDetail(entries[0].docName, entries[0].columnId)
  }
  // Forward passes cells that no longer need attention (reviewed with nothing
  // flagged, or gone); back returns to any cell still in the table, to look
  // again. Past the last cell the walk ends.
  const stillQueued = new Set(attention.queue.map((entry) => entry.key))
  const inTable = (entry: QueueEntry) => attention.cells.has(entry.key) && !!grid.cells[entry.key]
  const walkTo = (step: 1 | -1) => {
    if (!walk || composing) return
    for (let at = walk.at + step; at >= 0 && at < walk.entries.length; at += step) {
      const entry = walk.entries[at]
      if (step > 0 ? !stillQueued.has(entry.key) : !inTable(entry)) continue
      setWalk({ entries: walk.entries, at })
      showDetail(entry.docName, entry.columnId)
      return
    }
    if (step < 0) return
    setDetail(undefined)
    setWalkDone((count) => count + 1)
  }
  const markReviewedNext = () => {
    if (!inWalk || !detailKey || !canEdit || composing) return
    setReviewed(detailKey, true)
    walkTo(1)
  }
  const walkEntry = inWalk && detailKey ? attention.cells.get(detailKey) : undefined
  const walkReason = detailKey && stillQueued.has(detailKey) ? attention.queue.find((entry) => entry.key === detailKey)!.reason : undefined
  // A degraded source is always the cell's marker, which the walk's reason
  // already names and explains, so the citation repeats neither.
  const walkExplains = (check: SourceVerdict) => inWalk && !!walkReason && DEGRADED.includes(check)
  // The open cell's comparison, unless the walk's reason already gives it.
  const detailColumn = detailShown && grid.columns.find((col) => col.id === detailShown.columnId)
  const detailCompared = detailColumn && comparisonOf(detailColumn)?.cells[detailShown.docName]
  const detailDiffers = detailCompared?.label === "differs" && !(inWalk && walkReason && walkEntry?.marker?.state === "differs") ? detailCompared.reason : undefined
  // The header's one quiet line on its comparison: the summary, which opens the
  // differing cells, or why there is none and, where one can run, a compare.
  const compareLine = (col: GridColumn) => {
    const input = compareInputs.get(col.id)
    if (!input) return null
    if (input === "too_many") return <p className="review-compare-note">Not compared: more than {MAX_COMPARE_DOCUMENTS} documents</p>
    if (comparing.has(col.id)) return <p className="review-compare-note">Comparing answers…</p>
    const fresh = comparisonOf(col)
    if (fresh) return <CompareSummary comparison={fresh} onOpen={(docName) => openDetail(docName, col.id)} />
    const failed = compareFailed.has(col.id)
    return (
      <p className="review-compare-note">
        {failed ? "Could not compare · " : col.comparison ? "Comparison out of date · " : ""}
        <button type="button" className="linklike" disabled={!canEdit || running > 0} title="One AI call compares the cited answers across documents" onClick={() => void compareColumns([col.id], saveState.current.epoch)}>
          {failed || col.comparison ? "Compare again" : "Compare answers"}
        </button>
      </p>
    )
  }

  return createPortal(
    <dialog
      ref={dialog}
      className={`viewer-panel review-modal${closing ? " is-closing" : ""}`}
      inert={closing}
      aria-label={`Review table — ${title}`}
      onCancel={(event) => { event.preventDefault(); detail ? setDetail(undefined) : closeGrid() }}
      {...backdropClose((event) => {
        const bounds = event.currentTarget.getBoundingClientRect()
        if (event.clientX < bounds.left || event.clientX > bounds.right || event.clientY < bounds.top || event.clientY > bounds.bottom) closeGrid()
      })}
    >
      <ModalHeader title={`Review table — ${title}`} onClose={() => detail ? setDetail(undefined) : closeGrid()} />
      {dataRouter && <GridNavigationGuard canLeave={canLeaveGrid} bypass={bypassNavigation} />}
      <div className="card review">
        <WordWarning />
        {otherWindow && <p role="status">This review table is open in another window.</p>}
        <ReviewToolbar
          model={attention}
          checking={checkingSources}
          running={running}
          comparing={comparing.size}
          primary={primary}
          fillCount={fillCount}
          fillBlanks={fillBlanks}
          startDisabled={!canEdit || running > 0}
          fillDisabled={!canEdit || running > 0}
          canExport={documents.length > 0 && grid.columns.length > 0}
          searchable={attention.cells.size > 0}
          filter={filter}
          onFilter={setFilter}
          query={query}
          onQuery={setQuery}
          onStarter={() => { setError(undefined); setStarterOpen({}) }}
          onFill={() => { setError(undefined); setFillPreview(fillTargets()) }}
          onReviewNext={reviewNext}
          onExport={exportAs}
          notice={walkDone ? "End of the review list" : undefined}
        />
        {narrowed && <FilterSummary shown={matching} total={attention.cells.size} filter={filter} onClear={clearView} />}

        {error && <p role="alert">{error}</p>}
        {saveProblem && <div role="alert"><p>{saveProblem}</p><p>Export the table above or download your local copy before reloading the saved table.</p><button type="button" onClick={downloadPending}>Download unsaved copy</button> <button type="button" disabled={reloading} onClick={reloadGrid}>{reloading ? "Reloading…" : "Reload saved table"}</button></div>}

        {starterOpen && <ReviewStarter documents={documents} existingQuestions={grid.columns.map((column) => column.question)} problem={saveProblem ?? error} initialStarter={starterOpen.starter} initialStep={starterOpen.step} onCancel={() => setStarterOpen(undefined)} onConfirm={addStarterQuestions} />}
        {fillPreview && <FillPreview targets={fillPreview} checking={checkingFill} onCancel={() => setFillPreview(undefined)} onConfirm={confirmFill} />}

        {grid.columns.length === 0 && (
          <p className="review-hint">Use + to add a question for every document.</p>
        )}
        <div className="review-scroll" role="region" aria-label="Review table" tabIndex={0}>
          <table className={`review-table${grid.columns.length === 0 ? " is-empty" : ""}`} style={shownColumns.length ? { minWidth: 304 + shownColumns.length * 280 } : undefined} aria-label="Document review">
            <thead>
              <tr>
                <th className="review-doc-col" scope="col">Document</th>
                {shownColumns.map((col) => (
                  <th key={col.id} scope="col">
                    <div className="review-col-head">
                      <button className="review-col-q" disabled={!canEdit} onClick={() => renameColumn(col.id)} aria-label={`Edit question: ${col.question}`} title="Edit question">
                        {col.question}
                      </button>
                      <button className="icon-btn review-col-x" disabled={!canEdit} onClick={() => deleteColumn(col.id)} aria-label={`Delete column: ${col.question}`} title="Delete column">
                        ×
                      </button>
                    </div>
                    {compareLine(col)}
                  </th>
                ))}
                <th className="review-add-col" scope="col">
                  <button className="review-add-button" type="button" disabled={!canEdit} onClick={addColumn} aria-label="Add column" title="Add column">
                    <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeLinecap="round" aria-hidden="true">
                      <path d="M12 5v14M5 12h14" />
                    </svg>
                  </button>
                </th>
              </tr>
            </thead>
            <tbody>
              {shownDocuments.map((doc) => (
                <tr key={doc.id}>
                  <th className="review-doc-col" scope="row">{doc.name}<QualityNote quality={doc.quality} /></th>
                  {shownColumns.map((col) => {
                    const key = cellKey(doc.name, col.id)
                    const cell = grid.cells[key]
                    const entry = attention.cells.get(key)!
                    const dimmed = narrowed && !matches(entry, filter, query)
                    const check = sourceCheck(key, cell)
                    // The inset bar takes the marker's tone: a changed source is a warning, a missing or foreign one danger.
                    const degraded = !!check && DEGRADED.includes(check) && !!entry.marker
                    return (
                      <td key={col.id} className={`review-cell${cell?.status === "reviewed" ? " reviewed" : ""}${degraded ? ` degraded is-${MARKER_LOOK[entry.marker!.state].tone}` : ""}${dimmed ? " is-dimmed" : ""}`}>
                        {!cell ? (
                          <button
                            className="review-cell-fill"
                            disabled={!canEdit || running > 0}
                            onClick={() => fill([{ docName: doc.name, column: col }])}
                          >
                            Fill
                          </button>
                        ) : (
                          <div className="review-cell-body">
                            <button
                              className="review-answer"
                              onClick={() => openDetail(doc.name, col.id)}
                            >
                              {cell.answer}
                              {/* On its own line; the row names the document unless the citation is from another one. */}
                              {cell.citation && (
                                <span className="review-cite">
                                  <span className="sr-only"> Source:</span> {citesRow(cell.citation, { name: doc.name, docPath: doc.doc_path }) ? "" : `${cell.citation.documentName} `}§ {cell.citation.section}
                                  {degradedCitation(cell.citation.documentName)}
                                </span>
                              )}
                            </button>
                            <ReviewMarker entry={entry} recheck={cell.recheck && recheckOutcome(cell.recheck)} onOpen={() => openDetail(doc.name, col.id)} />
                            <div className="review-cell-actions">
                              <button
                                className={`icon-btn review-note${cell.comments?.length ? " has-comments" : ""}`}
                                aria-label={`${cell.comments?.length ?? 0} comments on ${doc.name}: ${col.question}`}
                                title={cell.comments?.map((c) => `${c.author || "Unattributed"}: ${c.text}`).join("\n") || "Add a comment"}
                                onClick={() => openDetail(doc.name, col.id)}
                              >
                                <svg width="12" height="12" viewBox="0 0 16 16" aria-hidden="true">
                                  <path
                                    d="M3 2.5h10A1.5 1.5 0 0 1 14.5 4v6a1.5 1.5 0 0 1-1.5 1.5H8.5L5 14.5v-3H3A1.5 1.5 0 0 1 1.5 10V4A1.5 1.5 0 0 1 3 2.5Z"
                                    fill="none"
                                    stroke="currentColor"
                                    strokeWidth="1.5"
                                    strokeLinejoin="round"
                                  />
                                </svg>
                                {cell.comments?.length ?? 0}
                              </button>
                              <button className={`icon-btn review-lock${cell.status === "reviewed" ? " locked" : ""}`} disabled={!canEdit} aria-pressed={cell.status === "reviewed"} title={cell.status === "reviewed" ? "Unlock" : "Mark reviewed"} onClick={() => setReviewed(key, cell.status !== "reviewed")}>
                                {cell.status === "reviewed" ? "Locked" : "Lock"}
                              </button>
                              {cell.status !== "reviewed" && (
                                <button className="icon-btn" title="Re-run this cell" disabled={!canEdit || running > 0} onClick={() => fill([{ docName: doc.name, column: col }])}>
                                  Refresh
                                </button>
                              )}
                            </div>
                          </div>
                        )}
                      </td>
                    )
                  })}
                  <td className="review-add-col" aria-hidden="true" />
                </tr>
              ))}
              {documents.length === 0 ? (
                <tr>
                  <td className="review-empty-documents" colSpan={shownColumns.length + 2}>Add documents to this matter to start reviewing.</td>
                </tr>
              ) : shownDocuments.length === 0 && (
                <tr>
                  <td className="review-empty-documents" colSpan={shownColumns.length + 2}>No matching cells.</td>
                </tr>
              )}
            </tbody>
          </table>
        </div>

        {detailShown && detailCell && (
          <dialog
            ref={detailDialog}
            className={`viewer-overlay review-detail-modal${detailClosing ? " is-closing" : ""}`}
            inert={detailClosing}
            aria-label={`Answer and comments: ${detailShown.question}`}
            onCancel={(event) => { event.preventDefault(); event.stopPropagation(); if (viewerOpen) setViewing(undefined); else setDetail(undefined) }}
            {...backdropClose(() => setDetail(undefined))}
            onKeyDown={(event) => {
              // The open document viewer keeps its own focus loop.
              if (event.key !== "Tab" || viewerOpen) return
              const panel = event.currentTarget.querySelector<HTMLElement>(".picker-panel")
              const first = panel?.querySelector<HTMLButtonElement>(".modal-close")
              const last = [...panel?.querySelectorAll<HTMLElement>("button:not(:disabled), textarea") ?? []].at(-1)
              if (event.shiftKey && document.activeElement === first) { event.preventDefault(); last?.focus() }
              else if (!event.shiftKey && document.activeElement === last) { event.preventDefault(); first?.focus() }
            }}
          >
            <div className="picker-panel" onClick={(e) => e.stopPropagation()}>
              <ModalHeader
                title={<span className="viewer-title-text"><span className="review-detail-question">{detailShown.question}</span><span className="viewer-title-meta">{detailShown.docName}</span></span>}
                actions={inWalk && <span className="review-walk-position">{walk.at + 1} of {walk.entries.length}</span>}
                onClose={() => setDetail(undefined)}
              />
              <div className="picker-body review-detail-body">
                {/* In the walk, a flagged cell leads with its marker, in the
                    table's own words; a plain unreviewed cell needs no reason. */}
                {inWalk && walkReason && walkEntry?.marker && (
                  <div className={`review-walk-why is-${MARKER_LOOK[walkEntry.marker.state].tone}`}>
                    <p className="review-walk-word"><MarkerGlyph state={walkEntry.marker.state} />{walkEntry.marker.word}{walkEntry.marker.state === "differs" && <span className="review-ai-tag">AI</span>}</p>
                    <p>{walkReason}</p>
                  </div>
                )}
                {inWalk && !walkReason && <p className="review-walk-reviewed">Reviewed</p>}
                <section className="review-detail-section" aria-labelledby={`${detailIds}-answer`}>
                  <h3 className="review-detail-label" id={`${detailIds}-answer`}>Answer</h3>
                  <p className="review-detail-answer">{detailCell.answer}</p>
                  {detailDiffers && <p className="review-compare-detail"><MarkerGlyph state="differs" /><span>AI comparison: Differs — {detailDiffers}</span></p>}
                  {detailCell.recheck && <p className="review-recheck-detail">{describeRecheck(detailCell.recheck)}</p>}
                </section>
                {detailCell.citation && detailKey && (() => {
                  const citation = detailCell.citation
                  const entry = sourceEntry(detailKey, detailCell)
                  const check = entry?.check
                  const label = `[${citation.documentName} § ${citation.section}]`
                  // The badge sits beside the reference; its sentence below says
                  // what the check found, unless the walk's reason already does.
                  const detail = check ? (walkExplains(check) ? undefined : SOURCE_DETAIL[check])
                    : checkingSources ? "Checking the excerpt against the document…" : "The excerpt could not be checked against the document."
                  return (
                    <section className="review-detail-section" aria-labelledby={`${detailIds}-source`}>
                      <h3 className="review-detail-label" id={`${detailIds}-source`}>Source</h3>
                      <div className={`citation review-source${check && DEGRADED.includes(check) ? ` degraded is-${check}` : ""}`}>
                        <div className="review-source-head">
                          {citationDocument(citation, documents) ? (
                            <button ref={sourceLink} type="button" className="ref" aria-label={`Open ${citation.documentName} at section ${citation.section}`} onClick={() => setViewing(detailKey)}>
                              {label}
                            </button>
                          ) : <div className="ref">{label}<span className="muted"> — Source unavailable in this matter</span></div>}
                          {check && !walkExplains(check) && <span className={`review-source-badge is-${check}`}>{SOURCE_BADGE[check]}</span>}
                        </div>
                        {degradedCitation(citation.documentName)}
                        {citation.structuralPath && <div className="review-source-path">{citation.structuralPath}</div>}
                        <div className="excerpt">{citation.excerpt}</div>
                        {detail && <p className="review-source-check">{detail}</p>}
                      </div>
                    </section>
                  )
                })()}
                {onAsk && (
                  <button type="button" className="linklike review-ask" onClick={() => { if (canLeaveGrid()) allowNavigation(() => onAsk(cellQuestion(detailShown.docName, detailShown.question, detailCell, detailKey ? sourceCheck(detailKey, detailCell) : undefined))) }}>
                    Ask about this answer in chat
                  </button>
                )}
                <section className="review-comments" aria-labelledby={`${detailIds}-comments`}>
                  <h3 className="review-detail-label" id={`${detailIds}-comments`}>Comments</h3>
                  {(detailCell.comments ?? []).map((c) => (
                    <div key={c.id} className="review-comment">
                      <div className="review-comment-meta">
                        <span><strong>{c.author || "Unattributed"}</strong><span className="muted"> · {formatDateTime(c.at)}</span></span>
                        <button className="icon-btn" disabled={!canEdit} aria-label="Delete comment" title="Delete comment" onClick={() => deleteComment(c.id)}>
                          ×
                        </button>
                      </div>
                      <div className="review-comment-text">{c.text}</div>
                    </div>
                  ))}
                  <div className="review-comment-compose">
                    <textarea
                      rows={2}
                      aria-describedby={`${detailIds}-signed`}
                      value={draft}
                      readOnly={commentSaving || !canEdit}
                      aria-label="Add a comment"
                      placeholder="Add a comment..."
                      onChange={(e) => setDraft(e.target.value)}
                      onKeyDown={(e) => {
                        if (e.key === "Enter" && !e.shiftKey) {
                          e.preventDefault()
                          addComment()
                        }
                      }}
                    />
                    <div className="review-comment-actions">
                      <small className="muted" id={`${detailIds}-signed`}>Comments use the attorney name in Settings → Drafting.</small>
                      <button className={!inWalk || composing ? "primary" : undefined} onClick={addComment} disabled={!canEdit || !draft.trim() || commentSaving}>
                        {commentSaving ? "Adding comment…" : "Add comment"}
                      </button>
                    </div>
                  </div>
                </section>
              </div>
              {/* One next step: Mark reviewed while the cell is open, else Next.
                  A reviewed cell is never unlocked or refilled from here. */}
              {inWalk && (() => {
                const open = detailCell.status !== "reviewed"
                return (
                  <div className="review-walk-bar">
                    {composing ? <small className="review-walk-hint">Add or clear the comment to move on.</small> : (
                      <>
                        <button type="button" className="ghost review-walk-previous" disabled={!walk.entries.slice(0, walk.at).some(inTable)} onClick={() => walkTo(-1)}>
                          Previous
                        </button>
                        {open && <button type="button" className="ghost review-walk-skip" onClick={() => walkTo(1)}>Skip</button>}
                      </>
                    )}
                    <button type="button" className={`review-walk-next${composing ? "" : " primary"}`} disabled={composing || (open && !canEdit)} onClick={open ? markReviewedNext : () => walkTo(1)}>
                      {open ? "Mark reviewed · Next" : "Next"}
                    </button>
                  </div>
                )
              })()}
            </div>
            {/* Inside the native detail so it stacks above it in the top layer. */}
            <Exit when={viewerOpen ? detailCell.citation : undefined}>{(citation) => {
              const source = citationDocument(citation, documents)
              return source && (
                <DocumentViewer
                  key={source.name}
                  matterId={matterId}
                  name={source.name}
                  citation={citation}
                  onChanged={() => setSourceRound((n) => n + 1)}
                  onClose={() => {
                    setViewing(undefined)
                    sourceLink.current?.focus({ preventScroll: true })
                  }}
                />
              )
            }}</Exit>
          </dialog>
        )}
      </div>
    </dialog>,
    document.body,
  )
}
