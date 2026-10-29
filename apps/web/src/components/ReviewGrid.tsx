import { useContext, useEffect, useLayoutEffect, useMemo, useRef, useState } from "react"
import { SplitButton } from "./SplitButton"
import { createPortal } from "react-dom"
import { UNSAFE_DataRouterContext, useBlocker } from "react-router-dom"
import { computeCell, describeRecheck, needsFill } from "../api/review-recheck"
import { cellKey, checkCellSources, getGrid, GridConflictError, hashQuestion, isStale, recheckPassages, saveGrid, type Grid, type GridCellData, type SourceCheck } from "../api/ingest/grid"
import { humanizeError } from "../api/ingest/errors"
import { getPreferences } from "../api/ingest/preferences"
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

// Tabular review: rows are the matter's documents, columns are natural-language
// questions, each cell is an extracted answer with a source citation. The grid is
// the durable diligence work product — columns and cells persist server-side
// (see api/ingest getGrid/saveGrid). Cells compute once and are cached; only
// empty or stale cells recompute, and `reviewed` cells are frozen. A cell left
// without a citation gets exactly one re-check of its own document as part of
// the same computation (api/review-recheck).

const POOL = 5 // concurrent cell extractions; the engine handles one prompt each

// A source check answers for the citation it was asked about, so a cell whose
// citation has since been recomputed shows no verdict until it is checked.
type Citation = NonNullable<GridCellData["citation"]>
const sourceSignature = (c: Citation) => JSON.stringify([c.documentName, c.excerpt, c.charStart, c.charEnd])
const DEGRADED: SourceCheck[] = ["not_found", "missing"]
const SOURCE_BADGE: Record<SourceCheck, string> = { verified: "Verified", reanchored: "Re-anchored", not_found: "Not found", missing: "Source missing" }
const SOURCE_DETAIL: Record<SourceCheck, string> = {
  verified: "The excerpt matches the document at the recorded position.",
  reanchored: "The excerpt is still in the document, but at a different position from the one recorded. The document may have changed since this answer.",
  not_found: "The document no longer contains this excerpt. It has changed since this answer; check the answer against the current text.",
  missing: "This document is no longer in the matter, so the excerpt cannot be checked.",
}

// The chat question for one cell: its question, answer and recorded citation,
// prefilled for the lawyer to edit and send. It carries what the run recorded;
// the assistant does not re-derive the answer from scratch.
function cellQuestion(docName: string, question: string, cell: GridCellData, check?: SourceCheck) {
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

export default function ReviewGrid({ matterId, title, directory, documents, onClose, onAsk }: { matterId: string; title: string; directory: string; documents: Document[]; onClose: () => void; onAsk?: (prompt: string) => void }) {
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
  const [error, setError] = useState<string>()
  const [loaded, setLoaded] = useState(false)
  const [saveProblem, setSaveProblem] = useState<string>()
  const [reloading, setReloading] = useState(false)
  const reloadInFlight = useRef(false)
  const bypassNavigation = useRef(false)
  const dataRouter = useContext(UNSAFE_DataRouterContext)
  const saveState = useRef<{ version: number; pending: boolean; saving: boolean; blocked: boolean; epoch: number; queue: Grid[] }>({ version: 0, pending: false, saving: false, blocked: false, epoch: 0, queue: [] })
  const otherWindow = useReviewGridPresence(directory, matterId)
  const canEdit = loaded && !saveProblem
  const [running, setRunning] = useState(0) // cells currently extracting
  // The detail overlay identifies its cell by key and reads it live from `grid`,
  // so the comment thread updates in place as comments are added or removed.
  const [detail, setDetail] = useState<{ docName: string; columnId: string; question: string }>()
  const [draft, setDraft] = useState("") // comment being composed in the overlay
  const [commentSaving, setCommentSaving] = useState(false)
  // Each dialog closes, and returns focus, as its exit starts.
  const [detailShown, detailClosing] = useExit(detail)
  const detailKey = detailShown ? cellKey(detailShown.docName, detailShown.columnId) : undefined
  const detailCell = detailKey ? grid.cells[detailKey] : undefined
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
  const sources = useMemo(() => Object.entries(grid.cells).flatMap(([key, cell]) => cell.citation
    ? [{ key, signature: sourceSignature(cell.citation), documentName: cell.citation.documentName, excerpt: cell.citation.excerpt, charStart: cell.citation.charStart, charEnd: cell.citation.charEnd }]
    : []), [grid.cells])
  // An edit invalidates the verdict in the same render, before the next effect
  // starts. Late replies remain tied to the document round they actually checked.
  const sourcesKey = JSON.stringify([matterId, sourceRound, sources])
  const sourceCheckFailed = sourceChecks?.key === sourcesKey && sourceChecks.failed
  const sourceEntry = (key: string, cell: GridCellData | undefined) => {
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

  useLayoutEffect(() => {
    if (!detail) return
    const opener = document.activeElement instanceof HTMLElement ? document.activeElement : null
    const modal = detailDialog.current
    modal?.showModal()
    modal?.querySelector<HTMLTextAreaElement>("textarea")?.focus({ preventScroll: true })
    return () => {
      modal?.close()
      if (opener?.isConnected) opener.focus({ preventScroll: true })
    }
  }, [detail])

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
      if (!saveState.current.pending && !saveState.current.saving && !running && !commentSaving && !draft.trim()) return
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
      const snapshot = { ...queued, version: state.version }
      try {
        const saved = await saveGrid(matterId, snapshot)
        if (state.epoch !== epoch) break
        state.version = saved.version ?? state.version + 1
        gridRef.current = { ...gridRef.current, version: state.version }
        setGrid(gridRef.current)
      } catch (e) {
        if (state.epoch !== epoch) break
        state.queue.unshift(queued)
        state.pending = true
        state.blocked = true
        setSaveProblem(e instanceof GridConflictError
          ? "This table conflicts with the saved table. Your changes are still shown here and have not been saved. Extractions in progress may need to be run again after reloading."
          : "The table could not be saved. Your changes are still shown here; the save may or may not have reached the server. Extractions in progress may need to be run again after reloading.")
      }
    }
    state.saving = false
  }

  // One local writer for concurrent cell completions and user changes.
  function commit(next: Grid) {
    if (!loaded || saveState.current.blocked) return
    gridRef.current = next
    setGrid(next)
    saveState.current.pending = true
    // Keep each intentional transition, including an unlock before a later
    // edit, so the server can distinguish it from replacing a reviewed cell.
    saveState.current.queue.push(next)
    void flush()
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
    return !(state.pending || state.saving || running || commentSaving || draft.trim()) || window.confirm("There are unsaved changes or work in progress. Leave the review table anyway?")
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

  function toggleReviewed(key: string) {
    const cell = gridRef.current.cells[key]
    if (!cell) return
    commitCell(key, { ...cell, status: cell.status === "reviewed" ? "filled" : "reviewed" })
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
    return computeCell({ document: docName, question, questionHash: hashQuestion(question) }, {
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

  function fillEmpty() {
    const targets = documents.flatMap((doc) =>
      grid.columns
        .filter((col) => {
          // Compute blank, stale and never re-checked negative cells; never touch a
          // reviewed (locked) cell or one already re-checked for this question.
          return needsFill(grid.cells[cellKey(doc.name, col.id)], hashQuestion(col.question))
        })
        .map((column) => ({ docName: doc.name, column })),
    )
    fill(targets)
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
        <div className="row review-toolbar">
          <div>
            <span className="review-summary">
              {documents.length} document{documents.length === 1 ? "" : "s"} · {grid.columns.length} question{grid.columns.length === 1 ? "" : "s"}
            </span>
          </div>
          <div className="row" style={{ gap: 8 }}>
            {running > 0 && <span className="review-running" role="status">Extracting {running}</span>}
            <button className="primary" onClick={fillEmpty} disabled={!canEdit || documents.length === 0 || grid.columns.length === 0 || running > 0}>
              Fill empty
            </button>
            <SplitButton
              label="Export Excel"
              onClick={() => exportAs("xlsx")}
              toggleLabel="More export formats"
              disabled={documents.length === 0 || grid.columns.length === 0}
              items={[
                { label: "Export CSV", onSelect: () => exportAs("csv") },
                { label: "Export Word", onSelect: () => exportAs("docx") },
              ]}
            />
          </div>
        </div>

        {error && <p role="alert">{error}</p>}
        {saveProblem && <div role="alert"><p>{saveProblem}</p><p>Export the table above or download your local copy before reloading the saved table.</p><button type="button" onClick={downloadPending}>Download unsaved copy</button> <button type="button" disabled={reloading} onClick={reloadGrid}>{reloading ? "Reloading…" : "Reload saved table"}</button></div>}

        {grid.columns.length === 0 && (
          <p className="review-hint">Use + to add a question for every document.</p>
        )}
        <div className="review-scroll" role="region" aria-label="Review table" tabIndex={0}>
          <table className={`review-table${grid.columns.length === 0 ? " is-empty" : ""}`} style={grid.columns.length ? { minWidth: 304 + grid.columns.length * 280 } : undefined} aria-label="Document review">
            <thead>
              <tr>
                <th className="review-doc-col" scope="col">Document</th>
                {grid.columns.map((col) => (
                  <th key={col.id} scope="col">
                    <div className="review-col-head">
                      <button className="review-col-q" disabled={!canEdit} onClick={() => renameColumn(col.id)} aria-label={`Edit question: ${col.question}`} title="Edit question">
                        {col.question}
                      </button>
                      <button className="icon-btn review-col-x" disabled={!canEdit} onClick={() => deleteColumn(col.id)} aria-label={`Delete column: ${col.question}`} title="Delete column">
                        ×
                      </button>
                    </div>
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
              {documents.map((doc) => (
                <tr key={doc.id}>
                  <th className="review-doc-col" scope="row">{doc.name}<QualityNote quality={doc.quality} /></th>
                  {grid.columns.map((col) => {
                    const key = cellKey(doc.name, col.id)
                    const cell = grid.cells[key]
                    const stale = isStale(cell, col.question)
                    const check = sourceCheck(key, cell)
                    const degraded = !!check && DEGRADED.includes(check)
                    return (
                      <td key={col.id} className={`review-cell${cell?.status === "reviewed" ? " reviewed" : ""}${degraded ? " degraded" : ""}`}>
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
                              onClick={() => {
                                setDraft("")
                                setViewing(undefined)
                                setDetail({ docName: doc.name, columnId: col.id, question: col.question })
                              }}
                            >
                              {cell.answer}
                              {cell.citation && <span className="review-cite"> [{cell.citation.documentName} § {cell.citation.section}]</span>}
                              {cell.citation && degradedCitation(cell.citation.documentName)}
                            </button>
                            {degraded && (
                              <span className="review-degraded" title={SOURCE_DETAIL[check]}>
                                {check === "missing" ? "Source missing" : "Source changed"}
                              </span>
                            )}
                            {cell.recheck && (
                              <span className="review-recheck" title={describeRecheck(cell.recheck)}>
                                {cell.recheck.outcome === "found" ? "Found on re-check" : "Not found after re-check"}
                              </span>
                            )}
                            <div className="review-cell-actions">
                              <button
                                className={`icon-btn review-note${cell.comments?.length ? " has-comments" : ""}`}
                                aria-label={`${cell.comments?.length ?? 0} comments on ${doc.name}: ${col.question}`}
                                title={cell.comments?.map((c) => `${c.author || "Unattributed"}: ${c.text}`).join("\n") || "Add a comment"}
                                onClick={() => {
                                  setDraft("")
                                  setViewing(undefined)
                                  setDetail({ docName: doc.name, columnId: col.id, question: col.question })
                                }}
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
                              {stale && <span className="review-stale" title="Question changed since this answer">stale</span>}
                              <button className={`icon-btn review-lock${cell.status === "reviewed" ? " locked" : ""}`} disabled={!canEdit} aria-pressed={cell.status === "reviewed"} title={cell.status === "reviewed" ? "Unlock" : "Mark reviewed"} onClick={() => toggleReviewed(key)}>
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
              {documents.length === 0 && (
                <tr>
                  <td className="review-empty-documents" colSpan={grid.columns.length + 2}>Add documents to this matter to start reviewing.</td>
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
              const first = event.currentTarget.querySelector<HTMLButtonElement>(".modal-close")
              const last = event.currentTarget.querySelector<HTMLButtonElement>(".review-comment-compose > button:not(:disabled)") ?? event.currentTarget.querySelector<HTMLTextAreaElement>("textarea")
              if (event.shiftKey && document.activeElement === first) { event.preventDefault(); last?.focus() }
              else if (!event.shiftKey && document.activeElement === last) { event.preventDefault(); first?.focus() }
            }}
          >
            <div className="picker-panel" onClick={(e) => e.stopPropagation()}>
              <ModalHeader title={detailShown.question} onClose={() => setDetail(undefined)} />
              <div className="picker-body">
                <p className="review-detail-answer">{detailCell.answer}</p>
                {detailCell.recheck && <p className="review-recheck-detail">{describeRecheck(detailCell.recheck)}</p>}
                {detailCell.citation && detailKey && (() => {
                  const citation = detailCell.citation
                  const entry = sourceEntry(detailKey, detailCell)
                  const check = entry?.check
                  const label = `[${citation.documentName} § ${citation.section}]`
                  return (
                    <div className={`citation review-source${check && DEGRADED.includes(check) ? " degraded" : ""}`}>
                      {citationDocument(citation, documents) ? (
                        <button ref={sourceLink} type="button" className="ref" aria-label={`Open ${citation.documentName} at section ${citation.section}`} onClick={() => setViewing(detailKey)}>
                          {label}
                        </button>
                      ) : <div className="ref">{label}<span className="muted"> — Source unavailable in this matter</span></div>}
                      {degradedCitation(citation.documentName)}
                      {citation.structuralPath && <div className="review-source-path">{citation.structuralPath}</div>}
                      <div className="excerpt">{citation.excerpt}</div>
                      <p className="review-source-check">
                        {check
                          ? <><span className={`review-source-badge is-${check}`}>{SOURCE_BADGE[check]}</span> {SOURCE_DETAIL[check]}</>
                          : sourceCheckFailed || entry ? "The excerpt could not be checked against the document." : "Checking the excerpt against the document…"}
                      </p>
                    </div>
                  )
                })()}
                {onAsk && (
                  <div className="review-ask">
                    <button type="button" onClick={() => { if (canLeaveGrid()) allowNavigation(() => onAsk(cellQuestion(detailShown.docName, detailShown.question, detailCell, detailKey ? sourceCheck(detailKey, detailCell) : undefined))) }}>
                      Ask about this answer in chat
                    </button>
                  </div>
                )}
                <div className="review-comments">
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
                  <small className="muted">Comments use the attorney name in Settings → Drafting.</small>
                  <div className="review-comment-compose">
                    <textarea
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
                    <button className="primary" onClick={addComment} disabled={!canEdit || !draft.trim() || commentSaving}>
                      {commentSaving ? "Adding comment…" : "Add comment"}
                    </button>
                  </div>
                </div>
              </div>
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
