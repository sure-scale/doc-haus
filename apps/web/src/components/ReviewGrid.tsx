import { useEffect, useLayoutEffect, useMemo, useRef, useState } from "react"
import { createPortal } from "react-dom"
import { computeCell, describeRecheck, needsFill } from "../api/review-recheck"
import { cellKey, getGrid, hashQuestion, isStale, recheckPassages, saveGrid, type Grid, type GridCellData } from "../api/ingest/grid"
import { humanizeError } from "../api/ingest/errors"
import type { Document } from "../api/ingest/documents"
import { createSession, getMessages, sendPrompt } from "../api/opencode/sessions"
import { matterClient, type Client } from "../api/opencode/client"
import ModalHeader from "./ModalHeader"
import { reviewCsv, reviewDocx, reviewExport, reviewXlsx } from "./review-export"
import { formatDateTime } from "../lib/format"

// Tabular review: rows are the matter's documents, columns are natural-language
// questions, each cell is an extracted answer with a source citation. The grid is
// the durable diligence work product — columns and cells persist server-side
// (see api/ingest getGrid/saveGrid). Cells compute once and are cached; only
// empty or stale cells recompute, and `reviewed` cells are frozen. A cell left
// without a citation gets exactly one re-check of its own document as part of
// the same computation (api/review-recheck).

const POOL = 5 // concurrent cell extractions; the engine handles one prompt each

export default function ReviewGrid({ matterId, title, directory, documents, onClose }: { matterId: string; title: string; directory: string; documents: Document[]; onClose: () => void }) {
  const dialog = useRef<HTMLDialogElement>(null)
  const client = useMemo<Client>(() => matterClient(directory), [directory])
  const [grid, setGrid] = useState<Grid>({ columns: [], cells: {} })
  const gridRef = useRef<Grid>(grid)
  const [error, setError] = useState<string>()
  const [running, setRunning] = useState(0) // cells currently extracting
  // The detail overlay identifies its cell by key and reads it live from `grid`,
  // so the comment thread updates in place as comments are added or removed.
  const [detail, setDetail] = useState<{ docName: string; columnId: string; question: string }>()
  const [draft, setDraft] = useState("") // comment being composed in the overlay
  const detailCell = detail ? grid.cells[cellKey(detail.docName, detail.columnId)] : undefined

  useLayoutEffect(() => {
    const modal = dialog.current
    modal?.showModal()
    return () => modal?.close()
  }, [])

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
    getGrid(matterId).then((g) => commit(g)).catch((e) => setError(humanizeError(e)))
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [matterId])

  // One writer for both state and the server, so concurrent cell completions
  // merge against the latest grid (each call reads+writes gridRef synchronously).
  function commit(next: Grid) {
    gridRef.current = next
    setGrid(next)
    saveGrid(matterId, next).catch((e) => setError(humanizeError(e)))
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
    commit({ columns: gridRef.current.columns.filter((c) => c.id !== id), cells })
    dialog.current?.querySelector<HTMLButtonElement>(".review-add-button")?.focus()
  }

  function toggleReviewed(key: string) {
    const cell = gridRef.current.cells[key]
    if (!cell) return
    commitCell(key, { ...cell, status: cell.status === "reviewed" ? "filled" : "reviewed" })
  }

  function addComment() {
    const text = draft.trim()
    if (!detail || !text) return
    const key = cellKey(detail.docName, detail.columnId)
    const cell = gridRef.current.cells[key]
    commitCell(key, {
      ...cell,
      comments: [...(cell.comments ?? []), { id: crypto.randomUUID().slice(0, 6), text, at: Date.now() }],
    })
    setDraft("")
  }

  function deleteComment(id: string) {
    if (!detail) return
    const key = cellKey(detail.docName, detail.columnId)
    const cell = gridRef.current.cells[key]
    commitCell(key, { ...cell, comments: (cell.comments ?? []).filter((c) => c.id !== id) })
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
    if (targets.length === 0) return
    setError(undefined)
    setRunning((n) => n + targets.length)
    let next = 0
    const worker = async () => {
      while (next < targets.length) {
        const t = targets[next++]
        try {
          const data = await extract(t.docName, t.column.question)
          const key = cellKey(t.docName, t.column.id)
          commitCell(key, { ...data, comments: gridRef.current.cells[key]?.comments })
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
      className="viewer-panel review-modal"
      aria-label={`Review table — ${title}`}
      onCancel={(event) => { event.preventDefault(); detail ? setDetail(undefined) : onClose() }}
      onClick={(event) => {
        if (event.target !== event.currentTarget) return
        const bounds = event.currentTarget.getBoundingClientRect()
        if (event.clientX < bounds.left || event.clientX > bounds.right || event.clientY < bounds.top || event.clientY > bounds.bottom) onClose()
      }}
    >
      <ModalHeader title={`Review table — ${title}`} onClose={() => detail ? setDetail(undefined) : onClose()} />
      <div className="card review">
        <div className="row review-toolbar">
          <div>
            <span className="review-summary">
              {documents.length} document{documents.length === 1 ? "" : "s"} · {grid.columns.length} question{grid.columns.length === 1 ? "" : "s"}
            </span>
          </div>
          <div className="row" style={{ gap: 8 }}>
            {running > 0 && <span className="review-running">Extracting {running}</span>}
            <button onClick={fillEmpty} disabled={documents.length === 0 || grid.columns.length === 0 || running > 0}>
              Fill empty
            </button>
            <button onClick={() => exportAs("xlsx")} disabled={documents.length === 0 || grid.columns.length === 0}>
              Export Excel
            </button>
            <button onClick={() => exportAs("csv")} disabled={documents.length === 0 || grid.columns.length === 0}>
              Export CSV
            </button>
            <button onClick={() => exportAs("docx")} disabled={documents.length === 0 || grid.columns.length === 0}>
              Export Word
            </button>
          </div>
        </div>

        {error && <p role="alert">{error}</p>}

        {grid.columns.length === 0 && (
          <p className="review-hint">Use + to add a question for every document.</p>
        )}
        <div className="review-scroll" role="region" aria-label="Review table" tabIndex={0}>
          <table className={`review-table${grid.columns.length === 0 ? " is-empty" : ""}`} aria-label="Document review">
            <thead>
              <tr>
                <th className="review-doc-col" scope="col">Document</th>
                {grid.columns.map((col) => (
                  <th key={col.id} scope="col">
                    <div className="review-col-head">
                      <button className="review-col-q" onClick={() => renameColumn(col.id)} title="Edit question">
                        {col.question}
                      </button>
                      <button className="icon-btn review-col-x" onClick={() => deleteColumn(col.id)} title="Delete column">
                        ×
                      </button>
                    </div>
                  </th>
                ))}
                <th className="review-add-col" scope="col">
                  <button className="review-add-button" type="button" onClick={addColumn} aria-label="Add column" title="Add column">
                    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" aria-hidden="true">
                      <path d="M12 5v14M5 12h14" />
                    </svg>
                  </button>
                </th>
              </tr>
            </thead>
            <tbody>
              {documents.map((doc) => (
                <tr key={doc.id}>
                  <th className="review-doc-col" scope="row">{doc.name}</th>
                  {grid.columns.map((col) => {
                    const key = cellKey(doc.name, col.id)
                    const cell = grid.cells[key]
                    const stale = isStale(cell, col.question)
                    return (
                      <td key={col.id} className={`review-cell${cell?.status === "reviewed" ? " reviewed" : ""}`}>
                        {!cell ? (
                          <button
                            className="review-cell-fill"
                            disabled={running > 0}
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
                                setDetail({ docName: doc.name, columnId: col.id, question: col.question })
                              }}
                            >
                              {cell.answer}
                              {cell.citation && <span className="review-cite"> [{cell.citation.documentName} § {cell.citation.section}]</span>}
                            </button>
                            {cell.recheck && (
                              <span className="review-recheck" title={describeRecheck(cell.recheck)}>
                                {cell.recheck.outcome === "found" ? "Found on re-check" : "Not found after re-check"}
                              </span>
                            )}
                            <div className="review-cell-actions">
                              <button
                                className={`icon-btn review-note${cell.comments?.length ? " has-comments" : ""}`}
                                title={cell.comments?.map((c) => c.text).join("\n") || "Add a comment"}
                                onClick={() => {
                                  setDraft("")
                                  setDetail({ docName: doc.name, columnId: col.id, question: col.question })
                                }}
                              >
                                <svg width="12" height="12" viewBox="0 0 16 16" aria-hidden="true">
                                  <path
                                    d="M3 2.5h10A1.5 1.5 0 0 1 14.5 4v6a1.5 1.5 0 0 1-1.5 1.5H8.5L5 14.5v-3H3A1.5 1.5 0 0 1 1.5 10V4A1.5 1.5 0 0 1 3 2.5Z"
                                    fill="none"
                                    stroke="currentColor"
                                    strokeWidth="1.4"
                                    strokeLinejoin="round"
                                  />
                                </svg>
                                {cell.comments?.length ?? 0}
                              </button>
                              {stale && <span className="review-stale" title="Question changed since this answer">stale</span>}
                              <button className={`icon-btn review-lock${cell.status === "reviewed" ? " locked" : ""}`} title={cell.status === "reviewed" ? "Unlock" : "Mark reviewed"} onClick={() => toggleReviewed(key)}>
                                {cell.status === "reviewed" ? "Locked" : "Lock"}
                              </button>
                              {cell.status !== "reviewed" && (
                                <button className="icon-btn" title="Re-run this cell" disabled={running > 0} onClick={() => fill([{ docName: doc.name, column: col }])}>
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

        {detail && detailCell && (
          <div className="viewer-overlay" onClick={() => setDetail(undefined)}>
            <div className="picker-panel" onClick={(e) => e.stopPropagation()}>
              <ModalHeader title={detail.question} onClose={() => setDetail(undefined)} />
              <div className="picker-body">
                <p className="review-detail-answer">{detailCell.answer}</p>
                {detailCell.recheck && <p className="review-recheck-detail">{describeRecheck(detailCell.recheck)}</p>}
                {detailCell.citation && (
                  <div className="citation">
                    <div className="ref">[{detailCell.citation.documentName} § {detailCell.citation.section}]</div>
                    <div className="excerpt">{detailCell.citation.excerpt}</div>
                  </div>
                )}
                <div className="review-comments">
                  {(detailCell.comments ?? []).map((c) => (
                    <div key={c.id} className="review-comment">
                      <div className="review-comment-meta">
                        <span className="muted">{formatDateTime(c.at)}</span>
                        <button className="icon-btn" title="Delete comment" onClick={() => deleteComment(c.id)}>
                          ×
                        </button>
                      </div>
                      <div className="review-comment-text">{c.text}</div>
                    </div>
                  ))}
                  <div className="review-comment-compose">
                    <textarea
                      value={draft}
                      placeholder="Add a comment..."
                      onChange={(e) => setDraft(e.target.value)}
                      onKeyDown={(e) => {
                        if (e.key === "Enter" && !e.shiftKey) {
                          e.preventDefault()
                          addComment()
                        }
                      }}
                    />
                    <button onClick={addComment} disabled={!draft.trim()}>
                      Send
                    </button>
                  </div>
                </div>
              </div>
            </div>
          </div>
        )}
      </div>
    </dialog>,
    document.body,
  )
}
