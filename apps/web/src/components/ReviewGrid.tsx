import { useEffect, useMemo, useRef, useState } from "react"
import type { Part } from "@opencode-ai/sdk"
import { getGrid, saveGrid, type Document, type Grid, type GridCellData } from "../api/ingest"
import { createSession, getMessages, matterClient, sendPrompt, type Citation, type Client } from "../api/opencode"

// Tabular review: rows are the matter's documents, columns are natural-language
// questions, each cell is an extracted answer with a source citation. The grid is
// the durable diligence work product — columns and cells persist server-side
// (see api/ingest getGrid/saveGrid). Cells compute once and are cached; only
// empty or stale cells recompute, and `reviewed` cells are frozen.

const POOL = 5 // concurrent cell extractions; the engine handles one prompt each

// Stable id for a question column. Editing a question changes its hash, not its
// id, so existing cells stay linked but flag themselves stale.
function hashQuestion(q: string) {
  let h = 5381
  for (let i = 0; i < q.length; i++) h = (h * 33) ^ q.charCodeAt(i)
  return (h >>> 0).toString(36)
}

function cellKey(docName: string, columnId: string) {
  return `${docName}::${columnId}`
}

// Reduce one extraction session's assistant parts to its answer text and the
// source citation. The agent ends with a `Source: §<section>` line naming the
// passage it used; we match that to the search-document result so the cited
// clause is the one the answer came from, not merely the top-ranked hit. A real
// answer without a declared source gets no citation rather than a wrong one.
function readAnswer(parts: Part[]) {
  const raw = parts
    .filter((p): p is Extract<Part, { type: "text" }> => p.type === "text")
    .map((p) => p.text)
    .join("")
    .trim()
  const citations = parts
    .filter((p): p is Extract<Part, { type: "tool" }> => p.type === "tool")
    .filter((p) => p.tool === "search-document" && p.state.status === "completed")
    .flatMap((p) => ((p.state as { metadata?: { citations?: Citation[] } }).metadata?.citations ?? []))
  const source = raw.match(/\n\s*Source:\s*§?\s*(.+?)\s*$/i)
  const text = source ? raw.slice(0, source.index).trim() : raw
  const section = source?.[1]?.trim()
  const citation = section ? (citations.find((c) => c.section === section) ?? citations[0]) : undefined
  return { text, citation }
}

export default function ReviewGrid({ matterId, directory, documents }: { matterId: string; directory: string; documents: Document[] }) {
  const client = useMemo<Client>(() => matterClient(directory), [directory])
  const [grid, setGrid] = useState<Grid>({ columns: [], cells: {} })
  const gridRef = useRef<Grid>(grid)
  const [running, setRunning] = useState(0) // cells currently extracting
  const [detail, setDetail] = useState<{ docName: string; question: string; cell: GridCellData }>()

  useEffect(() => {
    getGrid(matterId).then((g) => commit(g))
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [matterId])

  // One writer for both state and the server, so concurrent cell completions
  // merge against the latest grid (each call reads+writes gridRef synchronously).
  function commit(next: Grid) {
    gridRef.current = next
    setGrid(next)
    saveGrid(matterId, next)
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
  }

  function toggleReviewed(key: string) {
    const cell = gridRef.current.cells[key]
    if (!cell) return
    commitCell(key, { ...cell, status: cell.status === "reviewed" ? "filled" : "reviewed" })
  }

  // Extract one cell: a scoped, throwaway session titled so it stays out of the
  // sidebar's conversation list (which filters "doc.haus"-titled sessions).
  async function extract(docName: string, question: string): Promise<GridCellData> {
    const session = await createSession(client, `doc.haus grid: ${docName}`)
    await sendPrompt(client, session.id, "extract", `Document: "${docName}". Question: ${question}`)
    const messages = await getMessages(client, session.id)
    // A tool-using turn emits several assistant messages (one per model step): the
    // search-document call lands on an earlier one, the final answer text on the
    // last. Read every assistant part so the citation is not dropped.
    const { text, citation } = readAnswer(messages.filter((m) => m.info.role === "assistant").flatMap((m) => m.parts))
    return {
      answer: text || "Not addressed",
      citation: citation && {
        documentName: citation.documentName,
        docPath: citation.docPath,
        section: citation.section,
        excerpt: citation.excerpt,
        charStart: citation.charStart,
        charEnd: citation.charEnd,
      },
      status: "filled",
      questionHash: hashQuestion(question),
    }
  }

  // Run a set of (document, column) targets through a small concurrency pool so a
  // large grid does not open hundreds of sessions at once.
  async function fill(targets: Array<{ docName: string; column: { id: string; question: string } }>) {
    if (targets.length === 0) return
    setRunning((n) => n + targets.length)
    let next = 0
    const worker = async () => {
      while (next < targets.length) {
        const t = targets[next++]
        const data = await extract(t.docName, t.column.question)
        commitCell(cellKey(t.docName, t.column.id), data)
        setRunning((n) => n - 1)
      }
    }
    await Promise.all(Array.from({ length: Math.min(POOL, targets.length) }, worker))
  }

  function fillEmpty() {
    const targets = documents.flatMap((doc) =>
      grid.columns
        .filter((col) => {
          const cell = grid.cells[cellKey(doc.name, col.id)]
          // Compute blanks and stale cells; never touch a reviewed (locked) cell.
          return !cell || (cell.status !== "reviewed" && cell.questionHash !== hashQuestion(col.question))
        })
        .map((column) => ({ docName: doc.name, column })),
    )
    fill(targets)
  }

  function exportCsv() {
    const esc = (v: string) => `"${v.replace(/"/g, '""')}"`
    const header = ["Document", ...grid.columns.map((c) => c.question)].map(esc).join(",")
    const rows = documents.map((doc) =>
      [doc.name, ...grid.columns.map((c) => grid.cells[cellKey(doc.name, c.id)]?.answer ?? "")].map(esc).join(","),
    )
    const url = URL.createObjectURL(new Blob([[header, ...rows].join("\n")], { type: "text/csv" }))
    const a = document.createElement("a")
    a.href = url
    a.download = "tabular-review.csv"
    a.click()
    URL.revokeObjectURL(url)
  }

  if (documents.length === 0)
    return <p className="muted">Add documents to this matter to build a review grid.</p>

  return (
    <div className="card review">
      <div className="row review-toolbar">
        <h2 style={{ margin: 0 }}>Tabular review</h2>
        <div className="row" style={{ gap: 8 }}>
          {running > 0 && <span className="muted">Extracting {running}...</span>}
          <button onClick={addColumn}>Add column</button>
          <button onClick={fillEmpty} disabled={grid.columns.length === 0 || running > 0}>
            Fill empty
          </button>
          <button onClick={exportCsv} disabled={grid.columns.length === 0}>
            Export CSV
          </button>
        </div>
      </div>

      {grid.columns.length === 0 ? (
        <p className="muted">Add a column — a question asked of every document, e.g. "What is the governing law?"</p>
      ) : (
        <div className="review-scroll">
          <table className="review-table">
            <thead>
              <tr>
                <th className="review-doc-col">Document</th>
                {grid.columns.map((col) => (
                  <th key={col.id}>
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
              </tr>
            </thead>
            <tbody>
              {documents.map((doc) => (
                <tr key={doc.id}>
                  <td className="review-doc-col">{doc.name}</td>
                  {grid.columns.map((col) => {
                    const key = cellKey(doc.name, col.id)
                    const cell = grid.cells[key]
                    const stale = cell && cell.status !== "reviewed" && cell.questionHash !== hashQuestion(col.question)
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
                            <button className="review-answer" onClick={() => setDetail({ docName: doc.name, question: col.question, cell })}>
                              {cell.answer}
                              {cell.citation && <span className="review-cite"> [{cell.citation.documentName} § {cell.citation.section}]</span>}
                            </button>
                            <div className="review-cell-actions">
                              {stale && <span className="review-stale" title="Question changed since this answer">stale</span>}
                              <button className="icon-btn" title={cell.status === "reviewed" ? "Unlock" : "Mark reviewed"} onClick={() => toggleReviewed(key)}>
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
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {detail && (
        <div className="viewer-overlay" onClick={() => setDetail(undefined)}>
          <div className="picker-panel" onClick={(e) => e.stopPropagation()}>
            <div className="viewer-bar">
              <span className="viewer-title">{detail.question}</span>
              <button onClick={() => setDetail(undefined)}>Close</button>
            </div>
            <div className="picker-body">
              <p className="review-detail-answer">{detail.cell.answer}</p>
              {detail.cell.citation && (
                <div className="citation">
                  <div className="ref">[{detail.cell.citation.documentName} § {detail.cell.citation.section}]</div>
                  <div className="excerpt">{detail.cell.citation.excerpt}</div>
                </div>
              )}
            </div>
          </div>
        </div>
      )}
    </div>
  )
}
