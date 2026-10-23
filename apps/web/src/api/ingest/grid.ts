import { INGEST_URL } from "../../config"
import type { Citation } from "../opencode/parts"
import { ingestError } from "./errors"

// Tabular-review grid. Columns are questions; cells are keyed `<docName>::<colId>`.
// Rows are the matter's documents, so they are not stored here.
export type GridColumn = { id: string; question: string }
export type GridCellData = {
  answer: string
  citation?: {
    documentName: string
    docPath: string
    section: string
    // The clause hierarchy the search hit carried, when the index had one.
    structuralPath?: string
    excerpt: string
    charStart: number
    charEnd: number
  }
  status: "filled" | "reviewed"
  questionHash: string
  // Reviewer comment thread on the cell; survives recomputes.
  comments?: GridCellComment[]
  // The one re-check a cell computed without a citation receives, and what it
  // searched (mirrors services/ingest/src/grid.ts).
  recheck?: GridRecheck
}
export type GridCellComment = { id: string; text: string; at: number }
export type Grid = { columns: GridColumn[]; cells: Record<string, GridCellData> }

export function cellKey(docName: string, columnId: string) {
  return `${docName}::${columnId}`
}

// Stable id for a question column. Editing a question changes its hash, not its
// id, so existing cells stay linked but flag themselves stale.
export function hashQuestion(q: string) {
  let h = 5381
  for (let i = 0; i < q.length; i++) h = (h * 33) ^ q.charCodeAt(i)
  return (h >>> 0).toString(36)
}

/** An answer computed for an earlier wording of its question. A reviewed
 * (locked) answer is the lawyer's and is never stale. */
export function isStale(cell: GridCellData | undefined, question: string) {
  return !!cell && cell.status !== "reviewed" && cell.questionHash !== hashQuestion(question)
}
export type RecheckScope = { passages: number; read: number; candidates: number; reread: number; terms: string[]; channels: string[] }
export type GridRecheck = { outcome: "found" | "not_found" } & RecheckScope
export type RecheckPassages = { citations: Citation[]; passages: string; scope: RecheckScope }
export type SeenPassage = { docPath: string; charStart: number; charEnd: number }

export async function getGrid(id: string): Promise<Grid> {
  const res = await fetch(`${INGEST_URL}/matters/${id}/grid`)
  if (!res.ok) throw new Error(await ingestError(res, "Could not load the review grid"))
  return res.json()
}

export async function saveGrid(id: string, grid: Grid): Promise<void> {
  const res = await fetch(`${INGEST_URL}/matters/${id}/grid`, {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(grid),
  })
  if (!res.ok) throw new Error(await ingestError(res, "Could not save the review grid"))
}

// The deterministic re-check pass for one negative cell: passages of that one
// document holding the question's terms, skipping the passages the first
// extraction saw or, with `includeSeen`, leading with them.
export async function recheckPassages(
  id: string,
  request: { document: string; question: string; seen: SeenPassage[]; includeSeen: boolean },
): Promise<RecheckPassages> {
  const res = await fetch(`${INGEST_URL}/matters/${id}/grid/recheck`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(request),
  })
  if (!res.ok) throw new Error(await ingestError(res, "Could not re-check the review cell"))
  return res.json()
}

// Whether each cited cell's excerpt is still the document's text at its offsets
// (mirrors services/ingest/src/grid.ts SourceCheck). A cell left out of the
// result could not be read and shows no verdict.
export type SourceCheck = "verified" | "reanchored" | "not_found" | "missing"
export type SourceCheckCell = { key: string; documentName: string; excerpt: string; charStart: number; charEnd: number }

export async function checkCellSources(id: string, cells: SourceCheckCell[]): Promise<Record<string, SourceCheck>> {
  const res = await fetch(`${INGEST_URL}/matters/${id}/grid/sources`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ cells }),
  })
  if (!res.ok) throw new Error(await ingestError(res, "Could not check the review sources"))
  return res.json()
}
