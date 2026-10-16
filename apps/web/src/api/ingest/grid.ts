import { INGEST_URL } from "../../config"
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
    excerpt: string
    charStart: number
    charEnd: number
  }
  status: "filled" | "reviewed"
  questionHash: string
  // Reviewer comment thread on the cell; survives recomputes.
  comments?: GridCellComment[]
}
export type GridCellComment = { id: string; text: string; at: number }
export type Grid = { columns: GridColumn[]; cells: Record<string, GridCellData> }

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
