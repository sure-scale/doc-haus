import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs"
import path from "node:path"

// The tabular-review grid for a matter. Persisted as a single JSON file next to
// the matter's retrieval index (`<matter>/.dochaus/grid.json`) so the review
// work product lives with the documents it summarizes and survives across
// devices and sessions. The web app owns the merge; this service just persists.
//
// Rows are NOT stored here — they are the matter's documents, enumerated live
// from the index. We store the columns (the questions) and the computed cells.

export type GridColumn = { id: string; question: string }

export type GridCell = {
  answer: string
  citation?: {
    documentName: string
    docPath: string
    section: string
    excerpt: string
    charStart: number
    charEnd: number
  }
  // `reviewed` cells are locked: a lawyer has signed off, so they are never
  // recomputed. `questionHash` is the hash of the column's question at compute
  // time, so editing a column marks its cells stale without changing them.
  status: "filled" | "reviewed"
  questionHash: string
}

// Cells are keyed `<documentName>::<columnId>`.
export type Grid = { columns: GridColumn[]; cells: Record<string, GridCell> }

function gridFile(matterDir: string) {
  return path.join(matterDir, ".dochaus", "grid.json")
}

export function readGrid(matterDir: string): Grid {
  const file = gridFile(matterDir)
  if (!existsSync(file)) return { columns: [], cells: {} }
  return JSON.parse(readFileSync(file, "utf8")) as Grid
}

export function writeGrid(matterDir: string, grid: Grid) {
  mkdirSync(path.join(matterDir, ".dochaus"), { recursive: true })
  writeFileSync(gridFile(matterDir), JSON.stringify(grid, null, 2))
}
