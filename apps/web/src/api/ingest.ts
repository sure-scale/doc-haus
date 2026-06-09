import { INGEST_URL } from "../config"

// Matters and documents live in the ingest service: it owns matter directories
// under WORKSPACE_ROOT and turns uploaded DOCX into the per-matter embedding DB.
// OpenCode has no upload endpoint, so all document I/O goes through here.

export type Matter = { id: string; title: string; reference?: string; dir: string; created_at: number }
export type Document = { id: number; name: string; doc_path: string; created_at: number }
export type MatterDetail = Matter & { documents: Document[] }
export type IngestResult = { name: string; docPath: string; sections: number; chunks: number }

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
}
export type Grid = { columns: GridColumn[]; cells: Record<string, GridCellData> }

export async function listMatters(): Promise<Matter[]> {
  const res = await fetch(`${INGEST_URL}/matters`)
  return res.json()
}

export async function createMatter(title: string, reference?: string): Promise<Matter> {
  const res = await fetch(`${INGEST_URL}/matters`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ title, reference }),
  })
  return res.json()
}

export async function deleteMatter(id: string): Promise<void> {
  await fetch(`${INGEST_URL}/matters/${id}`, { method: "DELETE" })
}

export async function getMatter(id: string): Promise<MatterDetail> {
  const res = await fetch(`${INGEST_URL}/matters/${id}`)
  return res.json()
}

export async function uploadDocument(id: string, file: File): Promise<IngestResult> {
  const form = new FormData()
  form.append("file", file)
  const res = await fetch(`${INGEST_URL}/matters/${id}/documents`, { method: "POST", body: form })
  return res.json()
}

export async function deleteDocument(id: string, name: string): Promise<void> {
  await fetch(`${INGEST_URL}/matters/${id}/documents?name=${encodeURIComponent(name)}`, { method: "DELETE" })
}

export async function getGrid(id: string): Promise<Grid> {
  const res = await fetch(`${INGEST_URL}/matters/${id}/grid`)
  return res.json()
}

export async function saveGrid(id: string, grid: Grid): Promise<void> {
  await fetch(`${INGEST_URL}/matters/${id}/grid`, {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(grid),
  })
}

// Raw .docx bytes for a matter document. The viewer renders these client-side
// via WASM, so the file is fetched from our own ingest service and converted in
// the browser — it never leaves for any third-party service.
export async function fetchDocumentBytes(id: string, name: string): Promise<Uint8Array> {
  const res = await fetch(`${INGEST_URL}/matters/${id}/documents/content?name=${encodeURIComponent(name)}`)
  if (!res.ok) throw new Error(`Could not load ${name} (${res.status})`)
  return new Uint8Array(await res.arrayBuffer())
}
