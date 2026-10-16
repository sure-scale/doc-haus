import { afterEach, expect, test } from "bun:test"
import { Database } from "bun:sqlite"
import { mkdtempSync, mkdirSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { openDb, listDocuments, upsertDeliverable, refreshDeliverable } from "./db"
import { documentLabelRoutes } from "./document-labels"

const folders: string[] = []
afterEach(() => { folders.splice(0).forEach((folder) => rmSync(folder, { recursive: true, force: true })) })

test("labels migrate, persist, and survive delivery updates without changing document roles", async () => {
  const folder = mkdtempSync(path.join(tmpdir(), "document-labels-"))
  folders.push(folder)
  mkdirSync(path.join(folder, ".dochaus"))
  const legacy = new Database(path.join(folder, ".dochaus/legal.db"))
  legacy.run("CREATE TABLE documents (id INTEGER PRIMARY KEY, doc_path TEXT UNIQUE, name TEXT, created_at INTEGER)")
  legacy.run("INSERT INTO documents VALUES (1, ?, 'source.docx', 1)", [path.join(folder, "source.docx")])
  legacy.close()
  let db = openDb(folder)
  const file = path.join(folder, "draft.docx")
  upsertDeliverable(db, file, "draft.docx", 2)
  db.close()
  const app = documentLabelRoutes(() => folder)
  const patch = (body: unknown) => app.request("/matters/example/documents/label", {
    method: "PATCH", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body),
  })
  expect((await patch({ name: "draft.docx", label: "final" })).status).toBe(200)
  expect((await patch({ name: "draft.docx", label: "approved" })).status).toBe(400)
  expect((await patch({ name: "missing.docx", label: "draft" })).status).toBe(404)
  db = openDb(folder)
  expect(listDocuments(db)).toContainEqual(expect.objectContaining({ name: "draft.docx", label: "final", role: "deliverable" }))
  expect(listDocuments(db)).toContainEqual(expect.objectContaining({ name: "source.docx", label: null, role: "source" }))
  refreshDeliverable(db, file, "draft.docx", 3)
  expect(listDocuments(db)).toContainEqual(expect.objectContaining({ name: "draft.docx", label: "final", role: "deliverable" }))
  upsertDeliverable(db, file, "draft.docx", 4)
  expect(listDocuments(db)).toContainEqual(expect.objectContaining({ name: "draft.docx", label: "final", role: "deliverable" }))
  db.close()
})
