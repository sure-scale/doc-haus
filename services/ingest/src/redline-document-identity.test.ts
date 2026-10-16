import { expect, test } from "bun:test"
import { mkdtempSync, mkdirSync, readFileSync, realpathSync, rmSync, symlinkSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { recordRedlineProposal } from "../../../dochaus/lib/redlines"
import { listPendingRedlines, openDb, pendingRedlineCounts } from "./db"
import { documentRegistryIdentity } from "./document-identity"

function proposal(docPath: string, anchorId: string) {
  return { docPath, docName: "source.docx", scope: "phrase" as const, findText: "Original", oldText: "Original", newText: "Proposed", author: "Reviewer", reviewComment: "Clarify the wording.", anchorId }
}

test("legacy alias proposals remain visible and future writes use canonical paths without applying edits", () => {
  const root = mkdtempSync(path.join(tmpdir(), "redline-identity-"))
  const matter = path.join(root, "matter")
  const alias = path.join(root, "alias")
  mkdirSync(matter)
  symlinkSync(matter, alias, "dir")
  const file = path.join(matter, "source.docx")
  writeFileSync(file, "Original source bytes")
  const db = openDb(matter)
  try {
    const first = recordRedlineProposal(alias, proposal(path.join(alias, "source.docx"), "first"))
    expect((db.query("SELECT doc_path FROM redlines WHERE id = ?").get(first.id) as { doc_path: string }).doc_path).toBe(realpathSync.native(file))
    // A row created by the old transport is kept verbatim; no data rewrite is needed.
    db.run("UPDATE redlines SET doc_path = ? WHERE id = ?", [path.join(alias, "source.docx"), first.id])
    const second = recordRedlineProposal(matter, proposal(file, "second"))
    expect(listPendingRedlines(db, realpathSync.native(file)).map(row => row.id)).toEqual([first.id, second.id])
    expect(pendingRedlineCounts(db)[documentRegistryIdentity(file)]).toBe(2)
    expect((db.query("SELECT doc_path FROM redlines WHERE id = ?").get(first.id) as { doc_path: string }).doc_path).toBe(path.join(alias, "source.docx"))
    expect(readFileSync(file, "utf8")).toBe("Original source bytes")
    expect(listPendingRedlines(db, file).every(row => row.status === "pending")).toBe(true)
  } finally {
    db.close()
    rmSync(root, { recursive: true, force: true })
  }
})

test("a symlink escaping the matter cannot receive a redline proposal", () => {
  const root = mkdtempSync(path.join(tmpdir(), "redline-escape-"))
  const matter = path.join(root, "matter")
  mkdirSync(matter)
  const outside = path.join(root, "outside.docx")
  writeFileSync(outside, "Protected source")
  symlinkSync(outside, path.join(matter, "source.docx"))
  try {
    expect(() => recordRedlineProposal(matter, proposal(path.join(matter, "source.docx"), "outside"))).toThrow("inside its matter")
    expect(readFileSync(outside, "utf8")).toBe("Protected source")
  } finally {
    rmSync(root, { recursive: true, force: true })
  }
})
