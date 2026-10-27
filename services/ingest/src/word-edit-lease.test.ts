import { expect, test } from "bun:test"
import { mkdtempSync, readFileSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { Document, Packer, Paragraph } from "docx"
import { assertMatterWordEditsAllowed, assertWordEditAllowed, claimWordEditLease, clearWordEditLease, EditingLeaseError, releaseWordEditLease, renewWordEditLease, wordLeaseState } from "../../../dochaus/lib/word-edit-lease"
import { openDb } from "./db"
import { deleteMatterDocument } from "./document-delete"
import { readCurrentVersion } from "./document-versions"
import { DeliverableConflictError, writeDeliverable, writeDeliverableRevision } from "./deliverable"
import { ingestDocument } from "./ingest"
import { withArtifactMutationLocksAsync } from "../../../dochaus/capabilities/artifacts/mutation-lock"

const bytes = () => Packer.toBuffer(new Document({ sections: [{ children: [new Paragraph("Synthetic lease document.")] }] }))

test("authoritative writers reject an active Word lease before mutating; release and identity fences remain recoverable", async () => {
  const dir = mkdtempSync(path.join(tmpdir(), "dochaus-word-lease-"))
  try {
    await writeDeliverable(dir, "Draft.docx", await bytes())
    const file = path.join(dir, "Draft.docx")
    const before = readFileSync(file)
    const current = readCurrentVersion(dir, "Draft.docx")!
    const db = openDb(dir)
    const grant = "a".repeat(64)
    db.run("INSERT INTO word_grants (token_hash,document_id,origin,issued_version,created_at) VALUES (?,?,?,?,?)",
      [grant, current.document.id, "https://localhost:3444", current.current.id, Date.now()])
    const owner = await withArtifactMutationLocksAsync(dir, [file], async () => claimWordEditLease(db, { document: current.document.id, tokenHash: grant, session: "first-pane" }))
    expect(() => assertWordEditAllowed(dir, file)).toThrow(EditingLeaseError)
    expect(() => assertMatterWordEditsAllowed(dir)).toThrow(EditingLeaseError)
    expect(() => assertWordEditAllowed(dir, file, owner)).not.toThrow()
    expect(() => claimWordEditLease(db, { ...owner, session: "second-pane" })).toThrow(EditingLeaseError)
    await expect(writeDeliverableRevision(dir, "Draft.docx", Buffer.from("must never be parsed"))).rejects.toBeInstanceOf(EditingLeaseError)
    await expect(writeDeliverable(dir, "Draft.docx", Buffer.from("must never be parsed"))).rejects.toBeInstanceOf(EditingLeaseError)
    await expect(deleteMatterDocument(dir, "Draft.docx")).rejects.toBeInstanceOf(EditingLeaseError)
    expect(readFileSync(file)).toEqual(before)
    expect(readCurrentVersion(dir, "Draft.docx")!.current.id).toBe(current.current.id)
    expect(new EditingLeaseError()).not.toBeInstanceOf(DeliverableConflictError)
    // The source path must reject before extraction, embedding or registry changes.
    db.run("UPDATE documents SET role='source' WHERE doc_path=?", [file])
    await expect(ingestDocument(dir, "Draft.docx", Buffer.from("invalid source"), { existing: "replace" })).rejects.toBeInstanceOf(EditingLeaseError)
    db.run("UPDATE documents SET role='deliverable' WHERE doc_path=?", [file])
    releaseWordEditLease(db, owner)
    releaseWordEditLease(db, owner)
    expect(() => assertWordEditAllowed(dir, file)).not.toThrow()
    expect(() => assertWordEditAllowed(dir, file, owner)).toThrow(EditingLeaseError)
    const replacement = claimWordEditLease(db, { ...owner, session: "second-pane" })
    expect(replacement.generation).not.toBe(owner.generation)
    expect(() => releaseWordEditLease(db, owner)).toThrow(EditingLeaseError)
    expect(() => renewWordEditLease(db, owner)).toThrow(EditingLeaseError)
    expect(wordLeaseState(db, owner.document, replacement).state).toBe("owned")
    db.run("UPDATE word_grants SET document_id='different-document' WHERE token_hash=?", [grant])
    expect(() => assertWordEditAllowed(dir, file)).not.toThrow()
    expect(() => assertWordEditAllowed(dir, file, replacement)).toThrow(EditingLeaseError)
    db.run("UPDATE word_grants SET document_id=? WHERE token_hash=?", [owner.document, grant])
    // Revoked grants cannot keep ordinary writers blocked.
    db.run("UPDATE word_grants SET revoked_at=? WHERE token_hash=?", [Date.now(), grant])
    expect(() => assertWordEditAllowed(dir, file)).not.toThrow()
    expect(() => assertWordEditAllowed(dir, file, replacement)).toThrow(EditingLeaseError)
    clearWordEditLease(db, owner.document)
    db.close()
    await deleteMatterDocument(dir, "Draft.docx")
    await writeDeliverable(dir, "Draft.docx", await bytes())
    expect(readCurrentVersion(dir, "Draft.docx")!.document.id).not.toBe(owner.document)
    expect(() => assertWordEditAllowed(dir, file, owner)).toThrow(EditingLeaseError)
  } finally { rmSync(dir, { recursive: true, force: true }) }
})
