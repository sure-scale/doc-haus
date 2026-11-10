import { afterAll, expect, test } from "bun:test"
import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { Document, Packer, Paragraph } from "docx"
import { readMatterSourceVersions, readSourceDependencyStatus } from "../../../dochaus/lib/source-version-dependencies"
import { initLeanState, leanSourceVersions, readLeanState, updateLeanStage } from "../../../dochaus/workflows/lean/state"
import { openDb } from "./db"
import { captureVersion, detachVersionedDocument, readDocumentVersions, readVersion, sha256Hex, type VersionWrite } from "./document-versions"
import { registerLeanDeliverable } from "./lean-deliverables"
import { leanDraftingStatus } from "./lean-drafting"
import { recordDocumentReceived, relinkDocumentReceivedDocument } from "./document-exchanges"

const dirs: string[] = []
afterAll(() => dirs.forEach(dir => rmSync(dir, { recursive: true, force: true })))
function matter() { const dir = mkdtempSync(path.join(tmpdir(), "dochaus-source-versions-")); dirs.push(dir); return dir }
const docx = (text: string) => Packer.toBuffer(new Document({ sections: [{ children: [new Paragraph(text)] }] }))
function keep(dir: string, name: string, bytes: Uint8Array, version: VersionWrite = { origin: "upload" }) {
  const file = path.join(dir, name), db = openDb(dir)
  try {
    const saved = captureVersion(db, dir, { name, bytes, prior: existsSync(file) ? readFileSync(file) : undefined, version })
    writeFileSync(file, bytes)
    return saved
  } finally { db.close() }
}
function registerSource(dir: string, name: string) {
  const db = openDb(dir)
  try { db.run("INSERT OR IGNORE INTO documents(doc_path,name,created_at,role) VALUES(?,?,1,'source')", [path.join(dir, name), name]) }
  finally { db.close() }
}
function completedSourceStage(dir: string) {
  const state = initLeanState(dir, { request: "Prepare the matter documents.", ownerSessionID: "synthetic-run", model: null })
  return updateLeanStage(dir, state, "intake", { status: "done", outputs: ["intake.json"], source_versions: readMatterSourceVersions(dir) })
}

test("published memo, checklist and issue-register outputs keep their source versions and all become stale after a received source revision", async () => {
  const dir = matter(), sourceName = "Source agreement.docx"
  const first = keep(dir, sourceName, await docx("Delivery is on Monday."))
  registerSource(dir, sourceName)
  completedSourceStage(dir)
  const snapshot = leanSourceVersions(readLeanState(dir))
  expect(snapshot).toEqual({ versions: [{ documentId: first.document_id, versionId: first.id, sha256: sha256Hex(readFileSync(path.join(dir, sourceName))) }], complete: true })
  const names = ["Advice memorandum.docx", "Closing checklist.docx", "Issue register.docx"]
  for (const name of names) {
    expect(await registerLeanDeliverable(dir, name, await docx(`Derived output: ${name}`), null)).toBe(name)
    expect(readDocumentVersions(dir, name).versions[0]?.source_dependencies).toMatchObject({ status: "current", complete: true,
      sources: [{ versionId: first.id, versionNumber: 1, currentVersionNumber: 1, status: "current" }] })
  }
  const next = keep(dir, sourceName, await docx("Delivery is on Tuesday."), { origin: "revision", parent: first.id,
    event: { kind: "received", detail: { requestId: "synthetic-receipt" } } })
  for (const name of names) expect(readDocumentVersions(dir, name).versions[0]?.source_dependencies).toMatchObject({ status: "stale",
    sources: [{ versionId: first.id, currentVersionId: next.id, currentVersionNumber: 2, status: "newer-version" }] })
  expect(leanDraftingStatus(dir)?.source_dependencies.status).toBe("stale")
  // Registration after the source changed must not claim a fresh source basis
  // for work already made by the saved stage.
  await registerLeanDeliverable(dir, "Later delivery of old work.docx", await docx("The same earlier analysis."), null)
  expect(readDocumentVersions(dir, "Later delivery of old work.docx").versions[0]?.source_dependencies.status).toBe("stale")
})

test("a derived revision, copy and restore inherit the basis they actually came from", async () => {
  const dir = matter(), name = "Source.docx", source = keep(dir, name, await docx("Source version one."))
  registerSource(dir, name)
  const old = readMatterSourceVersions(dir)
  const output1 = keep(dir, "Memo.docx", await docx("Earlier memo."), { origin: "generated", sourceVersions: old })
  keep(dir, name, await docx("Source version two."), { origin: "revision" })
  const fresh = readMatterSourceVersions(dir)
  const output2 = keep(dir, "Memo.docx", await docx("Refreshed memo."), { origin: "generated", sourceVersions: fresh })
  expect(readDocumentVersions(dir, "Memo.docx").versions[0]?.source_dependencies.status).toBe("current")
  const changed = keep(dir, "Memo.docx", await docx("A lawyer amended the memo."), { origin: "revision", parent: output2.id })
  expect(readDocumentVersions(dir, "Memo.docx").versions[0]?.source_dependencies.status).toBe("current")
  keep(dir, "Memo copy.docx", await docx("A copy of the amended memo."), { origin: "copy", derivedFrom: "Memo.docx" })
  expect(readDocumentVersions(dir, "Memo copy.docx").versions[0]?.source_dependencies.sources[0]?.versionId).not.toBe(source.id)
  keep(dir, "Memo.docx", await docx("Earlier memo."), { origin: "restore", event: { kind: "restore", detail: { restored_from: output1.id } } })
  expect(readDocumentVersions(dir, "Memo.docx").versions[0]?.source_dependencies).toMatchObject({ status: "stale", sources: [{ versionId: source.id }] })
  expect(JSON.parse(readVersion(dir, changed.id)!.version.source_versions!)).toEqual(fresh)
})

test("legacy and missing source provenance stay unknown, while outside edits and removed identities make recorded dependents stale", async () => {
  const dir = matter(), name = "Source.docx"
  const source = keep(dir, name, await docx("Original source."))
  registerSource(dir, name)
  const snapshot = readMatterSourceVersions(dir)
  keep(dir, "Legacy memo.docx", await docx("A memo without version evidence."), { origin: "generated" })
  expect(readDocumentVersions(dir, "Legacy memo.docx").versions[0]?.source_dependencies.status).toBe("unknown")
  writeFileSync(path.join(dir, name), await docx("An external edit."))
  expect(readSourceDependencyStatus(dir, snapshot)).toMatchObject({ status: "stale", sources: [{ status: "unrecorded-change" }] })
  expect(readMatterSourceVersions(dir).complete).toBe(false)
  const db = openDb(dir)
  try { detachVersionedDocument(db, dir, name) } finally { db.close() }
  // A later file with this name is a different identity; do not relink prior work.
  keep(dir, name, await docx("Another document with the same name."))
  expect(readSourceDependencyStatus(dir, snapshot)).toMatchObject({ status: "stale", sources: [{ documentId: source.document_id, status: "removed" }] })
  const legacy = matter()
  registerSource(legacy, "Unversioned source.docx")
  writeFileSync(path.join(legacy, "Unversioned source.docx"), await docx("Legacy input."))
  expect(readMatterSourceVersions(legacy)).toEqual({ versions: [], complete: false })
  expect(readSourceDependencyStatus(legacy, readMatterSourceVersions(legacy)).status).toBe("unknown")
})

test("a linked received source branch makes dependent work stale while canonical current remains unchanged", async () => {
  const dir = matter(), source = keep(dir, "Agreement.docx", await docx("The current agreement source."))
  registerSource(dir, "Agreement.docx")
  completedSourceStage(dir)
  await registerLeanDeliverable(dir, "Advice.docx", await docx("Analysis of the current agreement."), null)
  const returned = keep(dir, "Unmatched return.docx", await docx("Their newly received agreement wording."))
  const receipt = recordDocumentReceived(dir, { versionId: returned.id, requestId: "received-separately" })
  relinkDocumentReceivedDocument(dir, { exchangeId: receipt.id, targetDocumentId: source.document_id, baseVersionId: source.id,
    expectedDocumentId: returned.document_id, expectedBaseVersionId: null, requestId: "link-new-source-branch" })
  expect(readDocumentVersions(dir, "Agreement.docx").document?.current_version).toBe(source.id)
  expect(readDocumentVersions(dir, "Advice.docx").versions[0]?.source_dependencies).toMatchObject({ status: "stale", sources: [{ status: "newer-received-branch",
    linkedReturn: { exchangeId: receipt.id, sourceName: "Unmatched return.docx", versionNumber: 1 } }] })
  expect(leanDraftingStatus(dir)?.source_dependencies.status).toBe("stale")
})
