import { afterEach, expect, test } from "bun:test"
import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { strFromU8, unzipSync } from "fflate"
import type { ChronologyCoverage, ChronologyDate } from "../../../dochaus/lib/chronology"
import { readPinnedSourceVersions, readSourceDependencyStatus } from "../../../dochaus/lib/source-version-dependencies"
import { chronologyRoutes } from "./chronology-routes"
import { currentChronologySources, mutateChronology, newChronologyEvent, readChronology } from "./chronology-store"
import { captureVersion, readDocumentVersions, readVersion, sha256Hex, type VersionWrite } from "./document-versions"
import { openDb, upsertDocument } from "./db"

const directories: string[] = []
afterEach(() => directories.splice(0).forEach(directory => rmSync(directory, { recursive: true, force: true })))
const date: ChronologyDate = { original: "2025", precision: "year", start: "2025" }
function matter() {
  const directory = mkdtempSync(path.join(tmpdir(), "chronology-source-versions-"))
  directories.push(directory)
  writeFileSync(path.join(directory, "matter.json"), JSON.stringify({ title: "Synthetic source-basis review" }))
  return directory
}
function keep(directory: string, name: string, text: string, version: VersionWrite = { origin: "upload" }) {
  const file = path.join(directory, name), bytes = Buffer.from(text), db = openDb(directory)
  try {
    const kept = captureVersion(db, directory, { name, bytes, prior: existsSync(file) ? readFileSync(file) : undefined, version })
    writeFileSync(file, bytes)
    upsertDocument(db, file, name, 1, null)
    return { ...kept, name, sha256: sha256Hex(bytes) }
  } finally { db.close() }
}
function event(directory: string, name: string, id = "entry") {
  const source = currentChronologySources(directory).find(source => source.documentName === name)!
  const text = readFileSync(path.join(directory, name), "utf8")
  return newChronologyEvent({ id, title: "A source reports delivery", detail: "The account remains qualified.", date,
    accounts: [{ id: `account-${id}`, title: "Delivery report", detail: "The source's account.", date, privilege: "cleared", significance: "relevant", status: "reported", people: [], gaps: [],
      evidence: [{ id: `evidence-${id}`, documentName: name, sourceId: source.sourceId, version: source.version, excerpt: text,
        section: "Account", charStart: 0, charEnd: text.length, verified: true, freshness: "current", provenance: "source" }] }] })
}
function coverage(directory: string, name: string): ChronologyCoverage {
  const source = currentChronologySources(directory).find(source => source.documentName === name)!
  return { sourceId: source.sourceId, documentName: name, version: source.version, status: "complete", uploaded: true, textExtracted: true,
    sections: 1, processedSections: [0], completedKeys: [], events: 1, issues: [], suppliedChronology: false }
}
async function exportSaved(directory: string, input?: { scope: "filtered"; eventIds: string[] }) {
  const request = { expectedRevision: readChronology(directory).revision, scope: "all-visible", includeHidden: false, ...input }
  const response = await chronologyRoutes(() => directory).request("/matters/synthetic/chronology/export", {
    method: "POST", headers: { "Content-Type": "application/json", "x-dochaus-actor": "Synthetic reviewer" }, body: JSON.stringify(request),
  })
  expect(response.status).toBe(200)
  const name = response.headers.get("X-DocHaus-Export-Name")!, bytes = new Uint8Array(await response.arrayBuffer())
  const history = readDocumentVersions(directory, name), version = history.versions[0]!
  expect(history.versions).toHaveLength(1)
  expect(version.origin).toBe("generated")
  expect(version.actor).toBe("Synthetic reviewer")
  expect(version.sha256).toBe(sha256Hex(bytes))
  expect(bytes).toEqual(new Uint8Array(readFileSync(path.join(directory, name))))
  return { version, bytes, basis: JSON.parse(readVersion(directory, version.id)!.version.source_versions!) }
}

test("actual chronology export retains saved source v1 after v2 arrives and does not adopt new projected pending sources", async () => {
  const directory = matter(), source = keep(directory, "Report.md", "The original report bears a teal seal.")
  const saved = mutateChronology(directory, 0, "synthetic fixture", state => {
    state.created = true; state.events = [event(directory, source.name)]; state.coverage = [coverage(directory, source.name)]
  })
  const first = await exportSaved(directory)
  expect(first.basis).toEqual({ versions: [{ documentId: source.document_id, versionId: source.id, sha256: source.sha256 }], complete: true })
  expect(first.version.source_dependencies.status).toBe("current")
  const next = keep(directory, source.name, "The revised report bears a violet seal.", { origin: "revision" })
  const unrelated = keep(directory, "New report.md", "A newly arrived unrelated report.")
  const preview = await chronologyRoutes(() => directory).request("/matters/synthetic/chronology/export/preview", {
    method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ expectedRevision: saved.revision, scope: "all-visible", includeHidden: false }),
  })
  expect((await preview.json()).coverage).toContainEqual(expect.objectContaining({ documentName: unrelated.name, status: "pending" }))
  const second = await exportSaved(directory)
  expect(second.basis).toEqual(first.basis)
  expect(second.version.source_dependencies).toMatchObject({ status: "stale", sources: [{ versionId: source.id, currentVersionId: next.id, status: "newer-version" }] })
  expect(strFromU8(unzipSync(second.bytes)["word/document.xml"]!)).not.toContain("original report bears a teal seal")
  expect(readChronology(directory)).toEqual(saved)
  expect(readDocumentVersions(directory, source.name).versions).toHaveLength(2)
  expect(readDocumentVersions(directory, unrelated.name).versions).toHaveLength(1)
})

test("filtered export includes recorded event and account date anchors plus saved coverage, excluding unselected event-only pins", async () => {
  const directory = matter(), source = keep(directory, "Report.md", "An account of delivery.")
  const anchor = keep(directory, "Anchor.md", "The starting date was recorded.")
  const excluded = keep(directory, "Excluded.md", "An unrelated excluded account.")
  const selected = event(directory, source.name), other = event(directory, excluded.name, "excluded")
  const relative = { expression: "Two days after the recorded date", anchorDocument: anchor.name, anchorVersion: anchor.sha256 }
  selected.date = { ...date, relative }
  selected.accounts[0]!.date = { ...date, relative: { ...relative, anchorDocument: source.name, anchorVersion: source.sha256 } }
  mutateChronology(directory, 0, "synthetic fixture", state => {
    state.created = true; state.events = [selected, other]; state.coverage = [coverage(directory, source.name)]
  })
  const output = await exportSaved(directory, { scope: "filtered", eventIds: [selected.id] })
  expect(output.basis.complete).toBe(true)
  expect(output.basis.versions.map((pin: { versionId: string }) => pin.versionId).sort()).toEqual([source.id, anchor.id].sort())
  expect(output.version.source_dependencies.status).toBe("current")
  // A legacy relative date that names a source without its executed hash does
  // not gain the current source version just because it can be read today.
  mutateChronology(directory, readChronology(directory).revision, "synthetic fixture", state => { delete state.events[0]!.date.relative!.anchorVersion })
  const legacy = await exportSaved(directory, { scope: "filtered", eventIds: [selected.id] })
  expect(legacy.basis.complete).toBe(false)
  expect(legacy.basis.versions.map((pin: { versionId: string }) => pin.versionId)).toEqual([source.id])
  expect(legacy.version.source_dependencies.status).toBe("unknown")
})

test("repeated retained hashes and unversioned source pins stay unknown rather than choosing current or backfilling", async () => {
  const directory = matter(), original = keep(directory, "Report.md", "The first wording.")
  mutateChronology(directory, 0, "synthetic fixture", state => { state.created = true; state.events = [event(directory, original.name)] })
  keep(directory, original.name, "The second wording.", { origin: "revision" })
  keep(directory, original.name, "The first wording.", { origin: "restore" })
  const ambiguous = await exportSaved(directory)
  expect(ambiguous.basis).toEqual({ versions: [], complete: false })
  expect(ambiguous.version.source_dependencies.status).toBe("unknown")
  const legacy = path.join(directory, "Legacy.md"), db = openDb(directory)
  try { writeFileSync(legacy, "Unversioned input."); upsertDocument(db, legacy, "Legacy.md", 1, null) } finally { db.close() }
  const pin = { name: "Legacy.md", sha256: sha256Hex(readFileSync(legacy)) }
  expect(readPinnedSourceVersions(directory, [pin])).toEqual({ versions: [], complete: false })
  expect(readDocumentVersions(directory, "Legacy.md").document).toBeNull()
})

test("known removed or damaged recorded inputs preserve their old pins and cannot be reported current", () => {
  const directory = matter(), source = keep(directory, "Report.md", "Recorded input.")
  const pin = { name: source.name, sha256: source.sha256 }
  expect(readPinnedSourceVersions(directory, [pin]).complete).toBe(true)
  const held = readVersion(directory, source.id)!.version
  rmSync(path.join(directory, ".dochaus", "revisions", source.document_id, `${source.sha256}${held.ext}`))
  const damaged = readPinnedSourceVersions(directory, [pin])
  expect(damaged).toEqual({ versions: [{ documentId: source.document_id, versionId: source.id, sha256: source.sha256 }], complete: false })
  expect(readSourceDependencyStatus(directory, damaged).status).toBe("unknown")
  const db = openDb(directory)
  try { db.run("UPDATE document_versions SET removed='redaction' WHERE id=?", [source.id]) } finally { db.close() }
  const removed = readPinnedSourceVersions(directory, [pin])
  expect(removed).toEqual(damaged)
  expect(readSourceDependencyStatus(directory, removed)).toMatchObject({ status: "stale", sources: [{ versionId: source.id, status: "removed" }] })
})
