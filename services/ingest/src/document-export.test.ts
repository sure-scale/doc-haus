import { afterEach, expect, spyOn, test } from "bun:test"
import { Hono } from "hono"
import { randomUUID } from "node:crypto"
import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { Document, Packer, Paragraph } from "docx"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { linkedWordDocx, readWordLink } from "../../../dochaus/lib/word-linked-docx"
import { compareVersions } from "../../../dochaus/lib/version-compare"
import { openDb } from "./db"
import { documentExportRoutes, prepareDocumentExportCopy, type DocumentExportView } from "./document-export"
import { nextDocumentDraftNumber } from "./document-exchanges"
import { captureVersion, purgeVersionBytes, readDocumentVersions, readVersion, readVersionBytes, sha256Hex, versionFile } from "./document-versions"
import { inspectDocumentReturn, readDocumentReturnMarker } from "./document-return-matching"
import * as preferences from "./preferences"

const directories: string[] = []
afterEach(() => { for (const directory of directories.splice(0)) rmSync(directory, { recursive: true, force: true }) })
const docx = (text: string) => Packer.toBuffer(new Document({ sections: [{ children: [new Paragraph(text)] }] }))
function fixture(proposals?: (directory: string, source: { name: string; docPath: string; bytes: Uint8Array }) => Promise<DocumentExportView>) {
  const directory = mkdtempSync(path.join(tmpdir(), "dochaus-export-copy-")); directories.push(directory)
  openDb(directory).close()
  const app = new Hono().route("/matters", documentExportRoutes({ directory: () => directory,
    proposals: proposals ?? (async (_directory, source) => ({ bytes: source.bytes })),
  }))
  const keep = (bytes: Uint8Array, name = "Agreement.docx") => {
    const file = path.join(directory, name), db = openDb(directory)
    try {
      const held = captureVersion(db, directory, { name, bytes, ...(existsSync(file) ? { prior: readFileSync(file) } : {}), version: { origin: "upload" } })
      writeFileSync(file, bytes)
      return readVersion(directory, held.id)!.version
    } finally { db.close() }
  }
  return { directory, app, keep, get: (query: string) => app.request(`/matters/synthetic/documents/export?${query}`) }
}
const body = (bytes: Uint8Array) => strFromU8(unzipSync(bytes)["word/document.xml"]!)
function edit(bytes: Uint8Array, before: string, after: string) {
  const parts = unzipSync(bytes); parts["word/document.xml"] = strToU8(body(bytes).replace(before, after)); return zipSync(parts)
}

test("ordinary current and older-version exports add truthful markers without touching exact sources, versions, events or draft counts", async () => {
  const s = fixture(), old = s.keep(await docx("The delivery is Monday."))
  const current = s.keep(await docx("The delivery is Tuesday."))
  const canonical = readFileSync(path.join(s.directory, "Agreement.docx")), oldRaw = readVersionBytes(s.directory, old)!
  const prefs = spyOn(preferences, "readDraftingPreferences").mockReturnValue({ ...preferences.readDraftingPreferences(), documentMatchingId: true })
  try {
    const currentResponse = await s.get("name=Agreement.docx"), currentCopy = new Uint8Array(await currentResponse.arrayBuffer())
    expect(currentResponse.status).toBe(200)
    expect(currentResponse.headers.get("X-DocHaus-Source-Version")).toBe(current.id)
    expect(readDocumentReturnMarker(currentCopy)).toEqual({ documentId: current.document_id, versionId: current.id })
    expect(body(currentCopy)).toBe(body(canonical))
    const historicalResponse = await s.get(`versionId=${old.id}`), oldCopy = new Uint8Array(await historicalResponse.arrayBuffer())
    expect(historicalResponse.status).toBe(200)
    expect(readDocumentReturnMarker(oldCopy)).toEqual({ documentId: old.document_id, versionId: old.id })
    const returned = await inspectDocumentReturn(s.directory, { name: "Renamed.docx", bytes: edit(oldCopy, "Monday", "Wednesday") })
    expect(returned.selected?.versionId).toBe(old.id)
    expect(readFileSync(path.join(s.directory, "Agreement.docx"))).toEqual(canonical)
    expect(readVersionBytes(s.directory, old)).toEqual(oldRaw)
    expect(sha256Hex(oldRaw)).toBe(old.sha256)
    expect(readDocumentVersions(s.directory, "Agreement.docx").versions).toHaveLength(2)
    expect(readDocumentVersions(s.directory, "Agreement.docx").events).toHaveLength(0)
    expect(nextDocumentDraftNumber(s.directory, old.document_id)).toBe(1)
    prefs.mockReturnValue({ ...preferences.readDraftingPreferences(), documentMatchingId: false })
    const off = await s.get(`versionId=${old.id}`)
    expect(off.status).toBe(200)
    expect(readDocumentReturnMarker(new Uint8Array(await off.arrayBuffer()))).toBeUndefined()
    expect(readVersionBytes(s.directory, old)).toEqual(oldRaw)
  } finally { prefs.mockRestore() }
})

test("proposal view exports preserve the rendered package and source pin, strip capabilities and refuse a changed source", async () => {
  let view: Uint8Array, changed = false
  const s = fixture(async (_directory, source) => {
    expect(body(source.bytes)).toContain("Monday")
    if (changed) writeFileSync(source.docPath, await docx("A concurrent outside edit."))
    return { bytes: view, notice: "Original tracked changes are disclosed." }
  })
  const version = s.keep(await docx("The delivery is Monday."))
  const trackedParts = unzipSync(readVersionBytes(s.directory, version)!)
  trackedParts["word/document.xml"] = strToU8(body(readVersionBytes(s.directory, version)!).replace("</w:p>", '<w:ins w:id="9" w:author="Synthetic reviewer" w:date="2026-01-01T00:00:00Z"><w:r><w:t> Proposed addition.</w:t></w:r></w:ins></w:p>'))
  const tracked = zipSync(trackedParts)
  view = linkedWordDocx(tracked, { version: 1, origin: "https://example.invalid", apiOrigin: "https://example.invalid", matter: randomUUID(), matterName: "Synthetic", document: version.document_id, documentName: "Agreement.docx", baseVersion: version.id, token: "synthetic-write-capability" })
  const response = await s.get("name=Agreement.docx&proposals=1"), copy = new Uint8Array(await response.arrayBuffer())
  expect(response.status).toBe(200)
  expect(body(copy)).toBe(body(view))
  expect(body(copy)).toContain("<w:ins")
  expect(body(readVersionBytes(s.directory, version)!)).not.toContain("Proposed addition")
  expect(readWordLink(copy)).toBeNull()
  expect(readDocumentReturnMarker(copy)?.versionId).toBe(version.id)
  expect(decodeURIComponent(response.headers.get("X-DocHaus-Revision-Notice")!)).toContain("Original tracked")
  expect(readDocumentVersions(s.directory, "Agreement.docx").versions).toHaveLength(1)
  changed = true
  expect((await s.get("name=Agreement.docx&proposals=1")).status).toBe(409)
  expect(readDocumentVersions(s.directory, "Agreement.docx").events).toHaveLength(0)
})

test("an externally changed or unrecorded Word file cannot reuse a stale marker or manufacture a version", async () => {
  const s = fixture(), version = s.keep(await docx("An original Monday term."))
  const prepared = prepareDocumentExportCopy(s.directory, { bytes: readVersionBytes(s.directory, version)!, sourceVersionId: version.id, sourceSha256: version.sha256 })
  writeFileSync(path.join(s.directory, "Agreement.docx"), edit(prepared.bytes, "Monday", "Thursday"))
  const response = await s.get("name=Agreement.docx"), bytes = new Uint8Array(await response.arrayBuffer())
  expect(response.status).toBe(200)
  expect(response.headers.get("X-DocHaus-Export-Identity")).toBe("unrecorded")
  expect(response.headers.get("X-DocHaus-Source-Version")).toBeNull()
  expect(readDocumentReturnMarker(bytes)).toBeUndefined()
  expect(body(bytes)).toContain("Thursday")
  expect(readDocumentVersions(s.directory, "Agreement.docx").versions).toHaveLength(1)
  writeFileSync(path.join(s.directory, "Legacy.docx"), prepared.bytes)
  expect((await s.get("name=Legacy.docx")).headers.get("X-DocHaus-Export-Identity")).toBe("unrecorded")
  expect(readDocumentVersions(s.directory, "Legacy.docx").versions).toHaveLength(0)
})

test("prepared blacklines retain their exact reviewed wording, qualify the right source across documents and revalidate both source pins", async () => {
  const s = fixture(), left = s.keep(await docx("The delivery is Monday.")), right = s.keep(await docx("The delivery is Tuesday."), "Counterpart.docx")
  const blackline = await compareVersions({ bytes: readVersionBytes(s.directory, left)!, label: "Before" }, { bytes: readVersionBytes(s.directory, right)!, label: "After" }, { detectMoves: true, ignoreFormatting: false, headersFooters: true }, "2026-01-01T00:00:00.000Z")
  const input = { bytes: blackline.blackline, sourceVersionId: right.id, sourceSha256: right.sha256, relatedSources: [{ versionId: left.id, sha256: left.sha256 }] }
  const prepared = prepareDocumentExportCopy(s.directory, input)
  expect(body(prepared.bytes)).toBe(body(blackline.blackline))
  expect(readDocumentReturnMarker(prepared.bytes)).toEqual({ documentId: right.document_id, versionId: right.id })
  expect(readVersion(s.directory, right.id)!.version.parent).toBeNull()
  expect(readDocumentVersions(s.directory, "Counterpart.docx").events).toHaveLength(0)
  const db = openDb(s.directory)
  try { db.run("UPDATE document_versions SET removed='redaction' WHERE id=?", [left.id]) } finally { db.close() }
  purgeVersionBytes(s.directory, left.document_id)
  expect(() => prepareDocumentExportCopy(s.directory, input)).toThrow("comparison source is no longer available")
  expect((await s.get(`versionId=${left.id}`)).status).toBe(410)
  writeFileSync(versionFile(s.directory, right), "damaged")
  expect((await s.get(`versionId=${right.id}`)).status).toBe(409)
  expect((await s.get("name=Agreement.docx&versionId=conflict")).status).toBe(400)
})
