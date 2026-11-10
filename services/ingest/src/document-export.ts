import { Hono } from "hono"
import { readFileSync } from "node:fs"
import path from "node:path"
import { withArtifactMutationLocksAsync } from "../../../dochaus/capabilities/artifacts/mutation-lock"
import { stripWordLink } from "../../../dochaus/lib/word-linked-docx"
import { attachmentDocumentDisposition } from "./content-disposition"
import { DocumentExchangeError, prepareDocumentDownload } from "./document-exchanges"
import { readCurrentVersion, readVersion, readVersionBytes, sha256Hex } from "./document-versions"
import { stampDocumentReturnMarker } from "./document-return-matching"
import { readDocumentTarget, registeredDocumentName } from "./registered-document"

export type DocumentExportView = { bytes: Uint8Array; notice?: string }

/** A sharing copy, separate from exact retained/content reads. The caller owns
 * the view and pins the source before rendering it. A changed/unrecorded working
 * file stays downloadable, but cannot impersonate a retained version. */
export function prepareDocumentExportCopy(directory: string, input: {
  bytes: Uint8Array; sourceVersionId?: string; sourceSha256?: string
  relatedSources?: { versionId: string; sha256: string }[]
}) {
  for (const pin of input.relatedSources ?? []) {
    const related = readVersion(directory, pin.versionId)
    if (!related || related.document.removed_at != null || related.version.removed || related.version.sha256 !== pin.sha256 || !readVersionBytes(directory, related.version))
      throw new DocumentExchangeError("A comparison source is no longer available. Refresh before downloading.", "export_source_changed")
  }
  if (input.sourceVersionId) {
    const source = readVersion(directory, input.sourceVersionId)
    if (!source || source.document.removed_at != null || source.version.removed ||
      source.version.sha256 !== input.sourceSha256 || !readVersionBytes(directory, source.version))
      throw new DocumentExchangeError("The export source is no longer available. Refresh the document before downloading.", "export_source_changed")
    if (source.version.ext !== ".docx") throw new DocumentExchangeError("Choose a retained Word source for this copy.", "not_docx", 400)
    const copy = prepareDocumentDownload(directory, { versionId: source.version.id, bytes: input.bytes })
    return { bytes: copy.bytes, versionId: source.version.id, sha256: source.version.sha256, matchingId: copy.marker !== null }
  }
  // In particular, discard a stale marker on an externally changed working
  // file. No download creates/backfills a version to manufacture an identity.
  try { return { bytes: stampDocumentReturnMarker(stripWordLink(input.bytes), null), versionId: null, sha256: null, matchingId: false } }
  catch { throw new DocumentExchangeError("This Word copy could not be prepared safely.", "invalid_word_copy", 400) }
}

export function documentExportRoutes(input: {
  directory: (id: string) => string
  proposals: (directory: string, source: { name: string; docPath: string; bytes: Uint8Array }) => Promise<DocumentExportView>
}) {
  const routes = new Hono()
  routes.onError((error, c) => {
    if (error instanceof DocumentExchangeError) return c.json({ error: error.message, code: error.code }, error.status)
    throw error
  })
  routes.get("/:id/documents/export", async c => {
    const directory = input.directory(c.req.param("id"))
    const versionId = c.req.query("versionId"), requestedName = c.req.query("name")
    const proposals = c.req.query("proposals")
    if ((!versionId && !requestedName) || versionId && requestedName || proposals !== undefined && proposals !== "1" || versionId && proposals)
      return c.json({ error: "Choose a saved version or a working document to export.", code: "invalid_export" }, 400)
    const name = versionId ? readVersion(directory, versionId)?.document.name : registeredDocumentName(requestedName ?? "")
    if (!name) return c.json({ error: "The export document is unavailable.", code: "version_missing" }, 404)
    if (path.extname(name).toLowerCase() !== ".docx") return c.json({ error: "Only Word files need a prepared export copy.", code: "not_docx" }, 400)
    return withArtifactMutationLocksAsync(directory, [name, ".dochaus/source-corpus"], async () => {
      let bytes: Uint8Array, sourceVersionId: string | undefined, sourceSha256: string | undefined
      let filename = name, notice: string | undefined
      if (versionId) {
        const source = readVersion(directory, versionId)
        if (!source || source.document.removed_at != null) throw new DocumentExchangeError("That retained version is unavailable.", "version_missing", 404)
        if (source.version.removed) throw new DocumentExchangeError("That version was removed by redaction.", "removed_by_redaction", 410)
        const retained = readVersionBytes(directory, source.version)
        if (!retained) throw new DocumentExchangeError("The kept bytes of this version are missing or damaged.", "version_damaged")
        bytes = retained; sourceVersionId = versionId; sourceSha256 = source.version.sha256
        filename = `${path.parse(name).name} (version ${source.version.number}).docx`
      } else {
        const target = readDocumentTarget(directory, name)
        if (!target) throw new DocumentExchangeError("This document is no longer available.", "version_missing", 404)
        try { bytes = readFileSync(target.docPath) }
        catch { throw new DocumentExchangeError("This document is no longer available.", "version_missing", 404) }
        const current = readCurrentVersion(directory, target.name)
        if (current && current.document.removed_at == null && !current.current.removed && current.current.sha256 === sha256Hex(bytes)) {
          sourceVersionId = current.current.id; sourceSha256 = current.current.sha256
        }
        if (proposals) {
          const view = await input.proposals(directory, { ...target, bytes })
          // The same source/name locks cover rendering and pin validation. An
          // outside edit during rendering must not relabel the prepared view.
          const after = (() => { try { return sha256Hex(readFileSync(target.docPath)) } catch { return undefined } })()
          if (after !== sha256Hex(bytes))
            throw new DocumentExchangeError("The working document changed while preparing this copy. Download it again.", "export_source_changed")
          bytes = view.bytes; notice = view.notice
        }
      }
      const copy = prepareDocumentExportCopy(directory, { bytes, sourceVersionId, sourceSha256 })
      return new Response(new Uint8Array(copy.bytes), { headers: {
        "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "Content-Disposition": attachmentDocumentDisposition(filename), "Cache-Control": "no-store",
        "X-DocHaus-Export-Identity": copy.versionId ? "retained-version" : "unrecorded",
        ...(copy.versionId ? { "X-DocHaus-Source-Version": copy.versionId, "X-DocHaus-Source-Sha256": copy.sha256! } : {}),
        ...(notice ? { "X-DocHaus-Revision-Notice": encodeURIComponent(notice) } : {}),
      } })
    })
  })
  return routes
}
