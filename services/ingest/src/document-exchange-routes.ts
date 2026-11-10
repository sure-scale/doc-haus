import { Hono } from "hono"
import { existsSync, readFileSync } from "node:fs"
import { attachmentDocumentDisposition } from "./content-disposition"
import { compareVersions } from "../../../dochaus/lib/version-compare"
import { stripWordLink, WordPackageError } from "../../../dochaus/lib/word-linked-docx"
import { EditingLeaseError } from "../../../dochaus/lib/word-edit-lease"
import { DeliverableConflictError } from "./deliverable"
import { SourceDocumentConflictError } from "./ingest"
import { type VersionEventInput } from "./document-versions"
import { readDocumentTarget } from "./registered-document"
import { oneLine, readDocumentVersions, readVersion, readVersionBytes, sha256Hex } from "./document-versions"
import { DocumentExchangeError, nextDocumentDraftNumber, prepareDocumentDownload, registerDocumentSent, prepareDocumentReceivedEvent, readDocumentExchangeRequest, readDocumentExchanges, readDocumentExchangePayload, relinkDocumentReceivedBase, relinkDocumentReceivedDocument, readDocumentReceivedBranch } from "./document-exchanges"
import { inspectDocumentReturn, readDocumentReturnMarker } from "./document-return-matching"
import { readDocumentReturnEmailSource } from "./document-return-email-source"
import { DocumentReturnTieBreakError, tieBreakDocumentReturn } from "./document-return-tie-break"

const text = (value: unknown) => typeof value === "string" ? value.trim().replace(/[\r\n]+/g, " ").slice(0, 500) : undefined
function formText(body: Record<string, unknown>, key: string) {
  if (body[key] !== undefined && typeof body[key] !== "string") throw new DocumentExchangeError(`Provide ${key} as text.`, "invalid_return_input", 400)
  return text(body[key])
}
function available(directory: string, id: string) {
  const kept = readVersion(directory, id)
  if (!kept || kept.document.removed_at != null) throw new DocumentExchangeError("That retained document version is unavailable.", "version_missing", 404)
  if (kept.version.removed) throw new DocumentExchangeError("That version was removed by redaction.", "removed_by_redaction", 410)
  if (!readVersionBytes(directory, kept.version)) throw new DocumentExchangeError("The kept version is missing or damaged.", "version_damaged")
  return kept
}
function comparison(directory: string, leftVersion: string, rightVersion: string) {
  const left = readVersion(directory, leftVersion), right = readVersion(directory, rightVersion)
  return { leftVersion, rightVersion, ...(left ? { left: left.document.name } : {}), ...(right ? { right: right.document.name } : {}) }
}
function sentResponse(bytes: Uint8Array, name: string, ext: string, draftNumber: number | undefined, exchangeId: number) {
  return new Response(new Uint8Array(bytes), { headers: {
    "Content-Type": ext === ".docx" ? "application/vnd.openxmlformats-officedocument.wordprocessingml.document" : "application/octet-stream",
    "Content-Disposition": attachmentDocumentDisposition(name),
    "X-DocHaus-Draft-Number": String(draftNumber), "X-DocHaus-Exchange-Id": String(exchangeId),
  } })
}

// Exchanges extend the version store. Inspecting a file and ordinary downloads
// never record a send; the explicit action records the exact downloaded bytes.
export function documentExchangeRoutes(input: {
  directory: (id: string) => string
  saveReturn: (directory: string, target: { name: string; bytes: Uint8Array; prior: string; parent: string; actor?: string; event: VersionEventInput }) => Promise<{ version?: { id: string; number: number } | null }>
}) {
  const routes = new Hono()
  routes.onError((error, c) => {
    if (error instanceof DocumentExchangeError) return c.json({ error: error.message, code: error.code }, error.status)
    if (error instanceof DocumentReturnTieBreakError) return c.json({ error: error.message, code: error.code }, error.status)
    if (error instanceof WordPackageError) return c.json({ error: error.message, code: error.code }, error.status)
    if (error instanceof EditingLeaseError) return c.json({ error: error.message, code: error.code, expires_at: error.expiresAt }, error.status)
    if (error instanceof DeliverableConflictError || error instanceof SourceDocumentConflictError) return c.json({ error: error.message, code: "return_current_changed" }, 409)
    throw error
  })
  routes.get("/:id/documents/exchanges", (c) => {
    const directory = input.directory(c.req.param("id"))
    const name = c.req.query("name") ?? ""
    const { document } = readDocumentVersions(directory, name)
    if (!document) return c.json({ exchanges: [] })
    return c.json({ exchanges: readDocumentExchanges(directory, document.id) })
  })
  routes.post("/:id/documents/versions/:version/send", async (c) => {
    const directory = input.directory(c.req.param("id"))
    const value: unknown = await c.req.json().catch(() => undefined)
    if (!value || typeof value !== "object" || Array.isArray(value)) return c.json({ error: "Provide the send details as a JSON object." }, 400)
    const body = value as Record<string, unknown>
    if (typeof body.requestId !== "string" || !body.requestId.trim() || !/^[A-Za-z0-9._:-]{1,160}$/.test(body.requestId))
      return c.json({ error: "A request identity is required to record this send." }, 400)
    if (body.marker !== undefined && typeof body.marker !== "boolean") return c.json({ error: "Choose whether to include the matching ID." }, 400)
    if (body.baseVersionId !== undefined && (typeof body.baseVersionId !== "string" || !body.baseVersionId.trim()) || body.party !== undefined && typeof body.party !== "string")
      return c.json({ error: "Provide the party and optional comparison base as text." }, 400)
    const versionId = c.req.param("version")
    const baseVersionId = text(body.baseVersionId)
    // A retry returns the durable exact send. Changed defaults (including the
    // firm's marker preference) do not prepare another view or alter that send.
    const previous = readDocumentExchangeRequest(directory, body.requestId)
    if (previous) {
      if (previous.kind !== "sent" || previous.versionId !== versionId || previous.baseVersionId !== (baseVersionId ?? null) || previous.party !== (text(body.party) || null))
        throw new DocumentExchangeError("This request identity already records a different send.", "request_conflict")
      const saved = readDocumentExchangePayload(directory, previous.id)
      if (body.marker !== undefined && saved.exchange.payload!.ext === ".docx") {
        const marker = readDocumentReturnMarker(saved.bytes)
        if (Boolean(marker) !== body.marker || marker && (marker.versionId !== versionId || marker.documentId !== previous.documentId))
          throw new DocumentExchangeError("This send request already records a different matching-ID choice.", "request_conflict")
      }
      return sentResponse(saved.bytes, saved.name, saved.exchange.payload!.ext, saved.exchange.draftNumber, saved.exchange.id)
    }
    let outbound: Uint8Array | undefined
    if (baseVersionId) {
      const left = readVersion(directory, baseVersionId), right = readVersion(directory, versionId)
      if (!left || !right || left.document.id !== right.document.id || left.document.removed_at != null || right.document.removed_at != null || left.version.removed || right.version.removed)
        return c.json({ error: "Choose two available versions of the same document." }, 409)
      const leftBytes = readVersionBytes(directory, left.version), rightBytes = readVersionBytes(directory, right.version)
      if (!leftBytes || !rightBytes) return c.json({ error: "A comparison version is missing or damaged." }, 409)
      const result = await compareVersions({ bytes: leftBytes, label: `Version ${left.version.number}` }, { bytes: rightBytes, label: `Version ${right.version.number}` }, { detectMoves: true, headersFooters: true, ignoreFormatting: false }, new Date(right.version.created_at).toISOString())
      outbound = result.blackline
    }
    const source = readVersion(directory, versionId)
    if (!source) return c.json({ error: "The selected version is unavailable." }, 404)
    const draftNumber = nextDocumentDraftNumber(directory, source.document.id, body.requestId)
    const prepared = prepareDocumentDownload(directory, { bytes: outbound, versionId, baseVersionId, marker: body.marker as boolean | undefined, draftNumber })
    const result = registerDocumentSent(directory, { prepared, requestId: body.requestId, party: text(body.party), actor: text(c.req.header("x-dochaus-actor")) })
    return sentResponse(prepared.bytes, prepared.name, prepared.ext, result.draftNumber, result.exchange.id)
  })
  routes.get("/:id/documents/exchanges/:exchange/content", (c) => {
    const result = readDocumentExchangePayload(input.directory(c.req.param("id")), Number(c.req.param("exchange")))
    return new Response(new Uint8Array(result.bytes), { headers: { "Content-Type": "application/octet-stream", "Content-Disposition": attachmentDocumentDisposition(result.name) } })
  })
  routes.post("/:id/documents/exchanges/:exchange/base", async (c) => {
    const value: unknown = await c.req.json().catch(() => undefined)
    if (!value || typeof value !== "object" || Array.isArray(value)) return c.json({ error: "Confirm the receipt and its existing and new base versions." }, 400)
    const body = value as Record<string, unknown>, exchangeId = Number(c.req.param("exchange"))
    if (!Number.isSafeInteger(exchangeId) || exchangeId < 1 || typeof body.baseVersionId !== "string" || !body.baseVersionId || !(body.expectedBaseVersionId === null || typeof body.expectedBaseVersionId === "string" && body.expectedBaseVersionId) || typeof body.requestId !== "string" || !/^[A-Za-z0-9._:-]{1,160}$/.test(body.requestId))
      return c.json({ error: "Confirm the receipt, its existing and new base versions, and a request identity." }, 400)
    return c.json(relinkDocumentReceivedBase(input.directory(c.req.param("id")), { exchangeId, baseVersionId: body.baseVersionId, expectedBaseVersionId: body.expectedBaseVersionId, requestId: body.requestId, actor: text(c.req.header("x-dochaus-actor")) }))
  })
  routes.post("/:id/documents/exchanges/:exchange/document", async (c) => {
    const value: unknown = await c.req.json().catch(() => undefined)
    if (!value || typeof value !== "object" || Array.isArray(value)) return c.json({ error: "Confirm the receipt's current and intended document and base." }, 400)
    const body = value as Record<string, unknown>, exchangeId = Number(c.req.param("exchange"))
    if (!Number.isSafeInteger(exchangeId) || exchangeId < 1 || ["targetDocumentId", "baseVersionId", "expectedDocumentId"].some(key => typeof body[key] !== "string" || !(body[key] as string).trim()) || !(body.expectedBaseVersionId === null || typeof body.expectedBaseVersionId === "string" && body.expectedBaseVersionId) || ["expectedSourceCurrentVersionId", "expectedTargetCurrentVersionId"].some(key => body[key] !== undefined && body[key] !== null && (typeof body[key] !== "string" || !(body[key] as string).trim())) || typeof body.requestId !== "string" || !/^[A-Za-z0-9._:-]{1,160}$/.test(body.requestId))
      return c.json({ error: "Confirm the receipt's current and intended document and base, and a request identity." }, 400)
    return c.json(relinkDocumentReceivedDocument(input.directory(c.req.param("id")), { exchangeId, targetDocumentId: body.targetDocumentId as string, baseVersionId: body.baseVersionId as string, expectedDocumentId: body.expectedDocumentId as string, expectedBaseVersionId: body.expectedBaseVersionId as string | null, requestId: body.requestId, actor: text(c.req.header("x-dochaus-actor")), expectedSourceCurrentVersionId: body.expectedSourceCurrentVersionId as string | null | undefined, expectedTargetCurrentVersionId: body.expectedTargetCurrentVersionId as string | null | undefined }))
  })
  routes.get("/:id/documents/exchanges/:exchange/branch/content", c => {
    const { branch, bytes } = readDocumentReceivedBranch(input.directory(c.req.param("id")), Number(c.req.param("exchange")))
    return new Response(new Uint8Array(bytes), { headers: { "Cache-Control": "no-store", "Content-Type": branch.ext === ".docx" ? "application/vnd.openxmlformats-officedocument.wordprocessingml.document" : "application/octet-stream", "Content-Disposition": attachmentDocumentDisposition(branch.sourceName) } })
  })
  routes.post("/:id/documents/returns/inspect", async (c) => {
    const directory = input.directory(c.req.param("id"))
    const body = await c.req.parseBody().catch(() => undefined)
    if (!body) return c.json({ error: "Choose the returned file using a file upload." }, 400)
    if (!(body.file instanceof File)) return c.json({ error: "Choose the returned file." }, 400)
    const bytes = new Uint8Array(await body.file.arrayBuffer())
    const emailName = formText(body, "emailName")
    const source = emailName ? await readDocumentReturnEmailSource(directory, emailName, bytes) : undefined
    return c.json({ ...await inspectDocumentReturn(directory, { name: body.file.name, bytes, source }), ...(source ? { source } : {}) })
  })
  // Explicitly requested phase-three judgment; ordinary inspect/save never
  // calls a model. The advisory carries no authority to save or relink a file.
  routes.post("/:id/documents/returns/tie-break", async (c) => {
    const directory = input.directory(c.req.param("id"))
    const body = await c.req.parseBody().catch(() => undefined)
    if (!body || !(body.file instanceof File) || !body.file.name.toLowerCase().endsWith(".docx")) return c.json({ error: "Choose the returned Word file using a file upload." }, 400)
    const bytes = new Uint8Array(await body.file.arrayBuffer()), emailName = formText(body, "emailName")
    const source = emailName ? await readDocumentReturnEmailSource(directory, emailName, bytes) : undefined
    try {
      const result = await tieBreakDocumentReturn(directory, { name: body.file.name, bytes, source }, c.req.raw.signal)
      if (emailName) await readDocumentReturnEmailSource(directory, emailName, bytes)
      return c.json({ ...result, ...(source ? { source } : {}) })
    } catch (error) { if (c.req.raw.signal.aborted) return new Response(null, { status: 499 }); throw error }
  })
  routes.post("/:id/documents/returns", async (c) => {
    const directory = input.directory(c.req.param("id"))
    const body = await c.req.parseBody().catch(() => undefined)
    if (!body) return c.json({ error: "Choose the returned file using a file upload." }, 400)
    if (!(body.file instanceof File) || !body.file.name.toLowerCase().endsWith(".docx")) return c.json({ error: "Choose a returned Word file." }, 400)
    const baseId = formText(body, "baseVersionId"), requestId = formText(body, "requestId"), prior = formText(body, "priorSha256"), party = formText(body, "party")
    if (!baseId || !requestId || !prior || !/^[a-f0-9]{64}$/.test(prior)) return c.json({ error: "Confirm a document, base version and current copy before saving this return." }, 400)
    const base = available(directory, baseId)
    const target = readDocumentTarget(directory, base.document.name)
    if (!target) return c.json({ error: "The selected document is no longer available." }, 409)
    // Exact repeat after a lost response: the event supplies the durable result.
    const previous = readDocumentExchangeRequest(directory, requestId)
    const incoming = new Uint8Array(await body.file.arrayBuffer())
    const emailName = formText(body, "emailName")
    // Authenticate attachment provenance against original uploaded bytes before
    // stripping the native save capability for storage.
    const source = emailName ? await readDocumentReturnEmailSource(directory, emailName, incoming) : undefined
    const bytes = stripWordLink(incoming)
    const inspection = await inspectDocumentReturn(directory, { name: body.file.name, bytes, source })
    // Inspection can await package readers. Recheck retained availability before
    // returning a prior receipt or allowing the normal writer to take its lock.
    const liveBase = available(directory, baseId)
    if (liveBase.document.id !== base.document.id) throw new DocumentExchangeError("The selected base changed. Inspect the return again.", "return_base_changed")
    if (previous) {
      const kept = available(directory, previous.versionId)
      if (previous.kind !== "received" || previous.documentId !== base.document.id || kept.version.sha256 !== sha256Hex(bytes) || previous.baseVersionId !== baseId || previous.party !== (party || null))
        return c.json({ error: "This request identity was already used for a different returned file." }, 409)
      if (source && (["documentName", "attachmentId", "sender", "date"] as const).some(key => oneLine(source[key]) !== (previous.source?.[key] ?? "")))
        return c.json({ error: "This request identity already records a different receipt source." }, 409)
      const comparisonBase = previous.association ? readDocumentReceivedBranch(directory, previous.id).branch.baseVersionId : baseId
      return c.json({ name: target.name, version: { id: kept.version.id, number: kept.version.number }, exchange: previous, comparison: comparison(directory, comparisonBase, kept.version.id) })
    }
    const hash = sha256Hex(bytes)
    const duplicate = readDocumentVersions(directory, target.name).versions.find(version => {
      const found = version.sha256 === hash && !version.removed ? readVersion(directory, version.id) : undefined
      return found && readVersionBytes(directory, found.version)
    })
    if (duplicate) return c.json({ error: `This exact file is already kept as version ${duplicate.number}. Nothing new was saved.`, code: "duplicate_version", version: { id: duplicate.id, number: duplicate.number } }, 409)
    if (!existsSync(target.docPath) || sha256Hex(readFileSync(target.docPath)) !== prior) return c.json({ error: "The current document changed. Inspect it again before saving this return." }, 409)
    const candidate = inspection.candidates.find((item) => item.versionId === baseId)
    const event = prepareDocumentReceivedEvent(directory, { baseVersionId: baseId, requestId, party, actor: text(c.req.header("x-dochaus-actor")), source, match: candidate ? { signal: `${candidate.signals.join(", ")}, lawyer-confirmed`, reason: `${candidate.reason} The lawyer explicitly selected this document and base version.` } : { signal: "lawyer-confirmed", reason: "The lawyer explicitly selected this document and base version." } })
    const written = await input.saveReturn(directory, { name: target.name, bytes, prior, parent: baseId, actor: text(c.req.header("x-dochaus-actor")), event })
    if (!written.version) return c.json({ error: "The returned version was not recorded. Refresh the version history before retrying." }, 409)
    const exchange = readDocumentExchangeRequest(directory, requestId)
    if (!exchange) return c.json({ error: "The returned receipt is unavailable. Refresh the history before retrying." }, 409)
    return c.json({ ...written, name: target.name, exchange, comparison: comparison(directory, baseId, written.version.id) })
  })
  return routes
}
