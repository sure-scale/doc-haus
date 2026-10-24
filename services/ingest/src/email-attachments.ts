import type { Database } from "bun:sqlite"
import { createHash } from "node:crypto"
import { existsSync, readFileSync } from "node:fs"
import { EMAIL_ATTACHMENT_EXTENSIONS, extension, parseEmail, safeAttachmentName } from "../../../dochaus/lib/email"
import { diagnosticErrorMetadata } from "../../../dochaus/lib/diagnostic-metadata"
import { draftingTraceReference } from "../../../dochaus/lib/drafting-trace-identity"
import { openDb } from "./db"
import { deleteMatterDocument, DocumentMutationConflictError, RegisteredDocumentNotFoundError } from "./document-delete"
import { documentRegistryIdentity } from "./document-identity"
import { readExtractionQuality } from "./extraction-quality"
import { extractSourceText, ingestDocument, sourceTextRevision, type ExistingSourceChoice, type ExtractedSourceText, type IngestVersion } from "./ingest"
import { detectInjection } from "./sanitize"
import { archiveBudget, ingestZipDocument, ZipTooLargeError, type ArchiveBudget } from "./zip-archive"

// Email attachments (#238). Every attachment of an uploaded email is either
// indexed as its own matter document that names the email (documents.parent_name)
// or listed on the email with what became of it (documents.attachments). None
// is dropped silently.
//
// The upload route ingests the email first, then each attachment it can read,
// each through ingestDocument like an upload of its own: same normalisation,
// injection scan, chunking and checks, each its own mutation. An attachment
// whose text matches a source already in the matter (by rev) links to it
// instead of adding a copy. A forwarded message becomes a child email and its
// own attachments follow, down to MAX_EMAIL_DEPTH. Expansion lives here, not in
// ingestDocument, so a revision or a re-ingest of the email never repeats it.
//
// A zip file is the same kind of container (zip-archive.ts): its files become
// documents that name it, through the same flow and record. Emails and zips nest
// in each other under one depth limit.

// How many levels of attached emails and zip files are ingested below the
// uploaded document. A container nested deeper is listed on its parent, not
// ingested.
export const MAX_EMAIL_DEPTH = 5

// An email or a zip file: a document whose parts become documents of their own.
export function isContainerName(name: string) {
  return /\.(eml|zip)$/i.test(name)
}

// One part of a container as this flow takes it: an email's attachment or a zip
// file's entry. `refused`: a part that is listed but never read (a
// password-protected zip entry, or one whose path climbs out of the archive).
export type ContainerPart = {
  filename: string
  name: string
  mimeType: string
  size: number
  supported: boolean
  tnef?: boolean
  refused?: "encrypted" | "unsafe_path"
  content: () => Uint8Array
}

// What became of one attachment, as stored on the email:
// indexed: held by matter document `document` (indexed from it, or the same
// text already in the matter); unsupported: not a format a matter document can
// take; failed: could not be read or indexed; nested: an email past
// MAX_EMAIL_DEPTH; pending: the email was stored but its attachments were not
// yet processed (an interrupted upload; uploading the email again finishes it);
// no_readable_text: kept as matter document `document` with nothing indexed,
// because the extraction quality gate found no readable text in it. For zip
// entries also: encrypted: password-protected, not added; unsafe_path: its path
// is absolute or climbs out of the archive, not added; too_large: a zip that
// would take the upload past the zip limits (zip-archive.ts), not added.
export type AttachmentOutcome =
  | "indexed"
  | "unsupported"
  | "failed"
  | "nested"
  | "pending"
  | "no_readable_text"
  | "encrypted"
  | "unsafe_path"
  | "too_large"

export type AttachmentEntry = {
  // The name, type and size as the email states them; name "" when it has none.
  name: string
  type: string
  size: number
  outcome: AttachmentOutcome
  document?: string
  // `document`'s rev when it was linked, to tell later whether the lawyer changed it.
  rev?: string
  // The incoming attachment's bytes, independent of later edits to `document`.
  contentHash?: string
  // Outlook's winmail.dat can carry attachments that are not unpacked.
  note?: "may-contain-attachments"
  // The stated name matched the prompt-injection scan.
  flagged?: true
}

export type EmailAttachmentsRecord = { attachments: AttachmentEntry[]; embedded: number }

// What the document list and the upload response show for one attachment. The
// status of a linked document is read from the registry, never stored:
// indexed: extracted from this email; same: the same text was already in the
// matter as `document`; removed: its document has since been removed.
// `degraded` is set on an attachment whose document's extracted text may be
// incomplete or garbled (extraction-quality.ts).
export type AttachmentStatus = "indexed" | "same" | "removed" | Exclude<AttachmentOutcome, "indexed">
export type AttachmentView = Omit<AttachmentEntry, "outcome" | "rev" | "contentHash"> & { status: AttachmentStatus; degraded?: true }

// `parentName` and `depth` are set when the container is itself an attachment
// or a zip entry.
export type ContainerOptions = {
  existing?: ExistingSourceChoice
  parentName?: string | null
  depth?: number
  findings?: ReturnType<typeof detectInjection>
  // What zip files may still expand to in this upload, shared by every zip in
  // it (zip-archive.ts).
  budget?: ArchiveBudget
  // The container's own version: who uploaded it, with what label (#268).
  version?: IngestVersion
}

// doc.haus's own commits of a container it already holds (its attachment
// outcomes, a re-ingest at start), which add no version of their own.
const SYSTEM_VERSION: IngestVersion = { origin: "system", actor: "doc.haus" }

// Ingest an uploaded email and then its attachments.
export async function ingestEmailDocument(matterDir: string, fileName: string, buffer: Buffer, options: ContainerOptions = {}) {
  const parsed = await parseEmail(buffer)
  const parts = parsed.attachments.map(({ content, ...attachment }): ContainerPart => ({ ...attachment, content: () => content }))
  return ingestContainerDocument(matterDir, fileName, buffer, parts, parsed.embedded, options)
}

// Ingest a container, then each of its parts that a matter document can take.
export async function ingestContainerDocument(
  matterDir: string,
  fileName: string,
  buffer: Buffer,
  parts: ContainerPart[],
  embedded: number,
  options: ContainerOptions,
) {
  const depth = options.depth ?? 0
  // One budget for every zip in this upload, however deep.
  const budget = options.budget ?? archiveBudget()
  const extraction = await extractSourceText(fileName, buffer)
  const extracted = { ...extraction, findings: [...extraction.findings, ...(options.findings ?? [])] }
  // Read before this email's own record is replaced, so an upload of the same
  // email again finds the documents its attachments went into.
  const earlier = earlierAttachments(matterDir, fileName, sourceTextRevision(extraction.text))
  const entries = parts.map((part) => initialEntry(part, depth))
  const record = () => JSON.stringify({ attachments: entries, embedded } satisfies EmailAttachmentsRecord)
  const email = await ingestDocument(matterDir, fileName, buffer, {
    existing: options.existing,
    extracted,
    links: { parentName: options.parentName ?? null, attachments: record() },
    version: options.version,
  })
  if (!entries.some((entry) => entry.outcome === "pending")) return { ...email, attachments: emailAttachmentsView(matterDir, email.name, entries) }
  for (const [index, part] of parts.entries()) {
    const entry = entries[index]!
    if (entry.outcome === "pending") Object.assign(entry, await ingestAttachment(matterDir, email.name, part, depth, earlier, budget))
  }
  // Record the outcomes on the email: one more commit of the email itself, which
  // keeps the registry and source-index.json in step through the usual path.
  try {
    await ingestDocument(matterDir, email.name, buffer, {
      existing: "replace",
      requireExistingSource: true,
      extracted,
      links: { attachments: record() },
      version: SYSTEM_VERSION,
    })
  } catch (error) {
    // The email and its attachments are in; it keeps showing them as pending,
    // and uploading it again records them.
    console.error("email attachment outcomes not recorded", { matter_ref: draftingTraceReference(matterDir), ...diagnosticErrorMetadata(error) })
  }
  return { ...email, attachments: emailAttachmentsView(matterDir, email.name, entries) }
}

function initialEntry(attachment: ContainerPart, depth: number): AttachmentEntry {
  const entry: AttachmentEntry = {
    name: attachment.filename,
    type: attachment.mimeType,
    size: attachment.size,
    outcome: attachment.refused
      ?? (!attachment.supported
        ? "unsupported"
        : isContainerName(attachment.name) && depth + 1 > MAX_EMAIL_DEPTH
          ? "nested"
          : "pending"),
  }
  if (attachment.tnef) entry.note = "may-contain-attachments"
  if (detectInjection(attachment.filename).length) entry.flagged = true
  return entry
}

async function ingestAttachment(
  matterDir: string,
  emailName: string,
  attachment: ContainerPart,
  depth: number,
  earlier: AttachmentEntry[],
  budget: ArchiveBudget,
): Promise<Partial<AttachmentEntry>> {
  // A suspicious stated name flags the document it becomes, like suspicious
  // text; the finding has no span because the name is not in the document.
  const nameFindings = detectInjection(attachment.filename).map(({ charStart: _start, charEnd: _end, ...finding }) => ({
    ...finding,
    detail: `Attachment name: ${finding.detail}`,
  }))
  try {
    const bytes = Buffer.from(attachment.content())
    const contentHash = createHash("sha256").update(bytes).digest("hex")
    const before = takeEarlier(matterDir, earlier, attachment, bytes, contentHash)
    if (before) return { ...before, contentHash }
    // An attachment passes the same extraction quality gate as an upload: one
    // with no readable text is kept as a document, with nothing indexed.
    const extraction = await extractSourceText(attachment.name, bytes)
    const same = sameSourceByText(matterDir, extraction, isContainerName(attachment.name) ? bytes : undefined)
    if (same) return { outcome: "indexed", ...same, contentHash }
    const container: ContainerOptions = { existing: "keep-both", parentName: emailName, depth: depth + 1, findings: nameFindings, budget }
    const lower = attachment.name.toLowerCase()
    const child = lower.endsWith(".eml")
      ? await ingestEmailDocument(matterDir, attachment.name, bytes, container)
      : lower.endsWith(".zip")
        ? await ingestZipDocument(matterDir, attachment.name, bytes, container)
        : await ingestDocument(matterDir, attachment.name, bytes, {
          existing: "keep-both",
          extracted: { ...extraction, findings: [...extraction.findings, ...nameFindings] },
          links: { parentName: emailName },
        })
    return { outcome: child.quality.status === "no_readable_text" ? "no_readable_text" : "indexed", document: child.name, rev: child.rev, contentHash }
  } catch (error) {
    if (error instanceof ZipTooLargeError) return { outcome: "too_large" }
    // Stated names are client data: log the failure without them.
    console.warn("email attachment not indexed", { matter_ref: draftingTraceReference(matterDir), ...diagnosticErrorMetadata(error) })
    return { outcome: "failed" }
  }
}

// A source already in the matter with this text, by documents.rev. Text that
// is empty or unreadable says nothing about identity, so it never matches.
// A container's body/listing omits its descendants: its bytes must match too.
function sameSourceByText(matterDir: string, { text, quality }: ExtractedSourceText, containerBytes?: Buffer) {
  if (!text.trim() || quality.status === "no_readable_text") return
  const rev = sourceTextRevision(text)
  const db = openDb(matterDir)
  try {
    const rows = db
      .query("SELECT name, doc_path FROM documents WHERE role = 'source' AND rev = ? ORDER BY created_at, id")
      .all(rev) as { name: string; doc_path: string }[]
    const row = rows.find((row) => !containerBytes || sameFileBytes(row.doc_path, containerBytes))
    if (row) return { document: row.name, rev }
  } finally {
    db.close()
  }
}

// The attachments already recorded for this email, under its name or as an
// earlier upload of the same message, whose documents are still in the matter.
// Uploading the same attachment again reuses it by its original bytes,
// whatever its current text: one the lawyer changed keeps their version, and one with
// no readable text gains no copy. The entry keeps its extraction-time rev, so
// a change still shows.
function earlierAttachments(matterDir: string, fileName: string, emailRev: string) {
  const db = openDb(matterDir)
  try {
    const rows = linkRows(db)
    const sources = new Set(rows.filter((row) => row.role === "source").map((row) => row.name))
    return rows
      .filter((row) => row.role === "source" && (row.name === fileName || row.rev === emailRev))
      .flatMap((row) => readRecord(row.attachments)?.attachments ?? [])
      .filter((entry) => (entry.outcome === "indexed" || entry.outcome === "no_readable_text") && entry.document && sources.has(entry.document))
  } finally {
    db.close()
  }
}

function takeEarlier(matterDir: string, earlier: AttachmentEntry[], attachment: ContainerPart, bytes: Buffer, contentHash: string): Partial<AttachmentEntry> | undefined {
  const index = earlier.findIndex(
    (entry) => entry.name === attachment.filename && entry.type === attachment.mimeType && entry.size === attachment.size
      && (entry.contentHash ? entry.contentHash === contentHash : sameLegacyAttachmentBytes(matterDir, entry, bytes)),
  )
  if (index < 0) return
  const [{ outcome, document, rev }] = earlier.splice(index, 1) as [AttachmentEntry]
  return { outcome, document, ...(rev ? { rev } : {}) }
}

// Legacy records have no original-byte identity. Reuse only when the linked
// file still matches; otherwise keep it and ingest the incoming source beside it.
function sameLegacyAttachmentBytes(matterDir: string, entry: AttachmentEntry, bytes: Buffer) {
  const db = openDb(matterDir)
  try {
    const row = db.query("SELECT doc_path FROM documents WHERE role = 'source' AND name = ?").get(entry.document!) as { doc_path: string } | null
    return !!row && sameFileBytes(row.doc_path, bytes)
  } finally {
    db.close()
  }
}

function sameFileBytes(file: string, bytes: Buffer) {
  try {
    return readFileSync(file).equals(bytes)
  } catch {
    return false
  }
}

type LinkRow = {
  name: string
  doc_path: string
  role: string
  rev: string | null
  parent_name: string | null
  attachments: string | null
  extraction_quality: string | null
}

function readRecord(value: string | null): EmailAttachmentsRecord | undefined {
  if (!value) return
  return JSON.parse(value) as EmailAttachmentsRecord
}

/** Whether the registered email or zip file at `docPath` has attachments still
 * to bring in: its record lists pending entries (an upload stopped before
 * recording them) or it has no record yet. Uploading it again brings them in. */
export function containerAttachmentsPending(matterDir: string, docPath: string) {
  const db = openDb(matterDir)
  try {
    const identity = documentRegistryIdentity(docPath)
    const row = (db.query("SELECT doc_path, attachments FROM documents WHERE role = 'source'").all() as { doc_path: string; attachments: string | null }[])
      .find((candidate) => documentRegistryIdentity(candidate.doc_path) === identity)
    if (!row) return false
    const record = readRecord(row.attachments)
    return !record || record.attachments.some((entry) => entry.outcome === "pending")
  } finally {
    db.close()
  }
}

function linkRows(db: Database) {
  return db.query("SELECT name, doc_path, role, rev, parent_name, attachments, extraction_quality FROM documents").all() as LinkRow[]
}

function attachmentView(entry: AttachmentEntry, emailName: string, rows: ReadonlyMap<string, LinkRow>): AttachmentView {
  const { outcome, rev: _rev, contentHash: _contentHash, ...rest } = entry
  if (outcome !== "indexed" && outcome !== "no_readable_text") return { ...rest, status: outcome }
  const document = entry.document ? rows.get(entry.document) : undefined
  if (!document) return { ...rest, status: "removed" }
  if (outcome === "no_readable_text") return { ...rest, status: outcome }
  const degraded = readExtractionQuality(document.extraction_quality)?.status === "degraded"
  return { ...rest, status: document.parent_name === emailName ? "indexed" : "same", ...(degraded ? { degraded: true as const } : {}) }
}

function emailAttachmentsView(matterDir: string, emailName: string, entries: AttachmentEntry[]) {
  const db = openDb(matterDir)
  try {
    const rows = new Map(linkRows(db).map((row) => [row.name, row]))
    return entries.map((entry) => attachmentView(entry, emailName, rows))
  } finally {
    db.close()
  }
}

// Removing an email removes the documents extracted from its attachments
// (and theirs, for an attached email), except one that another email still in
// the matter links to, or that the lawyer changed since extraction (its rev
// differs or it has redlines). A document the lawyer uploaded is never a child.
export type EmailRemovalPlan = {
  attachments: string[]
  kept: { name: string; reason: "linked" | "changed" }[]
}

export function emailRemovalPlan(db: Database, emailName: string, suppliedRows?: LinkRow[]): EmailRemovalPlan {
  const rows = suppliedRows ?? linkRows(db)
  const redlined = new Set(
    (db.query("SELECT DISTINCT doc_path FROM redlines").all() as { doc_path: string }[]).map((row) =>
      documentRegistryIdentity(row.doc_path),
    ),
  )
  const byName = new Map(rows.map((row) => [row.name, row]))
  const linksTo = (row: LinkRow, name: string) =>
    readRecord(row.attachments)?.attachments.some((item) => item.document === name) ?? false
  // The documents removing the email would remove, skipping those another
  // email keeps (`linked`, grown below) and those the lawyer changed.
  const closure = (linked: ReadonlySet<string>) => {
    const plan: EmailRemovalPlan = { attachments: [], kept: [] }
    const visited = new Set([emailName])
    const queue = [emailName]
    for (let email = queue.shift(); email !== undefined; email = queue.shift()) {
      const record = readRecord(byName.get(email)?.attachments ?? null)
      for (const child of rows.filter((row) => row.parent_name === email && row.role === "source")) {
        if (visited.has(child.name)) continue
        visited.add(child.name)
        const entry = record?.attachments.find((item) => item.document === child.name)
        if (linked.has(child.name)) plan.kept.push({ name: child.name, reason: "linked" })
        else if (!entry?.rev || entry.rev !== child.rev || redlined.has(documentRegistryIdentity(child.doc_path)))
          plan.kept.push({ name: child.name, reason: "changed" })
        else {
          plan.attachments.push(child.name)
          if (isContainerName(child.name)) queue.push(child.name)
        }
      }
    }
    return plan
  }
  // Links are judged against the whole set being removed, so a sibling that
  // goes too never keeps a document. Keeping one can keep a linker, so repeat
  // until nothing more is kept.
  const linked = new Set<string>()
  for (;;) {
    const plan = closure(linked)
    const removing = new Set([emailName, ...plan.attachments])
    const kept = plan.attachments.filter((name) =>
      rows.some((row) => !removing.has(row.name) && row.name !== name && linksTo(row, name)),
    )
    if (!kept.length) return plan
    kept.forEach((name) => linked.add(name))
  }
}

// The document list's view of email links: on an attachment, the email it came
// from and whether that email is still here; on an email, its attachments with
// their status and, when removing it would remove or keep attachments, the plan
// the remove confirmation names.
export function withEmailLinks<T extends { name: string }>(db: Database, documents: T[]) {
  const rows = linkRows(db)
  const byName = new Map(rows.map((row) => [row.name, row]))
  return documents.map((document) => {
    const { parent_name: _parent, attachments: _attachments, ...rest } = document as T & Partial<LinkRow>
    const row = byName.get(document.name)
    const record = readRecord(row?.attachments ?? null)
    const plan = record ? emailRemovalPlan(db, document.name, rows) : undefined
    return {
      ...rest,
      ...(row?.parent_name ? { attachment_of: { name: row.parent_name, removed: !byName.has(row.parent_name) } } : {}),
      ...(record
        ? {
            attachments: record.attachments.map((entry) => attachmentView(entry, document.name, byName)),
            embedded_images: record.embedded,
          }
        : {}),
      ...(plan && (plan.attachments.length || plan.kept.length) ? { removal: plan } : {}),
    }
  })
}

// Remove a document. For an email, unless `keepAttachments`, also remove the
// attachments its removal plan lists, each through the usual delete. The
// email goes first: it is what the lawyer asked to remove.
export async function removeMatterDocument(matterDir: string, name: string, keepAttachments: boolean) {
  const plan = !keepAttachments && isContainerName(name) ? (() => {
    const db = openDb(matterDir)
    try {
      return emailRemovalPlan(db, name)
    } finally {
      db.close()
    }
  })() : undefined
  const result = await deleteMatterDocument(matterDir, name)
  const purgeFailed = result.purge_failed ? [result.purge_failed] : []
  const removed: string[] = []
  // The email is already gone, so an attachment that cannot be removed now is
  // named with the reason, not turned into a failed request; it stays in the
  // matter shown as from a removed email, and can be removed on its own.
  const failed: { name: string; reason: string }[] = []
  for (const attachment of plan?.attachments ?? []) {
    try {
      const deleted = await deleteMatterDocument(matterDir, attachment)
      if (deleted.purge_failed) purgeFailed.push(deleted.purge_failed)
      removed.push(attachment)
    } catch (error) {
      // Already removed meanwhile.
      if (error instanceof RegisteredDocumentNotFoundError) continue
      console.warn("email attachment not removed", { matter_ref: draftingTraceReference(matterDir), ...diagnosticErrorMetadata(error) })
      failed.push({ name: attachment, reason: error instanceof DocumentMutationConflictError ? error.message : "It could not be removed. Try removing it on its own." })
    }
  }
  return { ...result, ...(purgeFailed.length ? { purge_failed: purgeFailed.join(" ") } : {}), attachments: { removed, kept: plan?.kept ?? [], failed } }
}

// Emails ingested before attachments were read have no attachments record, and
// their stored offsets index the old text (no attachments line, forwarded
// messages flattened into the body). Each is re-ingested once, which also
// brings in its attachments. So is an email whose record still has pending
// entries (an upload stopped before it recorded them); its attachments already
// in the matter are reused, not copied. So is an email with an attachment once
// listed as not supported that a matter document can now take (a zip file). A
// zip file left with pending entries is finished the same way. Runs in the
// background after startup; one that fails is logged and retried at the next
// start.
export async function reingestLegacyEmails(matterDir: string) {
  const db = openDb(matterDir)
  const legacy = (() => {
    try {
      return db
        .query("SELECT name, doc_path, attachments FROM documents WHERE role = 'source' AND (lower(name) LIKE '%.eml' OR lower(name) LIKE '%.zip')")
        .all() as { name: string; doc_path: string; attachments: string | null }[]
    } finally {
      db.close()
    }
  })()
  const nowSupported = (entry: AttachmentEntry, index: number) =>
    entry.outcome === "unsupported" &&
    (EMAIL_ATTACHMENT_EXTENSIONS as readonly string[]).includes(extension(safeAttachmentName(entry.name, entry.type, index + 1)))
  let count = 0
  for (const email of legacy) {
    const record = readRecord(email.attachments)
    const zip = email.name.toLowerCase().endsWith(".zip")
    if (record && !record.attachments.some((entry, index) => entry.outcome === "pending" || (!zip && nowSupported(entry, index)))) continue
    if (!existsSync(email.doc_path)) continue
    try {
      const ingest = zip ? ingestZipDocument : ingestEmailDocument
      await ingest(matterDir, email.name, readFileSync(email.doc_path), { existing: "replace", version: SYSTEM_VERSION })
      count++
    } catch (error) {
      console.error("email re-ingest failed", { matter_ref: draftingTraceReference(matterDir), ...diagnosticErrorMetadata(error) })
    }
  }
  return count
}
