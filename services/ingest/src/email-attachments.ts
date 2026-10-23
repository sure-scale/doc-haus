import type { Database } from "bun:sqlite"
import { existsSync, readFileSync } from "node:fs"
import { parseEmail, type EmailAttachment } from "../../../dochaus/lib/email"
import { diagnosticErrorMetadata } from "../../../dochaus/lib/diagnostic-metadata"
import { draftingTraceReference } from "../../../dochaus/lib/drafting-trace-identity"
import { openDb } from "./db"
import { deleteMatterDocument, RegisteredDocumentNotFoundError } from "./document-delete"
import { documentRegistryIdentity } from "./document-identity"
import { extractSourceText, ingestDocument, sourceTextRevision, type ExistingSourceChoice } from "./ingest"
import { detectInjection } from "./sanitize"

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

// How many levels of attached emails are ingested below the uploaded one. A
// message nested deeper is listed on its parent, not ingested.
export const MAX_EMAIL_DEPTH = 5

// What became of one attachment, as stored on the email:
// indexed: held by matter document `document` (indexed from it, or the same
// text already in the matter); unsupported: not a format a matter document can
// take; failed: could not be read or indexed; nested: an email past
// MAX_EMAIL_DEPTH; pending: the email was stored but its attachments were not
// yet processed (an interrupted upload; uploading the email again finishes it).
export type AttachmentOutcome = "indexed" | "unsupported" | "failed" | "nested" | "pending"

export type AttachmentEntry = {
  // The name, type and size as the email states them; name "" when it has none.
  name: string
  type: string
  size: number
  outcome: AttachmentOutcome
  document?: string
  // `document`'s rev when it was linked, to tell later whether the lawyer changed it.
  rev?: string
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
export type AttachmentStatus = "indexed" | "same" | "removed" | Exclude<AttachmentOutcome, "indexed">
export type AttachmentView = Omit<AttachmentEntry, "outcome" | "rev"> & { status: AttachmentStatus }

// Ingest an uploaded email and then its attachments. `parentName` and `depth`
// are set when the email is itself an attachment.
export async function ingestEmailDocument(
  matterDir: string,
  fileName: string,
  buffer: Buffer,
  options: { existing?: ExistingSourceChoice; parentName?: string | null; depth?: number; findings?: ReturnType<typeof detectInjection> } = {},
) {
  const depth = options.depth ?? 0
  const parsed = await parseEmail(buffer)
  const extraction = await extractSourceText(fileName, buffer)
  const extracted = { text: extraction.text, findings: [...extraction.findings, ...(options.findings ?? [])] }
  const entries = parsed.attachments.map((attachment) => initialEntry(attachment, depth))
  const record = () => JSON.stringify({ attachments: entries, embedded: parsed.embedded } satisfies EmailAttachmentsRecord)
  const email = await ingestDocument(matterDir, fileName, buffer, {
    existing: options.existing,
    extracted,
    links: { parentName: options.parentName ?? null, attachments: record() },
  })
  if (!entries.some((entry) => entry.outcome === "pending")) return { ...email, attachments: emailAttachmentsView(matterDir, email.name, entries) }
  for (const [index, attachment] of parsed.attachments.entries()) {
    const entry = entries[index]!
    if (entry.outcome === "pending") Object.assign(entry, await ingestAttachment(matterDir, email.name, attachment, depth))
  }
  // Record the outcomes on the email: one more commit of the email itself, which
  // keeps the registry and source-index.json in step through the usual path.
  try {
    await ingestDocument(matterDir, email.name, buffer, {
      existing: "replace",
      requireExistingSource: true,
      extracted,
      links: { attachments: record() },
    })
  } catch (error) {
    // The email and its attachments are in; it keeps showing them as pending,
    // and uploading it again records them.
    console.error("email attachment outcomes not recorded", { matter_ref: draftingTraceReference(matterDir), ...diagnosticErrorMetadata(error) })
  }
  return { ...email, attachments: emailAttachmentsView(matterDir, email.name, entries) }
}

function initialEntry(attachment: EmailAttachment, depth: number): AttachmentEntry {
  const entry: AttachmentEntry = {
    name: attachment.filename,
    type: attachment.mimeType,
    size: attachment.size,
    outcome: !attachment.supported
      ? "unsupported"
      : attachment.name.toLowerCase().endsWith(".eml") && depth + 1 > MAX_EMAIL_DEPTH
        ? "nested"
        : "pending",
  }
  if (attachment.tnef) entry.note = "may-contain-attachments"
  if (detectInjection(attachment.filename).length) entry.flagged = true
  return entry
}

async function ingestAttachment(
  matterDir: string,
  emailName: string,
  attachment: EmailAttachment,
  depth: number,
): Promise<Partial<AttachmentEntry>> {
  // A suspicious stated name flags the document it becomes, like suspicious
  // text; the finding has no span because the name is not in the document.
  const nameFindings = detectInjection(attachment.filename).map(({ charStart: _start, charEnd: _end, ...finding }) => ({
    ...finding,
    detail: `Attachment name: ${finding.detail}`,
  }))
  const bytes = Buffer.from(attachment.content)
  try {
    if (attachment.name.toLowerCase().endsWith(".eml")) {
      const same = await sameSource(matterDir, attachment.name, bytes)
      if (same) return { outcome: "indexed", ...same }
      const child = await ingestEmailDocument(matterDir, attachment.name, bytes, {
        existing: "keep-both",
        parentName: emailName,
        depth: depth + 1,
        findings: nameFindings,
      })
      return { outcome: "indexed", document: child.name, rev: child.rev }
    }
    const extraction = await extractSourceText(attachment.name, bytes)
    const same = sameSourceByText(matterDir, extraction.text)
    if (same) return { outcome: "indexed", ...same }
    const child = await ingestDocument(matterDir, attachment.name, bytes, {
      existing: "keep-both",
      extracted: { text: extraction.text, findings: [...extraction.findings, ...nameFindings] },
      links: { parentName: emailName },
    })
    return { outcome: "indexed", document: child.name, rev: child.rev }
  } catch (error) {
    // Stated names are client data: log the failure without them.
    console.warn("email attachment not indexed", { matter_ref: draftingTraceReference(matterDir), ...diagnosticErrorMetadata(error) })
    return { outcome: "failed" }
  }
}

async function sameSource(matterDir: string, name: string, bytes: Buffer) {
  return sameSourceByText(matterDir, (await extractSourceText(name, bytes)).text)
}

// A source already in the matter with this text, by documents.rev. Text that is
// empty says nothing about identity, so it never matches.
function sameSourceByText(matterDir: string, text: string) {
  if (!text.trim()) return
  const rev = sourceTextRevision(text)
  const db = openDb(matterDir)
  try {
    const row = db
      .query("SELECT name FROM documents WHERE role = 'source' AND rev = ? ORDER BY created_at, id LIMIT 1")
      .get(rev) as { name: string } | null
    return row ? { document: row.name, rev } : undefined
  } finally {
    db.close()
  }
}

type LinkRow = {
  name: string
  doc_path: string
  role: string
  rev: string | null
  parent_name: string | null
  attachments: string | null
}

function readRecord(value: string | null): EmailAttachmentsRecord | undefined {
  if (!value) return
  return JSON.parse(value) as EmailAttachmentsRecord
}

function linkRows(db: Database) {
  return db.query("SELECT name, doc_path, role, rev, parent_name, attachments FROM documents").all() as LinkRow[]
}

function attachmentView(entry: AttachmentEntry, emailName: string, rows: ReadonlyMap<string, LinkRow>): AttachmentView {
  const { outcome, rev: _rev, ...rest } = entry
  if (outcome !== "indexed") return { ...rest, status: outcome }
  const document = entry.document ? rows.get(entry.document) : undefined
  return { ...rest, status: !document ? "removed" : document.parent_name === emailName ? "indexed" : "same" }
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
  const removing = new Set([emailName])
  const plan: EmailRemovalPlan = { attachments: [], kept: [] }
  const queue = [emailName]
  for (let email = queue.shift(); email !== undefined; email = queue.shift()) {
    const record = readRecord(byName.get(email)?.attachments ?? null)
    for (const child of rows.filter((row) => row.parent_name === email && row.role === "source")) {
      const entry = record?.attachments.find((item) => item.document === child.name)
      const linked = rows.some(
        (row) =>
          !removing.has(row.name) &&
          row.name !== child.name &&
          readRecord(row.attachments)?.attachments.some((item) => item.document === child.name),
      )
      if (linked) plan.kept.push({ name: child.name, reason: "linked" })
      else if (!entry?.rev || entry.rev !== child.rev || redlined.has(documentRegistryIdentity(child.doc_path)))
        plan.kept.push({ name: child.name, reason: "changed" })
      else {
        plan.attachments.push(child.name)
        removing.add(child.name)
        if (child.name.toLowerCase().endsWith(".eml")) queue.push(child.name)
      }
    }
  }
  return plan
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
  const plan = !keepAttachments && name.toLowerCase().endsWith(".eml") ? (() => {
    const db = openDb(matterDir)
    try {
      return emailRemovalPlan(db, name)
    } finally {
      db.close()
    }
  })() : undefined
  const result = await deleteMatterDocument(matterDir, name)
  const removed: string[] = []
  for (const attachment of plan?.attachments ?? []) {
    try {
      await deleteMatterDocument(matterDir, attachment)
      removed.push(attachment)
    } catch (error) {
      // Already removed meanwhile.
      if (!(error instanceof RegisteredDocumentNotFoundError)) throw error
    }
  }
  return { ...result, attachments: { removed, kept: plan?.kept ?? [] } }
}

// Emails ingested before attachments were read have no attachments record, and
// their stored offsets index the old text (no attachments line, forwarded
// messages flattened into the body). Each is re-ingested once, which also
// brings in its attachments. Runs in the background after startup; one that
// fails is logged and retried at the next start.
export async function reingestLegacyEmails(matterDir: string) {
  const db = openDb(matterDir)
  const legacy = (() => {
    try {
      return db
        .query("SELECT name, doc_path FROM documents WHERE role = 'source' AND attachments IS NULL AND lower(name) LIKE '%.eml'")
        .all() as { name: string; doc_path: string }[]
    } finally {
      db.close()
    }
  })()
  let count = 0
  for (const email of legacy) {
    if (!existsSync(email.doc_path)) continue
    try {
      await ingestEmailDocument(matterDir, email.name, readFileSync(email.doc_path), { existing: "replace" })
      count++
    } catch (error) {
      console.error("email re-ingest failed", { matter_ref: draftingTraceReference(matterDir), ...diagnosticErrorMetadata(error) })
    }
  }
  return count
}
