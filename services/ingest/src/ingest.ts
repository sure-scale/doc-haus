import { existsSync, readdirSync, readFileSync } from "node:fs"
import path from "node:path"
import { strFromU8, unzipSync } from "fflate"
import mammoth from "mammoth"
import { read, utils } from "xlsx"
import { effectiveDocumentLinks, getDocumentRole, insertChunk, listDocuments, listSourceDocuments, openDb, upsertDocument, type DocumentLinks } from "./db"
import { embedChunk, migrateEmbeddings } from "./embed"
import { extractStructure, extractParties, migrateSections, migrateStructure, buildSectionPaths, chunkSection, recordSections, revision, sectionize } from "./structure"
import { readPdf } from "./pdf"
import { assessExtraction, type ExtractionFacts, type ExtractionQuality } from "./extraction-quality"
import { detectInjection, normalizeExtractedText, scanDocxHiddenContent, type InjectionFinding } from "./sanitize"
import { stageSourceChange, writeSourceArtifacts } from "./source-artifacts"
import { canonicalArtifactIdentity, withArtifactMutationLocksAsync } from "../../../dochaus/capabilities/artifacts/mutation-lock"
import { chunkText } from "../../../dochaus/lib/chunk-text"
import { diagnosticErrorMetadata } from "../../../dochaus/lib/diagnostic-metadata"
import { writeListLabels } from "../../../dochaus/lib/docx-list-labels"
import { readDocxMarkup, type DocxMarkup } from "../../../dochaus/lib/docx-markup"
import { markTrackedChanges } from "../../../dochaus/lib/docx-tracked-text"
import { draftingTraceReference } from "../../../dochaus/lib/drafting-trace-identity"
import { parseEmail, type ParsedEmail } from "../../../dochaus/lib/email"
import { parseMsg } from "../../../dochaus/lib/msg"
import { zipToText } from "../../../dochaus/lib/zip"
import { commitDocumentWriteAsync } from "./atomic-document-write"
import { MatterDocumentIdentityError, resolveMatterDocumentTarget } from "./document-identity"
import { captureVersion, stageVersionBytes, trackedChanges, trackedChangesNotice, type CapturedVersion, type VersionOrigin, type VersionWrite } from "./document-versions"
import {
  beginDocumentMutation,
  completeDocumentMutation,
  discardDocumentMutation,
  finalizeDocumentMutationLineage,
  recordDocumentMutationLineageIntent,
  recordDocumentMutationCommit,
  recoverDocumentMutations,
  type ActiveDocumentMutation,
} from "./document-mutation"

const MATTER_DOCUMENT_EXTENSIONS = new Set([".docx", ".pdf", ".xlsx", ".eml", ".msg", ".zip"])

export class SourceDocumentConflictError extends Error {}

export type IngestVersion = Omit<VersionWrite, "origin"> & { origin?: VersionOrigin }

// The upload names a source (or an unregistered file) already in the matter and
// the caller chose neither replace nor keep both. `existing` is its stored name.
export class SourceDocumentExistsError extends SourceDocumentConflictError {
  constructor(readonly existing: string) {
    super(`${existing} is already in this matter. Replace it, or keep both and add the upload under a new name.`)
  }
}

// The upload names a generated deliverable, which a source never replaces; the
// caller can keep both. `existing` is the deliverable's stored name.
export class DeliverableDocumentExistsError extends SourceDocumentConflictError {
  constructor(readonly existing: string) {
    super(`Cannot replace generated deliverable ${existing} with a source document.`)
  }
}

// What an upload does when a source of the same name is already registered, or a
// file of that name is already in the matter folder: refuse (the default, so a
// lawyer's file is never replaced silently), replace it, or keep both by storing
// the upload under the next free name.
export type ExistingSourceChoice = "refuse" | "replace" | "keep-both"

export function matterDocumentName(fileName: string) {
  const name = path.basename(fileName.trim())
  if (!name || name !== fileName.trim() || !MATTER_DOCUMENT_EXTENSIONS.has(path.extname(name).toLowerCase())) return
  return name
}

export async function ingestDocument(
  matterDir: string,
  fileName: string,
  buffer: Buffer,
  options: {
    existing?: ExistingSourceChoice
    requireExistingSource?: boolean
    updateRegistry?: (db: ReturnType<typeof openDb>) => void
    // Text the caller already extracted with extractSourceText, with any extra
    // findings it adds, so a document is not read twice.
    extracted?: ExtractedSourceText
    // The document's email links; a field left out keeps what it had.
    links?: DocumentLinks
    // The version this write records (document-versions.ts). Without an
    // origin the write is an upload: every caller that states none ingests a
    // file the lawyer brought in (an upload, an email and its attachments, a
    // zip's files), and a replace choice over an existing file makes it a replace.
    version?: IngestVersion
  } = {},
) {
  const requested = matterDocumentName(fileName)
  if (!requested) throw new Error(`Matter documents must use a supported document filename and extension.`)
  // Keep both picks a free name before locking it; the locked write still refuses
  // a source that took the name meanwhile, so it never replaces one.
  const { existing = "refuse" } = options
  const name = existing === "keep-both" ? freeDocumentName(matterDir, requested) : requested
  return withArtifactMutationLocksAsync(
    matterDir,
    [name, ".dochaus/source-corpus"],
    () => ingestDocumentLocked(matterDir, name, buffer, { ...options, existing: existing === "keep-both" ? "refuse" : existing }),
  )
}

// The requested name, or the first `Stem (n).ext` from 2 up that no matter file
// or registered document holds. Names compare by artifact identity, so a spelling
// the filesystem resolves to the same file (such as a decomposed Unicode name on
// macOS) is taken, and case-insensitively on every platform.
function freeDocumentName(matterDir: string, name: string) {
  const db = openDb(matterDir)
  const registered = (() => {
    try {
      return (listDocuments(db) as { name: string }[]).map((row) => row.name)
    } finally {
      db.close()
    }
  })()
  const identity = (entry: string) => canonicalArtifactIdentity(matterDir, entry).toLowerCase()
  const taken = new Set([...readdirSync(matterDir), ...registered].map(identity))
  const extension = path.extname(name)
  const stem = name.slice(0, -extension.length)
  for (let n = 1; ; n++) {
    const candidate = n === 1 ? name : `${stem} (${n})${extension}`
    if (!taken.has(identity(candidate))) return candidate
  }
}

async function ingestDocumentLocked(
  matterDir: string,
  requestedName: string,
  buffer: Buffer,
  options: {
    existing: "refuse" | "replace"
    requireExistingSource?: boolean
    updateRegistry?: (db: ReturnType<typeof openDb>) => void
    extracted?: ExtractedSourceText
    links?: DocumentLinks
    version?: IngestVersion
  },
) {
  const recovery = await recoverDocumentMutations(matterDir)
  if (recovery.sourceRolledBack) await writeSourceArtifacts(matterDir)
  const target = sourceDocumentTarget(matterDir, requestedName)
  const fileName = target.name
  const docPath = target.docPath
  const registry = openDb(matterDir)
  const role = getDocumentRole(registry, docPath)
  registry.close()
  if (role === "deliverable") throw new DeliverableDocumentExistsError(fileName)
  if (options.requireExistingSource && role !== "source")
    throw new SourceDocumentConflictError(`${fileName} is no longer a registered source document.`)
  if (options.existing === "refuse" && (role === "source" || existsSync(docPath)))
    throw new SourceDocumentExistsError(fileName)
  const version: VersionWrite = {
    ...options.version,
    origin: options.version?.origin ?? (options.existing === "replace" && existsSync(docPath) ? "replace" : "upload"),
  }

  // Untrusted-document defense (issue #17): normalize away invisible Unicode, then
  // scan the visible text and the raw DOCX for injected instructions. Findings are
  // stored on the document and flag overlapping chunks — the text itself is never
  // rewritten, so the lawyer always reviews exactly what the counterparty wrote.
  const extraction = options.extracted ?? await extractSourceText(fileName, buffer)
  const markup = await sourceMarkup(fileName, buffer)
  const findings = [
    ...extraction.findings,
    ...detectInjection(extraction.text),
    ...(fileName.toLowerCase().endsWith(".docx") ? scanDocxHiddenContent(buffer) : []),
    ...commentFindings(markup),
  ]
  // Extraction quality gate (#239): a file with no readable text is kept and
  // registered, with nothing indexed, so search and review never rest on it.
  const { quality } = extraction
  const text = quality.status === "no_readable_text" ? "" : extraction.text
  const sections = sectionize(text)

  const db = openDb(matterDir, { commit: true })
  let preparedChunks: Array<{
    section: string
    structuralPath: string
    text: string
    charStart: number
    charEnd: number
    embedding: Float32Array
    flagged: boolean
  }> = []
  const party = extractParties(text)
    .map((p) => (p.role ? `${p.name} (${p.role})` : p.name))
    .join("; ")
  const correlation = `ingest-${Date.now().toString(36)}`
  const mutation = beginDocumentMutation(matterDir, fileName, "write", true)
  let captured: CapturedVersion | undefined
  try {
    // Never mix vectors from two embedding models in one matter — re-embed any
    // stale chunks before this document's are written (no-op on current DBs).
    // Same for structure rows extracted by an older pattern version. Sections
    // recorded by an older version or not yet recorded are backfilled too, but
    // nothing reads them across documents, so a failed backfill (rolled back
    // whole, and retried at the next ingest or start) does not block this
    // upload; this document's own sections are still recorded strictly below.
    await migrateEmbeddings(db)
    migrateStructure(db)
    try {
      migrateSections(db)
    } catch (error) {
      console.error("sections backfill failed", { matter_ref: draftingTraceReference(matterDir), ...diagnosticErrorMetadata(error) })
    }
    const paths = buildSectionPaths(fileName, sections)
    preparedChunks = await Promise.all(
      sections.flatMap((section) => {
        const structuralPath = paths.get(section)!
        return chunkSection(section).map(async (chunk) => ({
          section: section.label,
          structuralPath,
          text: chunk.text,
          charStart: chunk.charStart,
          charEnd: chunk.charEnd,
          embedding: await embedChunk(fileName, structuralPath, chunk.text),
          flagged: findings.some(
            (finding) =>
              finding.charStart !== undefined &&
              finding.charEnd !== undefined &&
              finding.charStart < chunk.charEnd &&
              finding.charEnd > chunk.charStart,
          ),
        }))
      }),
    )
    const artifacts = await commitIngestedDocument(db, {
      matterDir,
      docPath,
      fileName,
      buffer,
      text,
      party,
      correlation,
      mutation,
      injectionReport: findings.length ? JSON.stringify({ findings }) : null,
      extractionQuality: JSON.stringify(quality),
      chunks: preparedChunks,
      updateRegistry: options.updateRegistry,
      links: options.links ?? {},
      version,
    })
    finalizeDocumentMutationLineage(db, mutation, artifacts.files)
    completeDocumentMutation(db, mutation)
    captured = artifacts.version
  } catch (error) {
    discardDocumentMutation(db, mutation)
    throw error
  } finally {
    db.close()
  }
  return {
    name: fileName,
    docPath,
    // documents.rev as recorded for this ingest.
    rev: revision(canonicalText(preparedChunks)),
    sections: sections.length,
    chunks: preparedChunks.length,
    injection: findings.length
      ? { findings, flaggedChunks: preparedChunks.filter((chunk) => chunk.flagged).length }
      : null,
    quality,
    // The version this ingest recorded; created is false when the file already
    // held these bytes.
    version: captured ?? null,
    ...versionNotice(fileName, buffer, captured, markup),
  }
}

/** A saved Word file's pending tracked changes, with the notice of how doc.haus
 * edits meet them (document-versions.ts); nothing for a file without. With a
 * source's `markup`, the changes are counted as the viewer lists them, where
 * the lawyer reviews them and doc.haus edits go beside them; otherwise by their
 * tags, resolved in Word before doc.haus edits the file. */
export function versionNotice(name: string, bytes: Uint8Array, version: CapturedVersion | undefined, markup?: DocxMarkup) {
  const tracked = markup
    ? { count: markup.revisions.length, authors: markup.authors.filter((author) => author.revisions && author.name).map((author) => author.name) }
    : trackedChanges(name, bytes)
  return tracked.count && version ? { tracked_changes: tracked, notice: trackedChangesNotice(version.number, tracked, markup ? "viewer" : "word") } : {}
}

/** A Word source's tracked changes and comments as the viewer lists them
 * (dochaus/lib/docx-markup.ts), read only when it has either. Undefined when
 * it has none or they cannot be read; the file then ingests as before and its
 * notice says to resolve its changes in Word. */
async function sourceMarkup(fileName: string, buffer: Buffer): Promise<DocxMarkup | undefined> {
  if (!fileName.toLowerCase().endsWith(".docx")) return
  try {
    if (!trackedChanges(fileName, buffer).count && !hasCommentsPart(buffer)) return
    return await readDocxMarkup(buffer)
  } catch (error) {
    console.warn("[ingest] tracked changes and comments not read", diagnosticErrorMetadata(error))
  }
}

/** Whether the Word body has a comments relationship; its part is read only
 * when it does. */
function hasCommentsPart(buffer: Buffer) {
  const relationships = unzipSync(buffer, { filter: (file) => file.name === "word/_rels/document.xml.rels" })["word/_rels/document.xml.rels"]
  return relationships !== undefined && /\/relationships\/comments["']/.test(strFromU8(relationships))
}

/** The injection findings in a source's comments, as document-level findings
 * (comments are not indexed, so they have no offsets). The detail is fixed
 * text: read-document prints it outside the untrusted frame, so it never names
 * the comment's author. A comment's quote is body text, which ingest scans
 * with offsets. */
function commentFindings(markup: DocxMarkup | undefined): InjectionFinding[] {
  return (markup?.comments ?? []).flatMap((comment) => comment.findings.filter((finding) => finding.field !== "quote").map((finding) => ({
    rule: finding.rule,
    detail: `${finding.detail}, in a comment`,
    ...(finding.snippet === undefined ? {} : { snippet: finding.snippet }),
  })))
}

function sourceDocumentTarget(matterDir: string, fileName: string) {
  try {
    return resolveMatterDocumentTarget(matterDir, fileName)
  } catch (error) {
    if (error instanceof MatterDocumentIdentityError) throw new SourceDocumentConflictError(error.message)
    throw error
  }
}

async function commitIngestedDocument(
  db: ReturnType<typeof openDb>,
  input: {
    matterDir: string
    docPath: string
    fileName: string
    buffer: Buffer
    text: string
    party: string
    correlation: string
    mutation: ActiveDocumentMutation
    injectionReport: string | null
    extractionQuality: string
    chunks: Array<{
      section: string
      structuralPath: string
      text: string
      charStart: number
      charEnd: number
      embedding: Float32Array
      flagged: boolean
    }>
    updateRegistry?: (db: ReturnType<typeof openDb>) => void
    links: DocumentLinks
    version: VersionWrite
  },
) {
  const links = effectiveDocumentLinks(db, input.fileName, input.links)
  // The source snapshot is staged before legal.db's write lock is taken (see
  // stageSourceChange); the transaction then holds the lock for the registry
  // writes and the snapshot's installation only. The uploaded document's text
  // is rebuilt from its chunks as reconstructText rebuilds it once stored.
  const staged = await stageSourceChange(input.matterDir, db, input.mutation.projectionDirectory, input.docPath, {
    name: input.fileName,
    text: canonicalText(input.chunks),
    ...links,
  })
  // The version's bytes, likewise stored before the lock (stageVersionBytes).
  const stage = stageVersionBytes(input.matterDir, input.fileName, input.buffer, input.version.origin)
  return commitDocumentWriteAsync(input.docPath, input.buffer, async () => {
    db.run("BEGIN IMMEDIATE")
    let settled = false
    try {
      const documentId = upsertDocument(
        db,
        input.docPath,
        input.fileName,
        Date.now(),
        input.injectionReport,
        { parentName: links.parent_name, attachments: links.attachments },
        input.extractionQuality,
      )
      input.chunks.forEach((chunk, chunkIndex) =>
        insertChunk(db, {
          documentId,
          docPath: input.docPath,
          docName: input.fileName,
          section: chunk.section,
          structuralPath: chunk.structuralPath,
          party: input.party,
          chunkIndex,
          text: chunk.text,
          charStart: chunk.charStart,
          charEnd: chunk.charEnd,
          embedding: chunk.embedding,
          flagged: chunk.flagged,
        }),
      )
      recordSections(db, input.docPath, input.fileName)
      extractStructure(db, input.docPath, input.fileName, input.text)
      input.updateRegistry?.(db)
      // The bytes this write replaced are its backup until it settles.
      const backup = input.mutation.documentBackup
      const version = captureVersion(db, input.matterDir, {
        name: input.fileName,
        bytes: input.buffer,
        prior: existsSync(backup) ? readFileSync(backup) : null,
        version: input.version,
        stage,
      })
      recordDocumentMutationLineageIntent(
        input.mutation,
        (listSourceDocuments(db) as { name: string }[]).map((document) => document.name),
        input.correlation,
        input.version.origin,
      )
      const artifacts = await writeSourceArtifacts(input.matterDir, input.correlation, {
        db,
        stampLineage: false,
        staged,
        settle: () => {
          recordDocumentMutationCommit(db, input.mutation)
          db.run("COMMIT")
          settled = true
        },
      })
      return { ...artifacts, version }
    } catch (error) {
      if (!settled) {
        try {
          db.run("ROLLBACK")
        } catch {
          // Preserve the original failure. SQLite may already have rolled the
          // transaction back while surfacing a failed commit.
        }
      }
      throw error
    }
  }, { backupPath: input.mutation.documentBackup })
}

// A document's canonical text as rebuilt from its chunks (reconstructText): the
// text chunk offsets, section spans and documents.rev index.
function canonicalText(chunks: ReadonlyArray<{ text: string; charStart: number }>) {
  return chunkText(
    chunks
      .map((chunk) => ({ text: chunk.text, char_start: chunk.charStart }))
      .toSorted((left, right) => left.char_start - right.char_start),
  ).text
}

// The documents.rev an ingest of this extracted text records, without indexing
// it: the same sections and chunks ingest cuts, rebuilt the same way.
export function sourceTextRevision(text: string) {
  return revision(canonicalText(sectionize(text).flatMap((section) => chunkSection(section))))
}

// `quality` is the extraction quality gate's result for the text
// (extraction-quality.ts).
export type ExtractedSourceText = { text: string; findings: InjectionFinding[]; quality: ExtractionQuality }

// A source document's normalized text, the findings normalization made and the
// text's quality, as ingest indexes it.
export async function extractSourceText(fileName: string, buffer: Buffer): Promise<ExtractedSourceText> {
  const raw = await extractRawText(fileName, buffer)
  return extractedText(raw.text, raw.facts)
}

// The same for an email already parsed, so one parse serves its text and its
// attachments.
export function extractedEmailText(email: ParsedEmail): ExtractedSourceText {
  return extractedText(email.text, emailFacts(email))
}

// An email (.eml or .msg) parsed by its format's one extractor.
export function parseEmailFile(fileName: string, buffer: Buffer) {
  return fileName.toLowerCase().endsWith(".msg") ? parseMsg(buffer) : parseEmail(buffer)
}

function extractedText(rawText: string, rawFacts: ExtractionFacts): ExtractedSourceText {
  const { text, findings } = normalizeExtractedText(rawText)
  const facts = rawFacts.format === "email" ? { ...rawFacts, body: normalizeExtractedText(rawFacts.body).text } : rawFacts
  return { text, findings, quality: assessExtraction(text, facts) }
}

function emailFacts(email: ParsedEmail): ExtractionFacts {
  return { format: "email", body: email.body, attachments: email.attachments.length, ...(email.bodyUnreadable ? { bodyUnreadable: true as const } : {}) }
}

// Pull plain text from a source document. DOCX goes through mammoth, with Word's
// list numbers written in first (see dochaus/lib/docx-list-labels.ts) and then
// pending tracked changes labelled (dochaus/lib/docx-tracked-text.ts); PDF through
// LiteParse/markitdown/unpdf with an OCR fallback for flat scans (see pdf.ts). Always
// normalized (see sanitize.ts) so every consumer — indexing, the document/template
// text routes, and through them the agents' read-document tool — sees the same
// instruction-stripped text that chunk offsets were computed over.
export async function extractDocumentText(fileName: string, buffer: Buffer) {
  return normalizeExtractedText((await extractRawText(fileName, buffer)).text).text
}

async function extractRawText(fileName: string, buffer: Buffer): Promise<{ text: string; facts: ExtractionFacts }> {
  const lower = fileName.toLowerCase()
  if (lower.endsWith(".pdf")) {
    const { text, pages } = await readPdf(buffer)
    return { text, facts: { format: "pdf", pages } }
  }
  if (lower.endsWith(".eml") || lower.endsWith(".msg")) {
    const email = await parseEmailFile(lower, buffer)
    return { text: email.text, facts: emailFacts(email) }
  }
  if (lower.endsWith(".xlsx")) return xlsxToText(buffer)
  if (lower.endsWith(".zip")) return { text: zipToText(buffer), facts: { format: "archive" } }
  return { text: (await mammoth.extractRawText({ buffer: markTrackedChanges(await writeListLabels(buffer)) })).value, facts: { format: "docx" } }
}

// Spreadsheet sources (cap tables, waterfalls, payout models) are zip-based but
// not Word documents, so they go through a workbook reader rather than mammoth.
// Real-world spreadsheets come from many exporters whose XML is not always
// pristine, so we use a tolerant reader. Each sheet becomes a labeled CSV block
// so the tabular relationships (holder, security, shares, percentage) survive as
// text the same normalize/index path can chunk and the agents can read. A
// formula saved without a calculated value (a workbook written by a program and
// never opened in a spreadsheet) reads as its formula, so those are counted for
// the quality gate. One saved with an empty value element (as openpyxl writes
// it) is dropped by the reader unless stub cells are kept; it then reads as a
// stub ("z") holding the formula, and is counted the same way.
function xlsxToText(buffer: Buffer): { text: string; facts: ExtractionFacts } {
  const workbook = read(buffer, { type: "buffer", sheetStubs: true })
  let cells = 0
  let formulasWithoutValue = 0
  const text = workbook.SheetNames.map((name) => {
    const sheet = workbook.Sheets[name]
    if (!sheet) return ""
    for (const [address, cell] of Object.entries(sheet)) {
      if (address.startsWith("!")) continue
      const { t, v, f } = cell as { t?: string; v?: unknown; f?: string }
      const noValue = v === undefined || t === "z"
      if (noValue && !f) continue
      cells++
      if (noValue) formulasWithoutValue++
    }
    const csv = utils.sheet_to_csv(sheet).trim()
    return csv ? `## ${name}\n${csv}` : ""
  })
    .filter(Boolean)
    .join("\n\n")
  return { text, facts: { format: "workbook", cells, formulasWithoutValue } }
}
