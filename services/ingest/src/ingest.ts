import path from "node:path"
import mammoth from "mammoth"
import PostalMime from "postal-mime"
import { read, utils } from "xlsx"
import { getDocumentRole, insertChunk, listSourceDocuments, openDb, upsertDocument } from "./db"
import { embedChunk, migrateEmbeddings } from "./embed"
import { extractStructure, extractParties, migrateSections, migrateStructure, buildSectionPaths, chunkSection, recordSections, sectionize } from "./structure"
import { pdfToText } from "./pdf"
import { detectInjection, normalizeExtractedText, scanDocxHiddenContent } from "./sanitize"
import { stageSourceChange, writeSourceArtifacts } from "./source-artifacts"
import { withArtifactMutationLocksAsync } from "../../../dochaus/capabilities/artifacts/mutation-lock"
import { chunkText } from "../../../dochaus/lib/chunk-text"
import { diagnosticErrorMetadata } from "../../../dochaus/lib/diagnostic-metadata"
import { draftingTraceReference } from "../../../dochaus/lib/drafting-trace-identity"
import { commitDocumentWriteAsync } from "./atomic-document-write"
import { MatterDocumentIdentityError, resolveMatterDocumentTarget } from "./document-identity"
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

const MATTER_DOCUMENT_EXTENSIONS = new Set([".docx", ".pdf", ".xlsx", ".eml"])

export class SourceDocumentConflictError extends Error {}

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
    requireExistingSource?: boolean
    updateRegistry?: (db: ReturnType<typeof openDb>) => void
  } = {},
) {
  const name = matterDocumentName(fileName)
  if (!name) throw new Error(`Matter documents must use a supported document filename and extension.`)
  return withArtifactMutationLocksAsync(
    matterDir,
    [name, ".dochaus/source-corpus"],
    () => ingestDocumentLocked(matterDir, name, buffer, options),
  )
}

async function ingestDocumentLocked(
  matterDir: string,
  requestedName: string,
  buffer: Buffer,
  options: {
    requireExistingSource?: boolean
    updateRegistry?: (db: ReturnType<typeof openDb>) => void
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
  if (role === "deliverable")
    throw new SourceDocumentConflictError(`Cannot replace generated deliverable ${fileName} with a source document.`)
  if (options.requireExistingSource && role !== "source")
    throw new SourceDocumentConflictError(`${fileName} is no longer a registered source document.`)

  // Untrusted-document defense (issue #17): normalize away invisible Unicode, then
  // scan the visible text and the raw DOCX for injected instructions. Findings are
  // stored on the document and flag overlapping chunks — the text itself is never
  // rewritten, so the lawyer always reviews exactly what the counterparty wrote.
  const extraction = normalizeExtractedText(await extractRawText(fileName, buffer))
  const text = extraction.text
  const findings = [
    ...extraction.findings,
    ...detectInjection(text),
    ...(fileName.toLowerCase().endsWith(".docx") ? scanDocxHiddenContent(buffer) : []),
  ]
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
    await commitIngestedDocument(db, {
      matterDir,
      docPath,
      fileName,
      buffer,
      text,
      party,
      correlation,
      mutation,
      injectionReport: findings.length ? JSON.stringify({ findings }) : null,
      chunks: preparedChunks,
      updateRegistry: options.updateRegistry,
    })
    finalizeDocumentMutationLineage(db, mutation)
    completeDocumentMutation(db, mutation)
  } catch (error) {
    discardDocumentMutation(db, mutation)
    throw error
  } finally {
    db.close()
  }
  return {
    name: fileName,
    docPath,
    sections: sections.length,
    chunks: preparedChunks.length,
    injection: findings.length
      ? { findings, flaggedChunks: preparedChunks.filter((chunk) => chunk.flagged).length }
      : null,
  }
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
  },
) {
  // The source snapshot is staged before legal.db's write lock is taken (see
  // stageSourceChange); the transaction then holds the lock for the registry
  // writes and the snapshot's installation only. The uploaded document's text
  // is rebuilt from its chunks as reconstructText rebuilds it once stored.
  const staged = await stageSourceChange(input.matterDir, db, input.mutation.projectionDirectory, input.docPath, {
    name: input.fileName,
    text: chunkText(
      input.chunks
        .map((chunk) => ({ text: chunk.text, char_start: chunk.charStart }))
        .toSorted((left, right) => left.char_start - right.char_start),
    ).text,
  })
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
      recordDocumentMutationLineageIntent(
        input.mutation,
        (listSourceDocuments(db) as { name: string }[]).map((document) => document.name),
        input.correlation,
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
      return artifacts.sourceDocuments
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

// Pull plain text from a source document. DOCX goes through mammoth; PDF through
// LiteParse/markitdown/unpdf with an OCR fallback for flat scans (see pdf.ts). Always
// normalized (see sanitize.ts) so every consumer — indexing, the document/template
// text routes, and through them the agents' read-document tool — sees the same
// instruction-stripped text that chunk offsets were computed over.
export async function extractDocumentText(fileName: string, buffer: Buffer) {
  return normalizeExtractedText(await extractRawText(fileName, buffer)).text
}

async function extractRawText(fileName: string, buffer: Buffer) {
  if (fileName.toLowerCase().endsWith(".pdf")) return pdfToText(buffer)
  if (fileName.toLowerCase().endsWith(".eml")) return emlToText(buffer)
  if (fileName.toLowerCase().endsWith(".xlsx")) return xlsxToText(buffer)
  return (await mammoth.extractRawText({ buffer })).value
}

// Spreadsheet sources (cap tables, waterfalls, payout models) are zip-based but
// not Word documents, so they go through a workbook reader rather than mammoth.
// Real-world spreadsheets come from many exporters whose XML is not always
// pristine, so we use a tolerant reader. Each sheet becomes a labeled CSV block
// so the tabular relationships (holder, security, shares, percentage) survive as
// text the same normalize/index path can chunk and the agents can read.
function xlsxToText(buffer: Buffer) {
  const workbook = read(buffer, { type: "buffer" })
  return workbook.SheetNames.map((name) => {
    const sheet = workbook.Sheets[name]
    if (!sheet) return ""
    const csv = utils.sheet_to_csv(sheet).trim()
    return csv ? `## ${name}\n${csv}` : ""
  })
    .filter(Boolean)
    .join("\n\n")
}

// Email sources (.eml) are MIME documents, not zip-based Office files, so they go
// through a real RFC822 parser rather than mammoth. The headers that carry legal
// meaning in a chain — who wrote, to whom, when, about what — sit above the body
// so the same normalize/index path treats them as part of the source text.
async function emlToText(buffer: Buffer) {
  const email = await PostalMime.parse(buffer)
  const header = [
    email.subject && `Subject: ${email.subject}`,
    email.from && `From: ${formatAddress(email.from)}`,
    email.to?.length && `To: ${email.to.map(formatAddress).join(", ")}`,
    email.cc?.length && `Cc: ${email.cc.map(formatAddress).join(", ")}`,
    email.date && `Date: ${email.date}`,
  ]
    .filter(Boolean)
    .join("\n")
  const body = email.text ?? (email.html ? email.html.replace(/<[^>]+>/g, " ") : "")
  return header ? `${header}\n\n${body}` : body
}

function formatAddress(addr: { name?: string; address?: string }) {
  if (addr.name && addr.address) return `${addr.name} <${addr.address}>`
  return addr.address ?? addr.name ?? ""
}
