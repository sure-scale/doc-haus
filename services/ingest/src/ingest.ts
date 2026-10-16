import path from "node:path"
import mammoth from "mammoth"
import PostalMime from "postal-mime"
import { read, utils } from "xlsx"
import { getDocumentRole, insertChunk, listSourceDocuments, openDb, upsertDocument } from "./db"
import { embedChunk, migrateEmbeddings } from "./embed"
import { extractStructure, extractParties, migrateStructure, buildSectionPaths, chunkSection } from "./structure"
import { pdfToText } from "./pdf"
import { detectInjection, normalizeExtractedText, scanDocxHiddenContent } from "./sanitize"
import { writeSourceArtifacts } from "./source-artifacts"
import { withArtifactMutationLocksAsync } from "../../../dochaus/capabilities/artifacts/mutation-lock"
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

// A clause line like "7.2 Termination" or "12. Governing Law": the dotted number
// is the hierarchy key, the trailing text is the heading title.
const CLAUSE_RE = /^\s*(\d+(?:\.\d+)*)[.)]?\s+(.+)$/
// A heading line: short, mostly uppercase, no trailing sentence punctuation.
const HEADING_RE = /^[A-Z0-9][A-Z0-9 ,'\-&/]{2,79}$/

// label is the hierarchy key (the dotted clause number, or the heading text when
// there is no number); title is the human heading line; number is the dotted
// numeric when present, which drives the parent/child path nesting.
export type Section = {
  label: string
  title: string
  number: string | null
  text: string
  charStart: number
  charEnd: number
}

export function sectionize(text: string): Section[] {
  const sections: Section[] = []
  let current: Section = { label: "Preamble", title: "Preamble", number: null, text: "", charStart: 0, charEnd: 0 }
  let offset = 0

  for (const line of text.split("\n")) {
    // PDF extractors can emit Markdown headings; strip hashes so "## 7.2 Termination"
    // sections the same as a plain "7.2 Termination" line.
    const trimmed = line.trim().replace(/^#{1,6}\s+/, "")
    const isMdHeading = trimmed !== line.trim()
    const clause = trimmed.match(CLAUSE_RE)
    const isHeading = clause || (isMdHeading ? isMarkdownHeading(trimmed) : trimmed.length > 0 && HEADING_RE.test(trimmed))
    if (isHeading && trimmed.length > 0) {
      if (current.text.trim()) {
        current.charEnd = offset
        sections.push(current)
      }
      const number = clause?.[1] ?? null
      const title = (clause?.[2] ?? trimmed).trim()
      current = { label: number ?? trimmed, title, number, text: line + "\n", charStart: offset, charEnd: 0 }
    } else {
      current.text += line + "\n"
    }
    offset += line.length + 1
  }
  if (current.text.trim()) {
    current.charEnd = offset
    sections.push(current)
  }
  return sections
}

function isMarkdownHeading(text: string) {
  if (CLAUSE_RE.test(text)) return true
  if (/^(appendix|annex|exhibit|schedule|article)\b/i.test(text)) return true
  const words = text.match(/[A-Za-z][A-Za-z'’-]*/g) ?? []
  if (words.length < 2) return false
  if (/^[A-Z]{1,4}\s+\d+[A-Z0-9.-]*$/i.test(text)) return false
  return !words.every((word) => ["a", "an", "and", "between", "for", "of", "or", "the"].includes(word.toLowerCase()))
}

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

  const db = openDb(matterDir)
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
    // Same for structure rows extracted by an older pattern version.
    await migrateEmbeddings(db)
    migrateStructure(db)
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
        transactionDirectory: input.mutation.projectionDirectory,
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
