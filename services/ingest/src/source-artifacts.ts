import type { Database } from "bun:sqlite"
import { createHash, randomUUID } from "node:crypto"
import { existsSync, mkdirSync, readFileSync, renameSync, rmSync, writeFileSync } from "node:fs"
import path from "node:path"
import { isDeepStrictEqual } from "node:util"
import { listSourceDocuments, openDb } from "./db"
import { documentRegistryIdentity } from "./document-identity"
import {
  extractIdentifiers,
  IDENTIFIERS_SCHEMA_VERSION,
  renderIdentifiersMarkdown,
  type IdentifierChunk,
  type MatterIdentifiers,
} from "./identifiers"
import { stampIngestLineage, stampSourceRebuildLineage } from "./lineage"
import { reconstructText, revision } from "./structure"

const MAX_SOURCE_CHUNK_CHARS = 4500

type SourceRecord = {
  document: string
  doc_path: string
  extracted_text: string
  characters: number
  // The reconstructed text trimmed: sources/*.txt holds it plus one newline, and
  // source-map.json's start_char/end_char and sha256 index and hash it (UTF-16),
  // unlike chunk offsets, section spans and documents.rev, which use it untrimmed.
  text: string
}

type SourceChunk = {
  id: string
  index: number
  heading_path: string[]
  text_path: string
  start_line: number
  end_line: number
  start_char: number
  end_char: number
  characters: number
  sha256: string
  preview: string
}

const SOURCE_ARTIFACT_PATHS = [
  "source-index.json",
  "source-map.md",
  "source-map.json",
  "source-ledger.json",
  "matter-identifiers.json",
  "matter-identifiers.md",
  "sources",
  "source-chunks",
] as const

// A source document as the registry lists it, with its rev (see db.ts).
type SourceRevision = { doc_path: string; name: string; rev: string | null }

// A snapshot staged in `transaction` before the commit that installs it, and
// the registry sources (listSourceDocuments order) it was staged from.
export type StagedSourceArtifacts = { transaction: string; sources: SourceRevision[]; chunks: number }

// `rebuild` re-derives the projection from sources already ingested (the
// startup self-heal): lineage then records only the artifacts whose bytes
// changed, as a rebuild rather than an upload. `staged`: a snapshot staged
// before this commit (stageSourceChange); it is installed when the registry
// now lists exactly the sources it was staged from, and otherwise staged
// again here from the registry.
export async function writeSourceArtifacts(
  matterDir: string,
  correlation?: string,
  options: {
    db?: Database
    settle?: () => void
    stampLineage?: boolean
    rebuild?: boolean
    staged?: StagedSourceArtifacts
  } = {},
) {
  const db = options.db ?? openDb(matterDir)
  const staged = options.staged
  const current = staged && sameSources(sourceRevisions(db), staged.sources) ? staged : undefined
  const sources = current ? [] : (listSourceDocuments(db) as { doc_path: string; name: string }[])
    .map((doc) => sourceRecord(doc.doc_path, doc.name, reconstructText(db, doc.doc_path)))
  if (!options.db) db.close()

  const transaction = staged?.transaction ?? path.join(
    matterDir,
    ".dochaus",
    "source-artifact-transactions",
    `${process.pid}-${randomUUID()}`,
  )
  let snapshot: { documents: string[]; chunks: number }
  try {
    snapshot = current
      ? { documents: current.sources.map((source) => source.name), chunks: current.chunks }
      : await stageSourceArtifacts(matterDir, transaction, sources)
    settleSourceArtifactSnapshot(matterDir, transaction, options.settle)
  } catch (error) {
    if (
      !existsSync(path.join(transaction, "prior-artifacts.json")) &&
      !existsSync(path.join(transaction, "backup"))
    ) rmSync(transaction, { recursive: true, force: true })
    throw error
  }
  if (options.stampLineage !== false) {
    const correlationId = correlation ?? `${options.rebuild ? "rebuild" : "ingest"}-${Date.now().toString(36)}`
    if (options.rebuild) stampSourceRebuildLineage(matterDir, snapshot.documents, correlationId)
    else stampIngestLineage(matterDir, snapshot.documents, correlationId)
  }
  return { documents: snapshot.documents.length, chunks: snapshot.chunks, sourceDocuments: snapshot.documents }
}

// Stage, before a source upload's or delete's write transaction, the snapshot
// its commit will install: every registered source other than `docPath` as it
// is now, then for an upload `uploaded`, the document's name and the untrimmed
// canonical text its chunks rebuild. Extracting the whole matter's identifiers
// is most of a commit's time on a large matter, and holding legal.db's write
// lock through it would make every engine write wait or fail. The caller holds
// the matter's source-corpus lock and passes the result to
// writeSourceArtifacts inside the transaction.
export async function stageSourceChange(
  matterDir: string,
  db: Database,
  transaction: string,
  docPath: string,
  uploaded?: { name: string; text: string },
): Promise<StagedSourceArtifacts> {
  // upsertDocument and deleteDocument both remove every row of the document's
  // identity; an upload's new row is listed last (it is created now).
  const identity = documentRegistryIdentity(docPath)
  const rows = [
    ...sourceRevisions(db)
      .filter((row) => documentRegistryIdentity(row.doc_path) !== identity)
      .map((row) => ({ doc_path: row.doc_path, name: row.name, text: reconstructText(db, row.doc_path) })),
    ...(uploaded ? [{ doc_path: docPath, ...uploaded }] : []),
  ]
  const { chunks } = await stageSourceArtifacts(
    matterDir,
    transaction,
    rows.map((row) => sourceRecord(row.doc_path, row.name, row.text)),
  )
  return {
    transaction,
    sources: rows.map((row) => ({ doc_path: row.doc_path, name: row.name, rev: revision(row.text) })),
    chunks,
  }
}

function sourceRevisions(db: Database) {
  return db
    .query("SELECT doc_path, name, rev FROM documents WHERE role = 'source' ORDER BY created_at")
    .all() as SourceRevision[]
}

function sameSources(left: SourceRevision[], right: SourceRevision[]) {
  return left.length === right.length && left.every((source, index) => {
    const other = right[index]!
    return source.doc_path === other.doc_path && source.name === other.name && source.rev === other.rev
  })
}

// `text`: the document's untrimmed canonical text (see SourceRecord).
function sourceRecord(docPath: string, name: string, text: string): SourceRecord {
  const trimmed = text.trim()
  return {
    document: name,
    doc_path: docPath,
    extracted_text: `sources/${sourceSlug(name)}.txt`,
    text: trimmed,
    characters: trimmed.length,
  }
}

// Write the snapshot of `sources` to `<transaction>/staged`, replacing anything
// staged there before. No sources stage an empty snapshot, which removes the
// projection when installed.
async function stageSourceArtifacts(matterDir: string, transaction: string, sources: SourceRecord[]) {
  const staged = path.join(transaction, "staged")
  rmSync(staged, { recursive: true, force: true })
  mkdirSync(staged, { recursive: true })
  const documents = sources.map((source) => source.document)
  if (!sources.length) return { documents, chunks: 0 }
  const previousIdentifiers = await readMatterIdentifiers(matterDir)
  const sourceMap = buildSourceMap(sources)
  mkdirSync(path.join(staged, "sources"), { recursive: true })
  mkdirSync(path.join(staged, "source-chunks"), { recursive: true })
  await Promise.all(sources.map((source) => Bun.write(path.join(staged, source.extracted_text), `${source.text}\n`)))
  const chunkFiles = sourceChunkFiles(sources, sourceMap)
  await writeSourceChunks(staged, chunkFiles)
  await writeMatterIdentifiers(staged, chunkFiles, previousIdentifiers)
  await Bun.write(path.join(staged, "source-ledger.json"), `${JSON.stringify(sourceMap, null, 2)}\n`)
  await Bun.write(path.join(staged, "source-map.json"), `${JSON.stringify(sourceMap, null, 2)}\n`)
  await Bun.write(path.join(staged, "source-map.md"), sourceMapMarkdown(sourceMap))
  await Bun.write(
    path.join(staged, "source-index.json"),
    `${JSON.stringify(
      {
        source_ledger: "source-ledger.json",
        source_map: "source-map.json",
        source_map_markdown: "source-map.md",
        source_chunk_directory: "source-chunks",
        sources: sources.map((source) => ({
          document: source.document,
          doc_path: source.doc_path,
          extracted_text: source.extracted_text,
          characters: source.characters,
        })),
      },
      null,
      2,
    )}\n`,
  )
  return { documents, chunks: sourceMap.documents.reduce((sum, source) => sum + source.chunks.length, 0) }
}

// Startup heal for one matter (ingest server boot). Returns what it did.
export async function healSourceArtifacts(matterDir: string, sourceRolledBack: boolean) {
  if (!sourceRolledBack && sourceArtifactsMatchRegistry(matterDir)) return "current" as const
  // A rolled-back mutation also restored legal.db, so it is stamped as an ingest like the other recovery
  // paths. Only a registry mismatch is a rebuild: it stamps just the files whose bytes changed.
  await writeSourceArtifacts(matterDir, undefined, sourceRolledBack ? {} : { rebuild: true })
  return sourceRolledBack ? "recovered" as const : "rebuilt" as const
}

export function sourceArtifactsMatchRegistry(matterDir: string, suppliedDb?: Database) {
  const db = suppliedDb ?? openDb(matterDir)
  let sources: SourceRecord[]
  try {
    sources = (listSourceDocuments(db) as { doc_path: string; name: string }[])
      .map((document) => sourceRecord(document.doc_path, document.name, reconstructText(db, document.doc_path)))
  } finally {
    if (!suppliedDb) db.close()
  }
  if (!sources.length) return SOURCE_ARTIFACT_PATHS.every((relative) => !existsSync(path.join(matterDir, relative)))
  if (SOURCE_ARTIFACT_PATHS.slice(0, 6).some((relative) => !existsSync(path.join(matterDir, relative)))) return false
  try {
    const expectedMap = buildSourceMap(sources)
    const expectedIndex = {
      source_ledger: "source-ledger.json",
      source_map: "source-map.json",
      source_map_markdown: "source-map.md",
      source_chunk_directory: "source-chunks",
      sources: sources.map((source) => ({
        document: source.document,
        doc_path: source.doc_path,
        extracted_text: source.extracted_text,
        characters: source.characters,
      })),
    }
    if (JSON.stringify(JSON.parse(readFileSync(path.join(matterDir, "source-index.json"), "utf8"))) !== JSON.stringify(expectedIndex)) return false
    if (JSON.stringify(JSON.parse(readFileSync(path.join(matterDir, "source-map.json"), "utf8"))) !== JSON.stringify(expectedMap)) return false
    if (JSON.stringify(JSON.parse(readFileSync(path.join(matterDir, "source-ledger.json"), "utf8"))) !== JSON.stringify(expectedMap)) return false
    if (readFileSync(path.join(matterDir, "source-map.md"), "utf8") !== sourceMapMarkdown(expectedMap)) return false
    if (sources.some((source) => readFileSync(path.join(matterDir, source.extracted_text), "utf8") !== `${source.text}\n`)) return false
    if (sourceChunkFiles(sources, expectedMap).some((chunk) =>
      !existsSync(path.join(matterDir, chunk.text_path)) ||
      readFileSync(path.join(matterDir, chunk.text_path), "utf8") !== chunk.file.text)) return false
    // The identifiers are not extracted again. writeSourceArtifacts extracts
    // them from exactly the chunk files checked above and installs them in the
    // same snapshot as the source map, whose sha256 of every document and chunk
    // was just matched against the current text; a partly installed snapshot
    // lacks a file and fails above. So an unchanged source corpus has the
    // identifiers it would get again, and re-running compromise over the whole
    // matter on every start (most of a large matter's start time) would only
    // repeat them. A change to the extraction reaches existing matters through
    // IDENTIFIERS_SCHEMA_VERSION: a file stored under another version does not
    // decode below, so the next start rebuilds it once, and it matches after.
    const identifiers = decodeMatterIdentifiers(
      JSON.parse(readFileSync(path.join(matterDir, "matter-identifiers.json"), "utf8")),
    )
    if (!identifiers || identifiers.documents !== new Set(sources.map((source) => source.document)).size) return false
    return readFileSync(path.join(matterDir, "matter-identifiers.md"), "utf8") ===
      `${renderIdentifiersMarkdown(identifiers)}\n`
  } catch {
    return false
  }
}

function settleSourceArtifactSnapshot(matterDir: string, transaction: string, settle?: () => void) {
  const installation = installSourceArtifactSnapshot(matterDir, transaction)
  try {
    settle?.()
  } catch (error) {
    installation.rollback()
    throw error
  }
  installation.commit()
}

function installSourceArtifactSnapshot(matterDir: string, transaction: string) {
  const staged = path.join(transaction, "staged")
  const backup = path.join(transaction, "backup")
  const prior = SOURCE_ARTIFACT_PATHS.filter((relative) => existsSync(path.join(matterDir, relative)))
  mkdirSync(backup, { recursive: true })
  writeFileSync(path.join(transaction, "prior-artifacts.json"), `${JSON.stringify(prior)}\n`, { flag: "wx" })
  try {
    SOURCE_ARTIFACT_PATHS.forEach((relative) => {
      const current = path.join(matterDir, relative)
      if (!existsSync(current)) return
      renameSync(current, path.join(backup, relative))
    })
    SOURCE_ARTIFACT_PATHS.forEach((relative) => {
      const candidate = path.join(staged, relative)
      if (!existsSync(candidate)) return
      renameSync(candidate, path.join(matterDir, relative))
    })
  } catch (error) {
    recoverSourceArtifactTransaction(matterDir, transaction, false)
    throw error
  }
  return {
    commit() {
      try {
        rmSync(transaction, { recursive: true, force: true })
      } catch {
        // The installed snapshot is authoritative. A leftover backup directory
        // is harmless and can be removed by maintenance without rolling back a
        // source corpus whose database settlement already succeeded.
      }
    },
    rollback() {
      recoverSourceArtifactTransaction(matterDir, transaction, false)
    },
  }
}

export function recoverSourceArtifactTransaction(
  matterDir: string,
  transaction: string,
  committed: boolean,
) {
  if (!existsSync(transaction)) return
  if (committed) {
    rmSync(transaction, { recursive: true, force: true })
    return
  }
  const manifest = path.join(transaction, "prior-artifacts.json")
  if (!existsSync(manifest)) {
    if (existsSync(path.join(transaction, "backup")))
      throw new Error(`Cannot safely recover source artifact transaction ${path.basename(transaction)}.`)
    rmSync(transaction, { recursive: true, force: true })
    return
  }
  const decoded = JSON.parse(readFileSync(manifest, "utf8")) as unknown
  const allowed = new Set<string>(SOURCE_ARTIFACT_PATHS)
  if (
    !Array.isArray(decoded) ||
    decoded.some((relative) => typeof relative !== "string" || !allowed.has(relative))
  ) throw new Error(`Cannot safely recover source artifact transaction ${path.basename(transaction)}.`)
  const prior = new Set(decoded as string[])
  SOURCE_ARTIFACT_PATHS.forEach((relative) => {
    const live = path.join(matterDir, relative)
    const saved = path.join(transaction, "backup", relative)
    if (!prior.has(relative)) {
      rmSync(live, { recursive: true, force: true })
      return
    }
    if (!existsSync(saved)) return
    rmSync(live, { recursive: true, force: true })
    renameSync(saved, live)
  })
  rmSync(transaction, { recursive: true, force: true })
}

function buildSourceMap(sourceRows: SourceRecord[]) {
  return {
    schema_version: 1,
    version: 1,
    chunking: {
      method: "heading-and-paragraph-ranges",
      max_characters: MAX_SOURCE_CHUNK_CHARS,
    },
    documents: sourceRows.map((source) => ({
      document: source.document,
      doc_path: source.doc_path,
      extracted_text: source.extracted_text,
      characters: source.text.length,
      sha256: sha256Text(source.text),
      chunks: chunkSource(source),
    })),
  }
}

// Each chunk's source slice and the file written for it. Chunk files are read
// by drafting agents through the generic `read` tool, which truncates any
// single line beyond 2000 characters mid-clause. Legal prose often emits an
// entire clause or recital as one line, so the file is soft-wrapped at spaces
// (see wrapForReadability). The authoritative offsets live in source-map.json
// (char ranges into the trimmed reconstructed text, see SourceRecord), never in
// these read-only files; identifier values are read from the unwrapped slice.
function sourceChunkFiles(sourceRows: SourceRecord[], sourceMap: ReturnType<typeof buildSourceMap>) {
  const textByDocument = new Map(sourceRows.map((source) => [source.document, source.text]))
  return sourceMap.documents.flatMap((document) => {
    const text = textByDocument.get(document.document)
    if (text === undefined) throw new Error(`Missing source text for ${document.document}.`)
    return document.chunks.map((chunk) => {
      const slice = text.slice(chunk.start_char, chunk.end_char)
      return { id: chunk.id, document: document.document, text_path: chunk.text_path, text: slice, file: wrapForReadability(slice) }
    })
  })
}

async function writeSourceChunks(matterDir: string, chunks: ReturnType<typeof sourceChunkFiles>) {
  await Promise.all(chunks.map((chunk) => Bun.write(path.join(matterDir, chunk.text_path), chunk.file.text)))
}

async function writeMatterIdentifiers(
  matterDir: string,
  chunks: IdentifierChunk[],
  previous: MatterIdentifiers | undefined,
) {
  const extracted = await extractIdentifiers(chunks, previous)
  const data = previous && sameIdentifierProjection(previous, extracted)
    ? { ...extracted, generated_at: previous.generated_at }
    : extracted
  await Bun.write(path.join(matterDir, "matter-identifiers.json"), `${JSON.stringify(data, null, 2)}\n`)
  await Bun.write(path.join(matterDir, "matter-identifiers.md"), `${renderIdentifiersMarkdown(data)}\n`)
}

async function readMatterIdentifiers(matterDir: string): Promise<MatterIdentifiers | undefined> {
  const file = path.join(matterDir, "matter-identifiers.json")
  if (!existsSync(file)) return
  const value = await Bun.file(file).json().catch(() => undefined) as unknown
  return decodeMatterIdentifiers(value)
}

function decodeMatterIdentifiers(value: unknown): MatterIdentifiers | undefined {
  if (!value || typeof value !== "object" || Array.isArray(value)) return
  const data = value as Record<string, unknown>
  if (data.schema_version !== IDENTIFIERS_SCHEMA_VERSION || typeof data.generated_at !== "string" || typeof data.documents !== "number" || !Array.isArray(data.rows)) return
  const generatedAt = Date.parse(data.generated_at)
  if (!Number.isFinite(generatedAt) || new Date(generatedAt).toISOString() !== data.generated_at) return
  return value as MatterIdentifiers
}

function sameIdentifierProjection(previous: MatterIdentifiers, current: MatterIdentifiers) {
  return previous.schema_version === current.schema_version &&
    previous.documents === current.documents &&
    isDeepStrictEqual(previous.rows, current.rows)
}

// Soft-wrap overly long lines so the generic read tool never truncates them.
// Quote and citation checks read these files, so the wrap never drops or
// changes a character. Each break replaces the last space (U+0020) that keeps
// the line within the cap, so the file keeps the source slice's length and
// every other character's offset. Only a run of more than READ_LINE_CAP
// characters without a space is cut, by inserting a line break (never between
// the two halves of a surrogate pair); `inserted` lists the offsets of `text`
// before which a cut was inserted. Short lines are untouched.
const READ_LINE_CAP = 1000

function wrapForReadability(text: string) {
  const lines: string[] = []
  const inserted: number[] = []
  let offset = 0
  for (const line of text.split("\n")) {
    const wrapped = wrapLine(line)
    lines.push(...wrapped.lines)
    inserted.push(...wrapped.cuts.map((cut) => offset + cut))
    offset += line.length + 1
  }
  return { text: lines.join("\n"), inserted }
}

function wrapLine(line: string) {
  const lines: string[] = []
  const cuts: number[] = []
  let start = 0
  while (line.length - start > READ_LINE_CAP) {
    const space = line.lastIndexOf(" ", start + READ_LINE_CAP)
    if (space > start) {
      lines.push(line.slice(start, space))
      start = space + 1
      continue
    }
    const end = start + (isSurrogatePair(line, start + READ_LINE_CAP - 1) ? READ_LINE_CAP - 1 : READ_LINE_CAP)
    lines.push(line.slice(start, end))
    cuts.push(end)
    start = end
  }
  lines.push(line.slice(start))
  return { lines, cuts }
}

function isSurrogatePair(text: string, at: number) {
  const high = text.charCodeAt(at)
  const low = text.charCodeAt(at + 1)
  return high >= 0xd800 && high <= 0xdbff && low >= 0xdc00 && low <= 0xdfff
}

function chunkSource(source: SourceRecord): SourceChunk[] {
  const blocks = textBlocks(source.text)
  if (!blocks.length) return [sourceChunk(source, 1, 0, source.text.length, 1, 1, [])]

  const chunks: SourceChunk[] = []
  const firstBlock = blocks[0]!
  const headingTrail: string[] = []
  let chunkStartChar = 0
  let chunkStartLine = firstBlock.startLine
  let chunkHeadingPath: string[] = []
  let chunkHasBlock = false
  let chunkCharacters = 0
  let lastBlockEndChar = 0
  let lastBlockEndLine = firstBlock.endLine

  for (const block of blocks) {
    const heading = headingText(block.text)
    const startsNewHeading = heading && chunkHasBlock
    const wouldOverflow = chunkHasBlock && chunkCharacters + (block.endChar - lastBlockEndChar) > MAX_SOURCE_CHUNK_CHARS
    if (startsNewHeading || wouldOverflow) {
      chunks.push(sourceChunk(source, chunks.length + 1, chunkStartChar, lastBlockEndChar, chunkStartLine, lastBlockEndLine, chunkHeadingPath))
      chunkStartChar = lastBlockEndChar
      chunkStartLine = block.startLine
      chunkCharacters = block.endChar - chunkStartChar
      chunkHasBlock = false
    }

    if (heading) headingTrail.push(heading)
    if (!chunkHasBlock) {
      chunkHeadingPath = headingTrail.slice(-4)
      chunkStartLine = block.startLine
      chunkCharacters = block.endChar - chunkStartChar
      chunkHasBlock = true
    } else {
      chunkCharacters = block.endChar - chunkStartChar
    }
    lastBlockEndChar = block.endChar
    lastBlockEndLine = block.endLine
  }

  chunks.push(sourceChunk(source, chunks.length + 1, chunkStartChar, source.text.length, chunkStartLine, lastBlockEndLine, chunkHeadingPath))
  return chunks
}

function sourceChunk(source: SourceRecord, index: number, startChar: number, endChar: number, startLine: number, endLine: number, headingPath: string[]) {
  const text = source.text.slice(startChar, endChar)
  const padded = String(index).padStart(3, "0")
  return {
    id: `${sourceSlug(source.document)}-c${padded}`,
    index,
    heading_path: headingPath,
    text_path: `source-chunks/${sourceSlug(source.document)}-c${padded}.txt`,
    start_line: startLine,
    end_line: endLine,
    start_char: startChar,
    end_char: endChar,
    characters: text.length,
    sha256: sha256Text(text),
    preview: preview(text),
  }
}

function textBlocks(text: string) {
  const lines = text.split("\n")
  const result: { text: string; startLine: number; endLine: number; startChar: number; endChar: number }[] = []
  let offset = 0
  let blockStartLine = 0
  let blockStartChar = 0
  let blockLines: string[] = []

  for (const [i, line] of lines.entries()) {
    const lineStart = offset
    const lineEnd = lineStart + line.length
    if (line.trim()) {
      if (!blockLines.length) {
        blockStartLine = i + 1
        blockStartChar = lineStart
      }
      blockLines.push(line)
    } else if (blockLines.length) {
      result.push({
        text: blockLines.join("\n"),
        startLine: blockStartLine,
        endLine: i,
        startChar: blockStartChar,
        endChar: lineEnd,
      })
      blockLines = []
    }
    offset = lineEnd + 1
  }

  if (blockLines.length) {
    result.push({
      text: blockLines.join("\n"),
      startLine: blockStartLine,
      endLine: lines.length,
      startChar: blockStartChar,
      endChar: text.length,
    })
  }
  return result
}

function headingText(text: string) {
  const line = text.trim()
  if (line.includes("\n")) return undefined
  if (line.length > 160) return undefined
  if (line.endsWith(".") && !/^(?:\d+(?:\.\d+)*|[IVXLCDM]+)\./i.test(line)) return undefined
  if (/^#{1,6}\s+\S/.test(line)) return line.replace(/^#{1,6}\s+/, "")
  if (/^(?:section|article|schedule|exhibit|annex|appendix|part)\b/i.test(line)) return line
  if (/^(?:\d+(?:\.\d+)*\.?|[IVXLCDM]+\.|[A-Z]\.)\s+\S/i.test(line)) return line
  if (/^[A-Z0-9][A-Z0-9 ,&/()':;-]{4,}$/.test(line) && /[A-Z]/.test(line)) return line
  return undefined
}

function sourceMapMarkdown(sourceMap: ReturnType<typeof buildSourceMap>) {
  return [
    "# Source Map",
    "",
    `Chunking: ${sourceMap.chunking.method}; max ${sourceMap.chunking.max_characters} characters per chunk unless a single structural block is longer.`,
    "",
    ...sourceMap.documents.flatMap((source) => [
      `## ${source.document}`,
      "",
      `Extracted text: \`${source.extracted_text}\``,
      "",
      `Characters: ${source.characters}`,
      "",
      "| chunk_id | file | lines | chars | heading path | preview |",
      "| --- | --- | ---: | ---: | --- | --- |",
      ...source.chunks.map((chunk) =>
        `| \`${chunk.id}\` | \`${chunk.text_path}\` | ${chunk.start_line}-${chunk.end_line} | ${chunk.characters} | ${tableCell(chunk.heading_path.join(" > ") || "(none)")} | ${tableCell(chunk.preview)} |`,
      ),
      "",
    ]),
  ].join("\n")
}

function preview(text: string) {
  return text.trim().replace(/\s+/g, " ").slice(0, 240)
}

function tableCell(value: string) {
  return value.replaceAll("|", "\\|").replaceAll("\n", " ")
}

function sourceSlug(value: string) {
  const name = path.basename(value)
  const readable = name.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "") || "source"
  return `${readable}-${createHash("sha256").update(name).digest("hex").slice(0, 12)}`
}

function sha256Text(value: string) {
  return createHash("sha256").update(value).digest("hex")
}
