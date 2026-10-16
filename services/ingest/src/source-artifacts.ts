import type { Database } from "bun:sqlite"
import { createHash, randomUUID } from "node:crypto"
import { existsSync, mkdirSync, readFileSync, renameSync, rmSync, writeFileSync } from "node:fs"
import path from "node:path"
import { isDeepStrictEqual } from "node:util"
import { listSourceDocuments, openDb } from "./db"
import { extractIdentifiers, renderIdentifiersMarkdown, type IdentifierChunk, type MatterIdentifiers } from "./identifiers"
import { stampIngestLineage } from "./lineage"
import { reconstructText } from "./structure"

const MAX_SOURCE_CHUNK_CHARS = 4500

type SourceRecord = {
  document: string
  doc_path: string
  extracted_text: string
  characters: number
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

export async function writeSourceArtifacts(
  matterDir: string,
  correlation?: string,
  options: { db?: Database; settle?: () => void; stampLineage?: boolean; transactionDirectory?: string } = {},
) {
  const previousIdentifiers = await readMatterIdentifiers(matterDir)
  const db = options.db ?? openDb(matterDir)
  const sources = (listSourceDocuments(db) as { doc_path: string; name: string }[])
    .map((doc) => ({
      document: doc.name,
      doc_path: doc.doc_path,
      extracted_text: `sources/${sourceSlug(doc.name)}.txt`,
      text: reconstructText(db, doc.doc_path).trim(),
    }))
    .map((source) => ({ ...source, characters: source.text.length }))
  if (!options.db) db.close()

  const documents = sources.map((s) => s.document)
  const correlationId = correlation ?? `ingest-${Date.now().toString(36)}`
  const transaction = options.transactionDirectory ?? path.join(
    matterDir,
    ".dochaus",
    "source-artifact-transactions",
    `${process.pid}-${randomUUID()}`,
  )
  const staged = path.join(transaction, "staged")
  mkdirSync(staged, { recursive: true })
  if (!sources.length) {
    settleSourceArtifactSnapshot(matterDir, transaction, options.settle)
    if (options.stampLineage !== false) stampIngestLineage(matterDir, documents, correlationId)
    return { documents: 0, chunks: 0, sourceDocuments: documents }
  }

  const sourceMap = buildSourceMap(sources)
  try {
    mkdirSync(path.join(staged, "sources"), { recursive: true })
    mkdirSync(path.join(staged, "source-chunks"), { recursive: true })
    await Promise.all(sources.map((source) => Bun.write(path.join(staged, source.extracted_text), `${source.text}\n`)))
    await writeSourceChunks(staged, sources, sourceMap)
    await writeMatterIdentifiers(staged, sourceMap, previousIdentifiers)
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
    settleSourceArtifactSnapshot(matterDir, transaction, options.settle)
  } catch (error) {
    if (
      !existsSync(path.join(transaction, "prior-artifacts.json")) &&
      !existsSync(path.join(transaction, "backup"))
    ) rmSync(transaction, { recursive: true, force: true })
    throw error
  }
  if (options.stampLineage !== false) stampIngestLineage(matterDir, documents, correlationId)
  return {
    documents: sources.length,
    chunks: sourceMap.documents.reduce((sum, source) => sum + source.chunks.length, 0),
    sourceDocuments: documents,
  }
}

export function sourceArtifactsMatchRegistry(matterDir: string, suppliedDb?: Database) {
  const db = suppliedDb ?? openDb(matterDir)
  const sources = (listSourceDocuments(db) as { doc_path: string; name: string }[])
    .map((document): SourceRecord => {
      const text = reconstructText(db, document.doc_path).trim()
      return {
        document: document.name,
        doc_path: document.doc_path,
        extracted_text: `sources/${sourceSlug(document.name)}.txt`,
        text,
        characters: text.length,
      }
    })
  if (!suppliedDb) db.close()
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
    const textByDocument = new Map(sources.map((source) => [source.document, source.text]))
    const identifierChunks = expectedMap.documents.flatMap((document) => {
      const text = textByDocument.get(document.document)
      if (text === undefined) throw new Error(`Missing source text for ${document.document}.`)
      return document.chunks.map((chunk) => {
        const expected = wrapForReadability(text.slice(chunk.start_char, chunk.end_char))
        if (
          !existsSync(path.join(matterDir, chunk.text_path)) ||
          readFileSync(path.join(matterDir, chunk.text_path), "utf8") !== expected
        ) throw new Error(`Stale source chunk ${chunk.id}.`)
        return { id: chunk.id, document: document.document, text: expected }
      })
    })
    const identifiers = decodeMatterIdentifiers(
      JSON.parse(readFileSync(path.join(matterDir, "matter-identifiers.json"), "utf8")),
    )
    if (!identifiers) return false
    const expectedIdentifiers = {
      ...extractIdentifiers(identifierChunks),
      generated_at: identifiers.generated_at,
    }
    if (!isDeepStrictEqual(identifiers, expectedIdentifiers)) return false
    return readFileSync(path.join(matterDir, "matter-identifiers.md"), "utf8") ===
      `${renderIdentifiersMarkdown(expectedIdentifiers)}\n`
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

async function writeSourceChunks(matterDir: string, sourceRows: SourceRecord[], sourceMap: ReturnType<typeof buildSourceMap>) {
  const textByDocument = new Map(sourceRows.map((source) => [source.document, source.text]))
  await Promise.all(
    sourceMap.documents.flatMap((source) =>
      source.chunks.map((chunk) =>
        // Chunk files are read by drafting agents through the generic `read`
        // tool, which truncates any single line beyond
        // 2000 characters mid-clause. Legal prose often emits an entire
        // clause or recital as one line, so the artifact is soft-wrapped on
        // sentence/clause boundaries before writing. This is safe: the
        // authoritative offsets live in source-map.json (char ranges into the
        // reconstructed extracted text), never in these read-only files.
        Bun.write(
          path.join(matterDir, chunk.text_path),
          wrapForReadability(textByDocument.get(source.document)?.slice(chunk.start_char, chunk.end_char) ?? ""),
        ),
      ),
    ),
  )
}

async function writeMatterIdentifiers(
  matterDir: string,
  sourceMap: ReturnType<typeof buildSourceMap>,
  previous: MatterIdentifiers | undefined,
) {
  const chunks: IdentifierChunk[] = []
  for (const doc of sourceMap.documents) {
    for (const chunk of doc.chunks) {
      const text = await Bun.file(path.join(matterDir, chunk.text_path)).text()
      chunks.push({
        id: chunk.id,
        document: doc.document,
        text,
      })
    }
  }
  const extracted = extractIdentifiers(chunks)
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
  if (data.schema_version !== 1 || typeof data.generated_at !== "string" || typeof data.documents !== "number" || !Array.isArray(data.rows)) return
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
// Splits first on sentence/clause boundaries, then greedily packs the pieces
// into lines under the cap, hard-wrapping any single piece that still exceeds
// it. Existing short lines are untouched.
const READ_LINE_CAP = 1000

function wrapForReadability(text: string): string {
  return text
    .split("\n")
    .flatMap((line) => (line.length <= READ_LINE_CAP ? [line] : wrapLine(line)))
    .join("\n")
}

function wrapLine(line: string): string[] {
  const pieces = line.split(/(?<=\.)\s+|(?<=;)\s+|\s+—\s+/).filter(Boolean)
  const out: string[] = []
  let current = ""
  for (const piece of pieces) {
    if (piece.length > READ_LINE_CAP) {
      if (current) {
        out.push(current)
        current = ""
      }
      for (let i = 0; i < piece.length; i += READ_LINE_CAP) out.push(piece.slice(i, i + READ_LINE_CAP))
      continue
    }
    if (current && current.length + 1 + piece.length > READ_LINE_CAP) {
      out.push(current)
      current = piece
    } else {
      current = current ? `${current} ${piece}` : piece
    }
  }
  if (current) out.push(current)
  return out
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
