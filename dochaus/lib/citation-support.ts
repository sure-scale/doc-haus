import { approvedResearch } from "../workflows/lean/research"
import { existsSync, readFileSync, readdirSync } from "node:fs"
import path from "node:path"
import { readAuthorityLedger } from "./authority-ledger"
import type { CitationSupport } from "./citation-check"
import { readContextManifest } from "./context-manifest"

// Citation support: the matter texts a citation may be found in (authority
// ledger, context manifest, verified research rows, source artifacts and source
// chunks), normalized for lookup by citation key.

export type SupportHaystack = {
  normalized_text: string
  support: CitationSupport
}

export function buildSupportIndex(directory: string): SupportHaystack[] {
  return [
    ...authorityLedgerSupport(directory),
    ...contextManifestSupport(directory),
    ...researchSupport(directory),
    ...approvedResearch(directory).flatMap(row => {
      const support: CitationSupport = { kind: "research-row", id: `lean-research:${row.id}`, status: "verified", artifact: ".dochaus/lean/research.json", url: row.url, research_row_id: row.id }
      return [haystack([row.url, row.citation, row.quote].join(" "), support)].filter((item): item is SupportHaystack => item !== null)
    }),
    ...sourceArtifactSupport(directory),
  ].sort((a, b) => a.support.id.localeCompare(b.support.id))
}

function authorityLedgerSupport(directory: string): SupportHaystack[] {
  return readAuthorityLedger(directory)
    .map((entry) => {
      const status = entry.status === "verified" ? "verified" as const : "unresolved" as const
      return haystack(entry.raw, {
        kind: "authority-ledger" as const,
        id: entry.id,
        status,
        artifact: entry.source.artifact,
        url: entry.kind === "url" ? entry.raw : undefined,
      })
    })
    .filter((item): item is SupportHaystack => item !== null)
}

function contextManifestSupport(directory: string): SupportHaystack[] {
  return readContextManifest(directory)
    .filter((row) => row.event === "source_fetch" && row.url)
    .map((row) =>
      haystack(row.url ?? "", {
        kind: "official-fetch",
        id: `official-fetch:${row.entry_id}`,
        status: "verified",
        url: row.url,
      }),
    )
    .filter((item): item is SupportHaystack => item !== null)
}

function researchSupport(directory: string): SupportHaystack[] {
  const file = path.join(directory, "research.md")
  if (!existsSync(file)) return []
  return parseResearchRows(readFileSync(file, "utf8")).flatMap((row, index) => {
    const rowID = stringField(row, "id") || `research-row-${index + 1}`
    const status = verifiedResearchRow(row) ? "verified" as const : "unresolved" as const
    const support: CitationSupport = {
      kind: "research-row",
      id: `research-row:${rowID}`,
      status,
      artifact: "research.md",
      research_row_id: rowID,
    }
    return [
      haystack(JSON.stringify(row), support),
      ...researchRowSources(row).map((source, sourceIndex) =>
        haystack(
          [stringField(source, "title"), stringField(source, "url")].filter(Boolean).join(" "),
          {
            ...support,
            kind: "source-index" as const,
            id: `research-source:${rowID}:${sourceIndex + 1}`,
            url: stringField(source, "url") || undefined,
          },
        ),
      ),
    ].filter((item): item is SupportHaystack => item !== null)
  })
}

function sourceArtifactSupport(directory: string): SupportHaystack[] {
  return [
    ...readSupportFile(directory, "source-index.json", "source-index"),
    ...readSupportFile(directory, "source-map.md", "source-map"),
    ...readSupportFile(directory, "source-map.json", "source-map"),
    ...readSupportFile(directory, "source-ledger.json", "source-map"),
    ...sourceIndexExtractedText(directory),
    ...sourceChunkFiles(directory),
  ]
}

function readSupportFile(
  directory: string,
  relative: string,
  kind: CitationSupport["kind"],
): SupportHaystack[] {
  const file = path.join(directory, relative)
  if (!existsSync(file)) return []
  const support: CitationSupport = { kind, id: `${kind}:${relative}`, status: "verified", artifact: relative }
  const item = haystack(readFileSync(file, "utf8"), support)
  return item ? [item] : []
}

function sourceIndexExtractedText(directory: string): SupportHaystack[] {
  const index = parseJsonFile(path.join(directory, "source-index.json"))
  if (!isRecord(index) || !Array.isArray(index.sources)) return []
  return index.sources.filter(isRecord).flatMap((source) => {
    const extracted = stringField(source, "extracted_text")
    if (!safeRelativePath(directory, extracted)) return []
    const file = path.join(directory, extracted)
    if (!existsSync(file)) return []
    const support: CitationSupport = {
      kind: "source-artifact",
      id: `source-artifact:${extracted}`,
      status: "verified",
      artifact: extracted,
    }
    const item = haystack(readFileSync(file, "utf8"), support)
    return item ? [item] : []
  })
}

function sourceChunkFiles(directory: string): SupportHaystack[] {
  const chunkDir = path.join(directory, "source-chunks")
  if (!existsSync(chunkDir)) return []
  return readdirSync(chunkDir)
    .filter((file) => file.endsWith(".txt"))
    .sort((a, b) => a.localeCompare(b))
    .flatMap((file) => {
      const relative = `source-chunks/${file}`
      const support: CitationSupport = {
        kind: "source-artifact",
        id: `source-artifact:${relative}`,
        status: "verified",
        artifact: relative,
      }
      const item = haystack(readFileSync(path.join(directory, relative), "utf8"), support)
      return item ? [item] : []
    })
}

function haystack(text: string, support: CitationSupport): SupportHaystack | null {
  const normalized_text = normalizeSupportText(text)
  if (!normalized_text) return null
  return { normalized_text, support }
}

/** With `wholeWord`, the key must be whole words of the support text: it
 * neither continues a word before it ("down v jones" is not in "ashdown v
 * jones", "ash-down v jones" or "o'down v jones") nor runs on into one after it
 * ("s. 5" is not "s. 51", "v jones" is not "v jonesco"). */
export function resolveSupport(normalizedKey: string, supportIndex: SupportHaystack[], options: { wholeWord?: boolean } = {}): CitationSupport[] {
  const byID = new Map<string, CitationSupport>()
  for (const item of supportIndex) {
    if (!(options.wholeWord ? wholeWordAt(item.normalized_text, normalizedKey) : item.normalized_text.includes(normalizedKey))) continue
    byID.set(item.support.id, item.support)
  }
  return [...byID.values()].sort((a, b) => a.id.localeCompare(b.id))
}

function wholeWordAt(text: string, key: string) {
  const word = (ch: string | undefined) => /[\p{L}\p{N}]/u.test(ch ?? "")
  for (let at = text.indexOf(key); at !== -1; at = text.indexOf(key, at + 1)) {
    // A hyphen, apostrophe or underscore after a letter or digit joins the word.
    const continues = word(text[at - 1]) || /[-_'\u2019]/.test(text[at - 1] ?? "") && word(text[at - 2])
    if ((!word(key[0]) || !continues) && (!word(key.at(-1)) || !word(text[at + key.length]))) return true
  }
  return false
}

function normalizeSupportText(text: string) {
  return text
    .replace(/\u00a0/g, " ")
    .replace(/[“”]/g, '"')
    .replace(/[’]/g, "'")
    .replace(/\s*§+\s*/g, " § ")
    .replace(/\bv\.?\b/gi, " v ")
    .replace(/\s+([,.;:)])/g, "$1")
    .replace(/([(])\s+/g, "$1")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase()
}

function parseResearchRows(content: string): Record<string, unknown>[] {
  const match = /```research\s*\n?([\s\S]*?)```/i.exec(content)
  if (!match?.[1]) return []
  const parsed = parseJson(match[1])
  return Array.isArray(parsed) ? parsed.filter(isRecord) : []
}

function parseJsonFile(file: string): unknown {
  if (!existsSync(file)) return null
  return parseJson(readFileSync(file, "utf8"))
}

function parseJson(text: string): unknown {
  try {
    return JSON.parse(text) as unknown
  } catch {
    return null
  }
}

function researchRowSources(row: Record<string, unknown>) {
  return Array.isArray(row.sources) ? row.sources.filter(isRecord) : []
}

function verifiedResearchRow(row: Record<string, unknown>) {
  const tier = stringField(row, "tier").toLowerCase()
  return tier === "verified" || tier === "corroborated"
}

function stringField(record: Record<string, unknown>, key: string): string {
  const value = record[key]
  return typeof value === "string" ? value.trim() : ""
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value)
}

function safeRelativePath(directory: string, relative: string) {
  if (path.isAbsolute(relative)) return false
  const resolved = path.resolve(directory, relative)
  return resolved !== directory && resolved.startsWith(path.resolve(directory) + path.sep)
}
