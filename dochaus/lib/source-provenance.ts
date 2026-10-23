import path from "node:path"
import { readRegularFile } from "./regular-file"

// Where an extracted matter source came from. Ingest indexes each attachment of
// an email and each file in a zip as its own source and records its container
// as `attachment_of` in source-index.json (services/ingest/src/source-artifacts.ts).
// Tools that name a source show this chain beside the name, e.g.
// "in Exhibits.zip, attachment of Thread.eml", so the model knows a schedule
// arrived with a particular email. It is display-only: the source's name, text
// and offsets are unchanged, so citations still name and verify the plain source.

type SourceLink = { document: string; attachment_of?: string }

// Each child source's provenance chain, keyed by its exact source name. Plain
// uploads are absent. A container missing from the index (removed from the
// matter) ends the chain; a container already on the chain ends it too, so a
// malformed index cannot loop.
export function provenanceChains(sources: readonly SourceLink[]): Map<string, string> {
  const parentOf = new Map(sources.map((source) => [source.document, source.attachment_of] as const))
  const chains = new Map<string, string>()
  for (const source of sources) {
    const parts: string[] = []
    const seen = new Set([source.document])
    let parent = source.attachment_of
    while (parent && !seen.has(parent)) {
      seen.add(parent)
      const present = parentOf.has(parent)
      parts.push(`${/\.zip$/i.test(parent) ? "in" : "attachment of"} ${displayName(parent)}${present ? "" : " (removed from matter)"}`)
      parent = present ? parentOf.get(parent) : undefined
    }
    if (parts.length) chains.set(source.document, parts.join(", "))
  }
  return chains
}

// The matter's provenance chains from its source-index.json. Provenance is
// supplementary context, so a missing or unreadable index yields none rather
// than failing the tool that asked.
export function sourceProvenance(directory: string): Map<string, string> {
  try {
    const bytes = readRegularFile(path.join(directory, "source-index.json"))
    if (!bytes) return new Map()
    const index = JSON.parse(bytes.toString("utf8")) as { sources?: unknown }
    if (!Array.isArray(index.sources)) return new Map()
    return provenanceChains(
      index.sources.flatMap((source): SourceLink[] =>
        source && typeof source === "object" && typeof source.document === "string"
          ? [{ document: source.document, ...(typeof source.attachment_of === "string" ? { attachment_of: source.attachment_of } : {}) }]
          : [],
      ),
    )
  } catch {
    return new Map()
  }
}

// Uploaded names are untrusted (ingest rejects only "/"): keep the chain on one line.
function displayName(name: string) {
  return name.replace(/[\p{Cc}\u2028\u2029]+/gu, " ")
}

// A source name in a list of names, with its chain in a labelled parenthetical
// that is not part of the name. A plain upload is just its name.
export function nameWithProvenance(name: string, provenance: ReadonlyMap<string, string>) {
  const chain = provenance.get(name)
  return chain ? `${name} (provenance: ${chain})` : name
}

// A "Provenance:" block after a tool's output, one line per distinct child
// document the output names, so every line above keeps the plain name. Empty
// when none of the names is a child.
export function provenanceFooter(names: Iterable<string>, provenance: ReadonlyMap<string, string>) {
  const lines = [...new Set(names)].flatMap((name) => {
    const chain = provenance.get(name)
    return chain ? [`${name}: ${chain}`] : []
  })
  return lines.length ? `\n\nProvenance:\n${lines.join("\n")}` : ""
}

// For prompts that show names as a JSON array consumers parse unchanged: a
// separate labelled line mapping each child name to its chain. Empty when the
// list has no child, so a plain matter's prompt is unchanged.
export function provenanceMapLine(names: Iterable<string>, provenance: ReadonlyMap<string, string>) {
  const map: Record<string, string> = {}
  for (const name of names) {
    const chain = provenance.get(name)
    if (chain) map[name] = chain
  }
  return Object.keys(map).length ? `Source provenance (where each attachment or archive entry came from; display only, not part of the source names or citations): ${JSON.stringify(map)}` : ""
}
