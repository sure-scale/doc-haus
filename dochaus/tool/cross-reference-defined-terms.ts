import { tool } from "./legacy-definition"
import path from "node:path"
import { discoverDeliverable } from "../lib/verification-loop"

// doc.haus cross-reference-defined-terms tool. Reads every original source
// chunk and the draft body (if available), extracts candidate named terms
// (multi-word proper-noun phrases), and cross-references them across sources
// to surface mismatches the source-reconciler's bounded search budget may miss.
//
// The reconciler has a bounded retrieval budget (6 searches, 2 chunk reads).
// When its search query doesn't match the source's wording, a named reference
// is missed entirely. This tool reads ALL chunks mechanically — no search
// needed — so it catches cross-source reference gaps that the reconciler cannot.
//
// Transfer-safe: extraction is pattern-based (proper-noun phrases), not
// document-family-specific. It works for any runtime-classified document
// family. The reconciler makes the final judgment on which candidates are
// genuine named terms.

export default tool({
  description:
    'Read all original source chunks and the draft body (if available), extract candidate named terms (multi-word proper-noun phrases), and cross-reference across sources. Flags: (1) single-source-reference — a named term found in only one source when multiple sources exist (potential stale reference, detected before draft assembly); (2) draft-only — a named term in the draft that appears in NO source; (3) single-source — a named term in the draft that appears in only one source while other sources exist. Writes defined-terms-cross-reference.md with findings.',
  args: {
    body: tool.schema
      .string()
      .optional()
      .describe('Optional draft filename (default: the discovered deliverable).'),
    output: tool.schema
      .string()
      .optional()
      .describe('Output report filename (default "defined-terms-cross-reference.md"; ".md" is appended if missing).'),
  },
  async execute(args, ctx) {
    const chunkFiles = [...new Bun.Glob("source-chunks/*.txt").scanSync({ cwd: ctx.directory, onlyFiles: true })]
      .filter(f => !/draft|memo|notes|strategy|research|register|spine|binding|review|audit|coverage|issue/i.test(f))
      .sort()

    if (!chunkFiles.length) {
      const outputFile = outputMarkdownFilename(args.output)
      await Bun.write(
        path.join(ctx.directory, outputFile),
        "# Defined-Terms Cross-Reference Report\n\nNo source chunks found in source-chunks/. The cross-reference runs after ingest creates source chunks.\n",
      )
      return `No source chunks found. Empty report written.`
    }

    type Chunk = { file: string; source: string; text: string }
    const chunks: Chunk[] = []
    const sourceNames = new Set<string>()

    for (const file of chunkFiles) {
      const text = await Bun.file(path.join(ctx.directory, file)).text()
      const source = file.replace(/-c\d+\.txt$/, "").replace(/\.txt$/, "").replace(/^source-chunks\//, "")
      chunks.push({ file, source, text })
      sourceNames.add(source)
    }

    const bodyName = args.body ? path.basename(args.body) : discoverDeliverable(ctx.directory)
    const draft = bodyName ? await Bun.file(path.join(ctx.directory, bodyName)).text() : ""

    const sourceTerms = new Map<string, Map<string, string[]>>()
    for (const c of chunks) {
      const terms = extractCandidateTerms(c.text)
      for (const { term, context } of terms) {
        if (!sourceTerms.has(term)) sourceTerms.set(term, new Map())
        const sources = sourceTerms.get(term)!
        if (!sources.has(c.source)) sources.set(c.source, [])
        sources.get(c.source)!.push(context)
      }
    }

    const draftTerms = extractCandidateTerms(draft)

    type Gap = { type: string; term: string; description: string; detail: string }
    const gaps: Gap[] = []

    for (const [term, sources] of sourceTerms) {
      if (sources.size === 1 && sourceNames.size > 1) {
        const inSource = [...sources.keys()][0]
        const notIn = [...sourceNames].filter(s => s !== inSource)
        gaps.push({
          type: "single-source-reference",
          term,
          description: `"${term}" found only in: ${inSource}`,
          detail: `Not found in: ${notIn.join(", ")}. This named reference may be stale, superseded, or specific to one source — verify currency and cross-source consistency before carrying it forward.`,
        })
      }
    }

    for (const { term } of draftTerms) {
      const sources = sourceTerms.get(term)
      if (!sources || sources.size === 0) {
        gaps.push({
          type: "draft-only",
          term,
          description: `"${term}" appears in the draft but in NO source document`,
          detail: `This named term was not found in any source chunk. It may be invented, misnamed, stale, or superseded. Verify provenance before keeping it.`,
        })
      } else if (sources.size === 1 && sourceNames.size > 1) {
        const inSource = [...sources.keys()][0]
        const notIn = [...sourceNames].filter(s => s !== inSource)
        gaps.push({
          type: "single-source",
          term,
          description: `"${term}" appears in the draft and in only 1 source: ${inSource}`,
          detail: `Not found in: ${notIn.join(", ")}. This named reference may be stale, superseded, or specific to one source — verify provenance and cross-source consistency.`,
        })
      }
    }

    const outputFile = outputMarkdownFilename(args.output)

    const report = formatReport(sourceTerms, draftTerms, gaps, [...sourceNames])
    await Bun.write(path.join(ctx.directory, outputFile), report)

    const summary = gaps.length > 0
      ? `Cross-reference: ${gaps.length} gap(s) found. See ${outputFile}.`
      : `Cross-reference: all named terms consistent across sources. Report written to ${outputFile}.`

    return summary
  },
})

function outputMarkdownFilename(output: string | undefined) {
  const filename = path.basename(output?.trim() || "defined-terms-cross-reference.md") || "defined-terms-cross-reference.md"
  return filename.endsWith(".md") ? filename : `${filename}.md`
}

// Words that indicate a structural/generic reference, not a named term.
const STRUCTURAL_STARTERS = new Set([
  "This", "That", "Such", "Said", "These", "Those",
  "Section", "Article", "Exhibit", "Schedule", "Appendix",
  "Paragraph", "Clause", "Table", "Figure", "Note", "Chapter",
  "Title", "Part", "Page", "Line",
])

// Determiners that indicate a generic two-word reference rather than a name.
const GENERIC_DETERMINERS = new Set([
  "this", "the", "a", "an", "such", "said", "any", "each",
  "every", "its", "their", "no", "under", "our", "your", "his", "her",
])

// Candidate labels come from source shape, not a document-family vocabulary:
// explicitly quoted labels and multi-word capitalized names. The reconciler
// makes the final materiality judgment, so this mechanical pass may be
// over-inclusive but must not silently privilege one legal document family.
const QUOTED_LABEL_RE = /[“"]([A-Z][A-Za-z0-9&'.\- ]{0,80})[”"]/g
const NAME_WORD = `[A-Z][A-Za-z0-9&'\\-]*(?:(?:\\.[A-Z][A-Za-z0-9&'\\-]*)+)?`
const CONNECTOR = `(?:of|the|and|or|for|in|to)`
const CAPITALIZED_NAME_RE = new RegExp(`\\b(${NAME_WORD}(?:\\s+(?:${CONNECTOR}\\s+)?${NAME_WORD}){1,6})\\b`, "g")

function extractCandidateTerms(text: string): { term: string; context: string }[] {
  const results: { term: string; context: string }[] = []
  const seen = new Set<string>()

  if (!text) return results

  for (const re of [QUOTED_LABEL_RE, CAPITALIZED_NAME_RE]) {
    re.lastIndex = 0
    let match: RegExpExecArray | null
    while ((match = re.exec(text)) !== null) {
      const term = normalizeCandidate(match[1])
      if (isStructural(term)) continue
      if (seen.has(term.toLowerCase())) continue
      seen.add(term.toLowerCase())

      const lineStart = text.lastIndexOf("\n", match.index)
      const lineEnd = text.indexOf("\n", match.index)
      const context = text.slice(lineStart + 1, lineEnd > 0 ? lineEnd : text.length).trim().slice(0, 200)
      results.push({ term, context })
    }
  }

  return results
}

function isStructural(term: string): boolean {
  const words = term.split(/\s+/)
  if (!term) return true

  // Reject phrases starting with structural words
  if (STRUCTURAL_STARTERS.has(words[0])) return true

  // Reject two-word phrases that begin with a generic determiner.
  if (words.length === 2 && GENERIC_DETERMINERS.has(words[0].toLowerCase())) return true

  return false
}

function normalizeCandidate(value: string) {
  return value
    .replace(/^(?:This|The|A|An|Such|Said|Its|Their|Under|Of|In|To|Our|Your|His|Her)\s+/i, "")
    .replace(/\s+/g, " ")
    .trim()
}

function formatReport(
  sourceTerms: Map<string, Map<string, string[]>>,
  draftTerms: { term: string; context: string }[],
  gaps: Gap[],
  sourceNames: string[],
): string {
  let report = `# Defined-Terms Cross-Reference Report\n\n`
  report += `Sources checked: ${sourceNames.join(", ")}\n\n`

  if (gaps.length === 0 && sourceTerms.size === 0) {
    report += `**No candidate named terms found.** Either no named references exist in the sources, or the extraction patterns did not match.\n`
    return report
  }

  report += `## Named terms in draft\n\n`
  if (draftTerms.length === 0) {
    report += `No draft body found or no candidate named terms in the draft.\n\n`
  } else {
    report += `| Term | Sources containing it |\n|------|---------------------|\n`
    for (const { term } of draftTerms) {
      const sources = sourceTerms.get(term)
      if (!sources || sources.size === 0) {
        report += `| ${term} | **NONE** |\n`
      } else {
        report += `| ${term} | ${[...sources.keys()].join(", ")} |\n`
      }
    }
    report += `\n`
  }

  if (gaps.length > 0) {
    report += `## Gaps requiring reconciler attention\n\n`
    for (const g of gaps) {
      report += `- **[${g.type}]** ${g.description}\n  ${g.detail}\n\n`
    }
  } else {
    report += `**No gaps detected.** All named terms in the draft appear in the expected sources.\n`
  }

  report += `\n## All candidate named terms across sources\n\n`
  if (sourceTerms.size > 0) {
    report += `| Term | Sources |\n|------|--------|\n`
    for (const [term, sources] of [...sourceTerms.entries()].sort((a, b) => a[0].localeCompare(b[0]))) {
      report += `| ${term} | ${[...sources.keys()].join(", ")} |\n`
    }
  }

  return report
}

type Gap = { type: string; term: string; description: string; detail: string }
