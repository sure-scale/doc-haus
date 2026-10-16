import { tool } from "./legacy-definition"
import { isOfficialSource, researchScope } from "../lib/research"
import { fetchSourceText } from "../lib/research-source-fetch"
import { displayedSourceSnippets, renderSourceSnippets, researchTerms } from "../lib/research-source-snippets"
import { preserveResearchSourceSnapshot } from "../lib/research-source-snapshot"

export default tool({
  description:
    "Fetch an official legal source URL and return compact snippets around supplied research terms. Use this when web-search has found, or the matter already supplies, an official statute, regulation, court, agency, or government URL and generic webfetch is too large or noisy. It is authority retrieval, not web discovery.",
  args: {
    url: tool.schema.string().describe("Official source URL to fetch."),
    terms: tool.schema
      .string()
      .optional()
      .describe("Comma-, semicolon-, pipe-, or newline-separated terms or short phrases to locate in the source."),
    maxSnippets: tool.schema.number().min(1).max(10).optional().describe("Maximum snippets to return (default 4)."),
  },
  async execute(args, ctx) {
    const official = isOfficialSource(args.url)
    if (!official && researchScope() !== "open")
      return {
        title: "Source outside research fence",
        output:
          `official-source is limited to official primary legal sources unless the firm enables open-web research. ` +
          `${new URL(args.url).hostname} is outside the current fence.`,
        metadata: { official, blocked: true },
      }

    const result = await fetchSourceText(args.url)
    if (result.status === "blocked")
      return {
        title: `Source blocked: ${result.descriptor.host}`,
        output: `Could not retrieve official source. Reason: ${result.reason}. URL: ${result.descriptor.finalUrl}`,
        metadata: { official, blocked: true, reason: result.reason, ...result.descriptor },
      }

    const terms = researchTerms(args.terms)
    const snippets = displayedSourceSnippets(result.text, terms, args.maxSnippets ?? 4)
    const snapshot = preserveResearchSourceSnapshot(ctx.directory, result, snippets)
    return {
      title: `Official source: ${result.descriptor.host}`,
      output: (result.descriptor.retrieval ? "Retrieval: Exa live-crawl text extraction. Original source bytes and extraction completeness are unverified; this is not a generated summary.\n" : "") + renderSourceSnippets({
        url: result.descriptor.finalUrl,
        contentType: result.descriptor.contentType,
        text: result.text,
        terms,
        maxSnippets: args.maxSnippets ?? 4,
        snapshot: snapshot.sha256,
      }),
      metadata: {
        official,
        partial: result.status === "partial" || undefined,
        reason: result.status === "partial" ? result.reason : undefined,
        ...result.descriptor,
        snippets: snippets.length,
        sourceSnapshot: snapshot,
      },
    }
  },
})
