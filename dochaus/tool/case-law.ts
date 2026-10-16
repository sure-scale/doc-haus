import { tool } from "./legacy-definition"
import { readCaseLawKey } from "../lib/research"

// doc.haus case-law research tool. Searches CourtListener's free legal-opinion
// database (https://www.courtlistener.com) — the same public corpus the major
// legal-AI tools verify citations against — and returns real, citable opinions.
//
// This is the one place a doc.haus agent reaches outside the matter: retrieval of
// a contract's own clauses stays local via `search-document`, but case law is
// public record that lives upstream, so this tool calls the live API rather than
// indexing it. It is read-only — it never writes to the matter — and returns only
// what CourtListener returns, with no fabricated citations.
//
// CourtListener works unauthenticated at a lower rate limit. A free token from a
// courtlistener.com account raises it; the firm pastes one in Settings > Research
// (read per-call from preferences.json), or it falls back to the
// COURTLISTENER_API_TOKEN env var. The tool sends it as an Authorization header
// when present, and never fails for lack of one.

const SEARCH_URL = "https://www.courtlistener.com/api/rest/v4/search/"
const SITE = "https://www.courtlistener.com"

type SearchResult = {
  caseName: string
  citation: string[]
  court: string
  dateFiled: string | null
  docketNumber: string | null
  absolute_url: string
  citeCount: number
  status: string
  opinions?: { snippet?: string }[]
}

export default tool({
  description:
    "Search U.S. case law (CourtListener) for judicial opinions relevant to a legal question, and return real, citable cases. Use this for legal research — finding precedent, how courts have treated a clause or doctrine — not for questions about the matter's own documents (use search-document for those). Every result is a real opinion; never cite a case this tool did not return.",
  args: {
    query: tool.schema
      .string()
      .describe("The legal question or terms to search for, e.g. 'enforceability of liquidated damages clause'"),
    court: tool.schema
      .string()
      .optional()
      .describe("Optional CourtListener court id to restrict to, e.g. 'scotus' for the U.S. Supreme Court."),
    k: tool.schema.number().int().min(1).max(20).optional().describe("Number of opinions to return (default 5)."),
  },
  async execute(args) {
    const params = new URLSearchParams({ q: args.query, type: "o", order_by: "score desc" })
    if (args.court) params.set("court", args.court)

    const token = readCaseLawKey() ?? process.env.COURTLISTENER_API_TOKEN
    const res = await fetch(`${SEARCH_URL}?${params}`, {
      headers: token ? { Authorization: `Token ${token}` } : {},
    })
    if (!res.ok) {
      return `CourtListener search failed (HTTP ${res.status}). The case-law database may be unreachable; report this to the user rather than inventing a citation.`
    }

    const k = args.k ?? 5
    const results = ((await res.json()) as { results?: SearchResult[] }).results ?? []
    const cases = results.slice(0, k).map((r) => ({
      caseName: r.caseName,
      citation: r.citation?.[0] ?? r.docketNumber ?? "no reporter citation",
      court: r.court,
      dateFiled: r.dateFiled,
      url: SITE + r.absolute_url,
      citeCount: r.citeCount,
      status: r.status,
      // Whitespace-collapse the highlighted match so the agent sees the passage,
      // not the opinion's raw line breaks.
      snippet: r.opinions?.[0]?.snippet?.replace(/\s+/g, " ").trim().slice(0, 400) ?? "",
    }))

    if (cases.length === 0) {
      return `No case law found for "${args.query}". Tell the user no opinions matched rather than inventing one.`
    }

    const output = cases
      .map(
        (c, i) =>
          `${i + 1}. ${c.caseName}, ${c.citation} (${c.court}${c.dateFiled ? `, ${c.dateFiled}` : ""})` +
          ` — cited by ${c.citeCount}, ${c.status}\n${c.url}` +
          (c.snippet ? `\n"${c.snippet}"` : ""),
      )
      .join("\n\n")

    return {
      title: `${cases.length} opinion(s) for "${args.query}"`,
      output,
      metadata: { cases },
    }
  },
})
