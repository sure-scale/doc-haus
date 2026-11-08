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
// CourtListener's search refuses a request without a token (HTTP 401). A free
// token comes from a courtlistener.com account; the firm pastes one in
// Settings > Research (read per-call from preferences.json), or it falls back to
// the COURTLISTENER_API_TOKEN env var. The tool sends it as an Authorization
// header when present. A refused request says that a token is needed and that
// nothing was searched, so the gap is disclosed rather than read as an outage.

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
    // "Case law", not the tool's own name: tool names in a reply are flagged as injection.
    if (res.status === 401 || res.status === 403) {
      return `CourtListener refused the search (HTTP ${res.status}): it needs a CourtListener API token, added under Settings > Research. No case law was searched. Disclose this as a research limitation and never cite a case this search did not return.`
    }
    if (!res.ok) {
      return `CourtListener search failed (HTTP ${res.status}); the service may be unreachable. No case law was searched. Disclose this as a research limitation and never cite a case this search did not return.`
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
