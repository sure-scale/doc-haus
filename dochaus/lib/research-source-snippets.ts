export type SourceSnippet = {
  term: string
  text: string
  start: number
  end: number
}

const SNIPPET_CHARS = 1_800

export function researchTerms(input?: unknown) {
  const text = typeof input === "string" ? input : Array.isArray(input) ? input.filter((term) => typeof term === "string").join("\n") : ""
  return [
    ...new Set(
      text
        .split(/[\n,;|]+/)
        .map((term) => term.trim())
        .filter((term) => term.length >= 3),
    ),
  ]
}

export function sourceSnippets(text: string, terms: string[], maxSnippets: number) {
  const lower = text.toLowerCase()
  const seen = new Set<string>()
  return terms
    .flatMap((term) => {
      const at = firstTermIndex(lower, term)
      if (at === -1) return []
      const start = Math.max(0, at - Math.floor(SNIPPET_CHARS / 2))
      const end = Math.min(text.length, at + term.length + Math.floor(SNIPPET_CHARS / 2))
      const snippet = sourceSnippet(text, start, end, term)
      const key = snippet.text.slice(0, 120)
      if (seen.has(key)) return []
      seen.add(key)
      return [snippet]
    })
    .slice(0, Math.max(1, Math.min(maxSnippets, 10)))
}

export function renderSourceSnippets(args: {
  url: string
  contentType: string
  text: string
  terms: string[]
  maxSnippets: number
  snapshot?: string
}) {
  const snippets = displayedSourceSnippets(args.text, args.terms, args.maxSnippets)
  // A snippet's primary search term is only its anchor. Other requested terms
  // may be visible in the same window, including overlapping capped windows.
  const visible = snippets.map((snippet) => normalizedTermText(snippet.text))
  const missing = args.terms.filter((term) => !visible.some((text) => text.includes(normalizedTermText(term))))
  const body = snippets.map((snippet, index) =>
    `<snippet index="${index + 1}" matched="${escapeAttr(snippet.term)}" locator="snippet-${index + 1}">\n${escapeText(snippet.text)}\n</snippet>`,
  ).join("\n\n")

  return (
    `<official_source url="${escapeAttr(args.url)}" content_type="${escapeAttr(args.contentType || "unknown")}"${args.snapshot ? ` source_snapshot="${escapeAttr(args.snapshot)}"` : ""}>\n` +
    `${body}\n` +
    `<missing_terms>${missing.map((term) => `<term>${escapeText(term)}</term>`).join("")}</missing_terms>\n` +
    `</official_source>\n` +
    `Use only text shown inside <snippet> as fetched authority. If the relevant rule is not visible, mark the research row practice_standard_unretrieved or open rather than relying on memory. ` +
    `Text matching does not establish semantic or legal support.`
  )
}

export function displayedSourceSnippets(text: string, terms: string[], maxSnippets: number) {
  const snippets = sourceSnippets(text, terms, maxSnippets)
  return snippets.length ? snippets : [sourceSnippet(text, 0, SNIPPET_CHARS, "first-readable-text")]
}

function sourceSnippet(text: string, start: number, end: number, term: string): SourceSnippet {
  const slice = text.slice(start, end)
  // Remove only tokens actually cut by the window. A complete first or last
  // word can be a qualification; BOF/EOF never imply a clipped token.
  const leading = start > 0 && /\S/.test(text[start - 1]!) && /^\S/.test(slice)
    ? slice.match(/^\S+/)?.[0].length ?? 0 : 0
  const trailing = end < text.length && /\S/.test(text[end]!) && /\S$/.test(slice)
    ? slice.match(/\S+$/)?.[0].length ?? 0 : 0
  const retained = slice.slice(leading, Math.max(leading, slice.length - trailing))
  const excerpt = retained.trim()
  const offset = start + leading + retained.indexOf(excerpt)
  return { term, text: excerpt, start: offset, end: offset + excerpt.length }
}

function firstTermIndex(lower: string, term: string) {
  const direct = lower.indexOf(term.toLowerCase())
  if (direct !== -1) return direct
  return term
    .toLowerCase()
    .split(/\s+/)
    .filter((token) => token.length >= 4)
    .map((token) => lower.indexOf(token))
    .find((at) => at !== -1) ?? -1
}

function normalizedTermText(text: string) {
  return text.normalize("NFKC").replace(/\s+/g, " ").trim().toLowerCase()
}

function escapeAttr(text: string) {
  return escapeText(text).replace(/"/g, "&quot;")
}

function escapeText(text: string) {
  return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
}
