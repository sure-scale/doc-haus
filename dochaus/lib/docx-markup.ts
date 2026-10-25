import { DOMParser } from "@xmldom/xmldom"
import { strFromU8, unzipSync } from "fflate"
import { docxodus } from "./docxodus"
import { readDocxComments } from "./docx-comments"
import { detectInjection, normalizeUntrustedText } from "./untrusted-text"

// Counterparty markup in one Word file: its pending tracked changes and comment
// threads, read for the triage panel, the markup route and the agents' comments
// view. A computed view of these bytes; nothing is stored. Docxodus groups and
// resolves the revisions; this module locates and screens them.
//
// - Every third-party string (text, author, quote, initials) passes the shared
//   untrusted-text rules: invisible characters are stripped and reported, the
//   visible text is otherwise verbatim, and `findings` names each match by field.
//   Location `context` is body text (ingest scans it) and is only normalized.
//   A date is the recorded timestamp as a normalized ISO string (a change's
//   w16du:dateUtc before its w:date), or null when it does not parse.
// - MarkupRevision is one Docxodus unit, decided as a whole: a move with both
//   its locations, an inserted paragraph with its mark, adjacent same-author
//   insertions together. `id` is Docxodus's opaque `rev2-…` id for these bytes;
//   resolving a nested revision changes the enclosing one's id.
// - MarkupLocation.paragraph is the 1-based ordinal among its part's paragraphs
//   as marked up (deleted, moved-from, empty and table-cell paragraphs count);
//   null when the anchor is unknown. Every revision has at least one location.
// - MarkupComment.id is the native w:id (the rendered data-comment-id); `parent`
//   is the id of the comment it replies to; `resolved` is null when the file has
//   no commentsExtended entry for it. `quote` is the commented text as marked up,
//   deleted text included, capped; `paragraph` counts body paragraphs as above.
// - authors: each name with its counts, in order of first appearance; "" is a
//   change or comment that names no author.
export type MarkupField = "text" | "author" | "quote" | "initials"
export type MarkupFinding = { field: MarkupField; rule: string; detail: string; snippet?: string }
export type MarkupLocation = { paragraph: number | null; context: string }
export type MarkupRevision = {
  id: string
  kind: "insert" | "delete" | "move" | "format" | "structure"
  /** Docxodus RevisionFamily, e.g. content_insert, paragraph_mark, properties_change, row_delete. */
  family: string
  author: string
  date: string | null
  text: string
  scope: "body" | "header" | "footer" | "footnote" | "endnote"
  locations: MarkupLocation[]
  /** Only `supported` units can be accepted or rejected here; the rest are resolved in Word. */
  status: "supported" | "unsupported" | "malformed" | "ambiguous"
  diagnostic: { code: string; message: string } | null
  findings: MarkupFinding[]
}
export type MarkupComment = {
  id: number
  author: string
  initials: string | null
  date: string | null
  text: string
  parent: number | null
  resolved: boolean | null
  quote: string
  paragraph: number | null
  findings: MarkupFinding[]
}
export type MarkupAuthor = { name: string; revisions: number; comments: number }
export type DocxMarkup = { revisions: MarkupRevision[]; comments: MarkupComment[]; authors: MarkupAuthor[] }

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
// Anchor kinds that are paragraphs (h: heading, li: numbered list item).
const PARAGRAPHS = new Set(["p", "h", "li"])
// Property subtrees hold formatting, never text.
const PROPERTIES = new Set(["pPr", "rPr", "tblPr", "tcPr", "trPr", "sectPr"])
const QUOTE_CAP = 200
// A range collects up to this many UTF-16 units, room for stripped characters.
const QUOTE_COLLECTED = QUOTE_CAP * 4

export async function readDocxMarkup(bytes: Uint8Array): Promise<DocxMarkup> {
  const revisions = await readMarkupRevisions(bytes)
  const comments = await readMarkupComments(bytes)
  const authors = new Map<string, MarkupAuthor>()
  const author = (name: string) => authors.get(name) ?? authors.set(name, { name, revisions: 0, comments: 0 }).get(name)!
  for (const revision of revisions) author(revision.author).revisions++
  for (const comment of comments) author(comment.author).comments++
  return { revisions, comments, authors: [...authors.values()] }
}

async function readMarkupRevisions(bytes: Uint8Array): Promise<MarkupRevision[]> {
  // Anchor ids belong to the session that listed them, so locations come from
  // the same session; deleted and moved-from paragraphs keep their anchors.
  const session = (await docxodus()).openDocxSession(bytes, { persistAnchorIds: false, trackedChanges: "render_inline" })
  try {
    const rows = session.listRevisions()
    if (!rows.length) return []
    const ordinals = new Map<string, number>(), located = new Map<string, MarkupLocation>()
    for (const [id, anchor] of Object.entries(session.project().anchorIndex)) {
      if (!PARAGRAPHS.has(anchor.kind)) continue
      const paragraph = (ordinals.get(anchor.scope) ?? 0) + 1
      ordinals.set(anchor.scope, paragraph)
      located.set(id, { paragraph, context: normalizeUntrustedText(anchor.textPreview).text })
    }
    return rows.map(row => {
      // A table, row or cell revision also lists its paragraphs; the container
      // itself has no paragraph ordinal.
      const anchors = row.affectedAnchors.length ? row.affectedAnchors : row.anchorId ? [{ id: row.anchorId, kind: row.anchorId.split(":")[0]! }] : []
      const locations = [...new Set(anchors.filter(anchor => PARAGRAPHS.has(anchor.kind)).map(anchor => anchor.id))]
        .map(id => located.get(id) ?? { paragraph: null, context: "" })
      const findings: MarkupFinding[] = []
      return {
        id: row.id, kind: row.type, family: row.family,
        author: screened(findings, "author", row.author),
        date: timestamp(row.dateUtc ?? row.date),
        text: screened(findings, "text", row.text),
        scope: scopeOf(row.scope),
        locations: locations.length ? locations : [{ paragraph: null, context: "" }],
        status: row.resolutionStatus,
        diagnostic: row.diagnostic ? { code: row.diagnostic.code, message: row.diagnostic.message } : null,
        findings,
      }
    })
  } finally {
    session.close()
  }
}

/** Comment threads; `native` is readDocxComments of the same bytes when the caller already has it. */
export async function readMarkupComments(bytes: Uint8Array, native = readDocxComments(bytes)): Promise<MarkupComment[]> {
  const listed = await (await docxodus()).getComments(bytes)
  if (!listed.length) return []
  const ids = new Map(listed.map(row => [row.anchorId, row.id]))
  const quotes = commentQuotes(bytes, new Set(listed.map(row => String(row.id))))
  // The first native comment with each id, as a lookup.
  const natives = new Map(native.map(comment => [comment.comment_id, comment] as const).reverse())
  return listed.map(row => {
    const own = natives.get(String(row.id))
    // The range start places a comment; a reply often has only a reference.
    const marker = ["commentRangeStart", "commentReference", "commentRangeEnd"]
      .map(kind => own?.markers.find(marker => marker.kind === kind && marker.xml_paragraph_index !== null)).find(Boolean)
    const findings: MarkupFinding[] = []
    return {
      id: row.id,
      author: screened(findings, "author", row.author),
      initials: row.initials === undefined ? null : screened(findings, "initials", row.initials),
      date: timestamp(row.date),
      // The native text keeps the comment's paragraph breaks; Docxodus joins them with spaces.
      text: screened(findings, "text", own?.text ?? row.text),
      parent: row.parentAnchorId === undefined ? null : ids.get(row.parentAnchorId) ?? null,
      resolved: row.resolved ?? null,
      // Scanned as collected; only the shown quote is capped.
      quote: capped(screened(findings, "quote", quotes.get(String(row.id)) ?? "")),
      paragraph: marker ? marker.xml_paragraph_index! + 1 : null,
      findings,
    }
  })
}

function screened(findings: MarkupFinding[], field: MarkupField, raw: string) {
  const normalized = normalizeUntrustedText(raw)
  for (const { rule, detail, snippet } of [...normalized.findings, ...detectInjection(normalized.text)])
    findings.push({ field, rule, detail, ...(snippet === undefined ? {} : { snippet }) })
  return normalized.text
}

function timestamp(raw: string | undefined) {
  // A time without a zone keeps the clock it records, not the server's offset.
  const time = raw === undefined ? NaN : Date.parse(/T[\d:.]+$/.test(raw) ? `${raw}Z` : raw)
  return Number.isNaN(time) ? null : new Date(time).toISOString().replace(/\.000Z$/, "Z")
}

function capped(text: string) {
  const points = Array.from(text)
  return points.length > QUOTE_CAP ? `${points.slice(0, QUOTE_CAP).join("")}…` : text
}

function scopeOf(scope: string): MarkupRevision["scope"] {
  return scope.startsWith("hdr") ? "header" : scope.startsWith("ftr") ? "footer" : scope.startsWith("fn") ? "footnote" : scope.startsWith("en") ? "endnote" : "body"
}

// The body text inside each listed comment's range, in document order: inserted
// and deleted text alike, paragraphs joined by newlines. Ranges naming no listed
// comment are skipped, and a range closes once its quote is collected.
function commentQuotes(bytes: Uint8Array, listed: Set<string>) {
  const part = unzipSync(bytes)["word/document.xml"]
  if (!part) throw new Error("invalid_docx_package")
  const fail = () => { throw new Error("invalid_docx_xml") }
  const doc = new DOMParser({ errorHandler: { warning: fail, error: fail, fatalError: fail } }).parseFromString(strFromU8(part), "application/xml")
  const open = new Set<string>(), quotes = new Map<string, string>()
  const add = (text: string) => {
    if (text) for (const id of open) {
      const quote = quotes.get(id)!
      quotes.set(id, quote + text.slice(0, QUOTE_COLLECTED - quote.length))
      if (quote.length + text.length >= QUOTE_COLLECTED) open.delete(id)
    }
  }
  const visit = (node: Element) => {
    const name = node.namespaceURI === W ? node.localName : ""
    if (PROPERTIES.has(name)) return
    if (name === "commentRangeStart" || name === "commentRangeEnd") {
      const id = node.getAttributeNS(W, "id") ?? ""
      if (name === "commentRangeEnd") open.delete(id)
      else if (listed.has(id) && (quotes.get(id) ?? "").length < QUOTE_COLLECTED) { open.add(id);quotes.set(id, quotes.get(id) ?? "") }
      return
    }
    if (name === "t" || name === "delText") return add(node.textContent ?? "")
    if (name === "tab") return add("\t")
    if (name === "br" || name === "cr") return add("\n")
    for (const child of Array.from(node.childNodes)) if (child.nodeType === 1) visit(child as Element)
    if (name === "p") add("\n")
  }
  visit(doc.documentElement)
  return new Map([...quotes].map(([id, text]) => [id, text.replace(/^\n+|\n+$/g, "")]))
}
