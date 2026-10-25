import { DOMParser } from "@xmldom/xmldom"
import type { RevisionListEntry } from "docxodus"
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
//   its locations (each with its `side` when that paragraph holds only one side
//   of it), an inserted paragraph with its mark, adjacent same-author
//   insertions together. `id` is Docxodus's opaque `rev2-…` id for these bytes;
//   resolving a nested revision changes the enclosing one's id.
// - MarkupLocation.paragraph is the 1-based ordinal among its part's paragraphs
//   as marked up (deleted, moved-from, empty and table-cell paragraphs count;
//   separator notes do not); in scope `comment`, among its comment's own
//   paragraphs. Null when the anchor is unknown. Every revision has at least one
//   location. A change with no paragraph of its own (a table's, row's, cell's or
//   section's properties) is placed at its container's first paragraph.
// - MarkupComment.id is the native w:id (the rendered data-comment-id); `parent`
//   is the id of the comment it replies to; `resolved` is null when the file has
//   no commentsExtended entry for it. `quote` is the commented text as marked up,
//   deleted text included, capped; `paragraph` counts its part's paragraphs as
//   above, `label` is that paragraph's list number as a location's is, and
//   `scope` names that part when it is a footnote or endnote.
// - authors: each name with its counts, in order of first appearance; "" is a
//   change or comment that names no author.
export type MarkupField = "text" | "author" | "quote" | "initials"
export type MarkupFinding = { field: MarkupField; rule: string; detail: string; snippet?: string }
export type MarkupLocation = {
  paragraph: number | null
  context: string
  /** The list number Word prints before the paragraph ("5.", "(a)"), when it has one; not part of `context`. */
  label?: string
  /** What a change with no paragraph of its own changed; `paragraph` is that container's first. */
  container?: "table" | "row" | "cell" | "section"
  /** Scope `comment`: the w:id of the comment whose own text this paragraph is. */
  comment?: number
  /** A move's place: the paragraph it leaves (`from`) or lands in (`to`); absent when the paragraph holds both or neither. */
  side?: "from" | "to"
}
export type MarkupRevision = {
  id: string
  kind: "insert" | "delete" | "move" | "format" | "structure"
  /** Docxodus RevisionFamily, e.g. content_insert, paragraph_mark, properties_change, row_delete. */
  family: string
  author: string
  date: string | null
  text: string
  /** `comment`: an edit inside a comment's own text, not the document's. */
  scope: "body" | "header" | "footer" | "footnote" | "endnote" | "comment"
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
  /** The list number Word prints before `paragraph`, when it has one. */
  label?: string
  /** Absent for a comment in the body. */
  scope?: "footnote" | "endnote"
  findings: MarkupFinding[]
}
export type MarkupAuthor = { name: string; revisions: number; comments: number }
export type DocxMarkup = { revisions: MarkupRevision[]; comments: MarkupComment[]; authors: MarkupAuthor[] }
/** One change or open comment thread the other side made, as a register input. `source.sha256` names the
 * version whose ids `ref` uses; a thread's `text`, `author`, `date` and place are its root comment's. */
export type CounterpartyItem = {
  source: { name: string; sha256: string }
  ref: { kind: "revision"; id: string } | { kind: "comment"; id: number }
  author: string
  date: string | null
  change: "insert" | "delete" | "move" | "comment"
  text: string
  replies?: { author: string; date: string | null; text: string }[]
  /** A comment's anchor text. */
  quote?: string
  places: { part: MarkupRevision["scope"]; paragraph: number | null; label?: string; side?: "from" | "to" }[]
  status?: Exclude<MarkupRevision["status"], "supported">
  /** The untrusted-text rules the item's carried strings matched. */
  flagged?: string[]
}

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
// Anchor kinds that are paragraphs (h: heading, li: numbered list item).
const PARAGRAPHS = new Set(["p", "h", "li"])
const CONTAINERS: Record<string, MarkupLocation["container"]> = { tbl: "table", tr: "row", tc: "cell", sec: "section" }
// Where comments other than the body's can be anchored.
const NOTES = [["word/footnotes.xml", "footnote"], ["word/endnotes.xml", "endnote"]] as const
const MARKERS = ["commentRangeStart", "commentReference", "commentRangeEnd"]
// The parts a comment can be placed in, as MarkupRevision scopes.
const COMMENTED = new Set<MarkupRevision["scope"]>(["body", "footnote", "endnote"])
// Where a saved session keeps each element's anchor unid.
const PT = "http://powertools.codeplex.com/2011"
// Property subtrees hold formatting, never text.
const PROPERTIES = new Set(["pPr", "rPr", "tblPr", "tcPr", "trPr", "sectPr"])
const QUOTE_CAP = 200
// A range collects up to this many UTF-16 units, room for stripped characters.
const QUOTE_COLLECTED = QUOTE_CAP * 4

export async function readDocxMarkup(bytes: Uint8Array): Promise<DocxMarkup> {
  const { revisions, labels } = await readMarkupRevisions(bytes)
  const comments = await readMarkupComments(bytes, undefined, labels)
  const authors = new Map<string, MarkupAuthor>()
  const author = (name: string) => authors.get(name) ?? authors.set(name, { name, revisions: 0, comments: 0 }).get(name)!
  for (const revision of revisions) author(revision.author).revisions++
  for (const comment of comments) author(comment.author).comments++
  return { revisions, comments, authors: [...authors.values()] }
}

// The negotiable points of a markup: each insert, delete (a paragraph mark's
// included) and move outside comments' own text, then each unresolved thread
// with its replies folded in. Format and structure changes and resolved threads
// are left out. A reply joins the thread of a comment listed before it (Word
// lists replies after their parent); any other comment starts its own thread.
export function counterpartyItems(source: { name: string; sha256: string }, markup: DocxMarkup): CounterpartyItem[] {
  // Initials are not carried, so their findings are not either.
  const flagged = (findings: MarkupFinding[]) => {
    const rules = [...new Set(findings.filter(finding => finding.field !== "initials").map(finding => finding.rule))]
    return rules.length ? { flagged: rules } : {}
  }
  const units = markup.revisions.flatMap((revision): CounterpartyItem[] => {
    const { kind } = revision
    if (revision.scope === "comment" || kind === "format" || kind === "structure") return []
    return [{
      source, ref: { kind: "revision", id: revision.id }, author: revision.author, date: revision.date, change: kind, text: revision.text,
      places: revision.locations.map(({ paragraph, label, side }) => ({ part: revision.scope, paragraph, ...label ? { label } : {}, ...side ? { side } : {} })),
      ...revision.status === "supported" ? {} : { status: revision.status },
      ...flagged(revision.findings),
    }]
  })
  const threads: { root: MarkupComment; replies: MarkupComment[] }[] = [], member = new Map<number, (typeof threads)[number]>()
  for (const comment of markup.comments) {
    const joined = comment.parent === null ? undefined : member.get(comment.parent)
    if (joined) joined.replies.push(comment)
    else threads.push({ root: comment, replies: [] })
    member.set(comment.id, joined ?? threads.at(-1)!)
  }
  const comments = threads.flatMap(({ root, replies }): CounterpartyItem[] => root.resolved === true ? [] : [{
    source, ref: { kind: "comment", id: root.id }, author: root.author, date: root.date, change: "comment", text: root.text,
    ...replies.length ? { replies: replies.map(({ author, date, text }) => ({ author, date, text })) } : {},
    ...root.quote ? { quote: root.quote } : {},
    places: [{ part: root.scope ?? "body", paragraph: root.paragraph, ...root.label ? { label: root.label } : {} }],
    ...flagged([root, ...replies].flatMap(comment => comment.findings)),
  }])
  return [...units, ...comments]
}

// The revisions, and each list number by `${scope}:${paragraph}` for the parts
// comments are placed in, which takes them from the same projection.
async function readMarkupRevisions(bytes: Uint8Array): Promise<{ revisions: MarkupRevision[]; labels: Map<string, string> }> {
  // Anchor ids belong to the session that listed them, so locations come from
  // the same session; deleted and moved-from paragraphs keep their anchors, and
  // so does what it saves, which tells a move's two places apart.
  const session = (await docxodus()).openDocxSession(bytes, { persistAnchorIds: true, trackedChanges: "render_inline" })
  try {
    const rows = session.listRevisions(), listed = session.listComments(), labels = new Map<string, string>()
    if (!rows.length && !listed.length) return { revisions: [], labels }
    // The session projection carries Word's resolved list number (its declared
    // type omits the field). Ordinals count per part, and per comment in the
    // comments part, where each comment's anchor precedes its paragraphs. So
    // does a table's, row's or cell's; a section's follows its paragraphs, so a
    // section starts at the first paragraph after the previous one's anchor.
    const index = session.project().anchorIndex as Record<string, { kind: string; scope: string; textPreview: string; autoNumberPrefix?: string }>
    const comments = new Map(listed.map(row => [row.anchorId, row.id] as const))
    const sides = rows.some(row => row.type === "move") ? moveSides(session.save(), rows) : new Map<string, MarkupLocation["side"]>()
    const ordinals = new Map<string, number>(), located = new Map<string, MarkupLocation>()
    const waiting = new Map<string, string[]>(), sections = new Map<string, MarkupLocation>()
    let comment: string | undefined
    for (const [id, anchor] of Object.entries(index)) {
      if (anchor.kind === "cmt") comment = id
      const part = anchor.scope === "cmt" && comment ? comment : anchor.scope
      if (anchor.kind === "sec") {
        const start = sections.get(part)
        if (start) located.set(id, { ...start, container: "section" })
        sections.delete(part)
        continue
      }
      if (!PARAGRAPHS.has(anchor.kind)) {
        (waiting.get(part) ?? waiting.set(part, []).get(part)!).push(id)
        continue
      }
      const paragraph = (ordinals.get(part) ?? 0) + 1
      ordinals.set(part, paragraph)
      const label = normalizeUntrustedText(anchor.autoNumberPrefix ?? "").text.trim(), owner = comments.get(part)
      const location: MarkupLocation = { paragraph, context: normalizeUntrustedText(anchor.textPreview).text, ...label ? { label } : {}, ...owner === undefined ? {} : { comment: owner } }
      located.set(id, location)
      if (label && COMMENTED.has(scopeOf(anchor.scope))) labels.set(`${scopeOf(anchor.scope)}:${paragraph}`, label)
      if (!sections.has(part)) sections.set(part, location)
      // A container's own preview runs its cells' text together; its first paragraph's reads.
      for (const container of waiting.get(part) ?? []) {
        const kind = CONTAINERS[index[container]!.kind]
        located.set(container, kind ? { ...location, container: kind } : location)
      }
      waiting.delete(part)
    }
    const revisions = rows.map((row): MarkupRevision => {
      // A table, row or cell revision also lists its paragraphs; a property
      // change with none of its own is placed by its container.
      const anchors = row.affectedAnchors.length ? row.affectedAnchors : row.anchorId ? [{ id: row.anchorId, kind: row.anchorId.split(":")[0]! }] : []
      const own = anchors.filter(anchor => PARAGRAPHS.has(anchor.kind)).map(anchor => anchor.id), container = row.anchorId ?? anchors[0]?.id
      const locations = [...new Set(own.length ? own : container ? [container] : [])].map(id => {
        const location = located.get(id) ?? { paragraph: null, context: "" }, side = sides.get(`${row.id} ${id}`)
        return side ? { ...location, side } : location
      })
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
    return { revisions, labels }
  } finally {
    session.close()
  }
}

// Which side of its move each of a move's paragraphs is, by `${row.id} ${anchor id}`.
// Neither the listed anchors nor their order say (a move up lists its
// destination first); the carriers do: constituentKeys name each moveFrom and
// moveTo by its native w:id, and the saved package marks every paragraph with
// its anchor's unid. A paragraph holding both sides, or neither, has none.
function moveSides(saved: Uint8Array, rows: RevisionListEntry[]) {
  const parts = unzipSync(saved), paragraphs = new Map<string, Map<string, Element>>(), sides = new Map<string, MarkupLocation["side"]>()
  const byUnid = (partUri: string) => {
    const found = paragraphs.get(partUri)
    if (found) return found
    const part = parts[partUri.replace(/^\//, "")], own = new Map<string, Element>()
    if (part) for (const node of Array.from(new DOMParser().parseFromString(strFromU8(part), "application/xml").getElementsByTagNameNS(W, "p"))) {
      const unid = node.getAttributeNS(PT, "Unid")
      if (unid) own.set(unid, node)
    }
    return paragraphs.set(partUri, own).get(partUri)!
  }
  for (const row of rows) {
    if (row.type !== "move") continue
    // A content or paragraph-mark carrier, as `${role}:${w:id}`; range markers are not paragraphs' own.
    const carriers = new Set(row.constituentKeys.flatMap(key => /:(moveFrom|moveTo):(\d+)(?::|$)/.exec(key)?.slice(1, 3).join(":") ?? []))
    for (const anchor of row.affectedAnchors) {
      const node = byUnid(row.partUri).get(anchor.unid)
      const held = node ? (["moveFrom", "moveTo"] as const).filter(role => Array.from(node.getElementsByTagNameNS(W, role)).some(carrier => carriers.has(`${role}:${carrier.getAttributeNS(W, "id")}`))) : []
      if (held.length === 1) sides.set(`${row.id} ${anchor.id}`, held[0] === "moveFrom" ? "from" : "to")
    }
  }
  return sides
}

/** Comment threads; `native` is readDocxComments of the same bytes when the caller already has it,
 * and `labels` the list numbers readMarkupRevisions read from them. */
export async function readMarkupComments(bytes: Uint8Array, native = readDocxComments(bytes), labels = new Map<string, string>()): Promise<MarkupComment[]> {
  const listed = await (await docxodus()).getComments(bytes)
  if (!listed.length) return []
  const ids = new Map(listed.map(row => [row.anchorId, row.id]))
  const { quotes, places } = commentRanges(bytes, new Set(listed.map(row => String(row.id))))
  // The first native comment with each id, as a lookup.
  const natives = new Map(native.map(comment => [comment.comment_id, comment] as const).reverse())
  return listed.map(row => {
    const own = natives.get(String(row.id))
    // The range start places a comment; a reply often has only a reference.
    // The native markers are the body's; a note's come from the scan.
    const marker = MARKERS.map(kind => own?.markers.find(marker => marker.kind === kind && marker.xml_paragraph_index !== null)).find(Boolean)
    const place = marker ? undefined : places.get(String(row.id))
    const paragraph = marker ? marker.xml_paragraph_index! + 1 : place?.paragraph ?? null
    const label = paragraph === null ? undefined : labels.get(`${place?.scope ?? "body"}:${paragraph}`)
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
      paragraph,
      ...label ? { label } : {},
      ...place ? { scope: place.scope } : {},
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
  return scope.startsWith("hdr") ? "header" : scope.startsWith("ftr") ? "footer" : scope.startsWith("fn") ? "footnote" : scope.startsWith("en") ? "endnote" : scope.startsWith("cmt") ? "comment" : "body"
}

// The text inside each listed comment's range, in document order: inserted and
// deleted text alike, paragraphs joined by newlines. Ranges naming no listed
// comment are skipped, and a range closes once its quote is collected. A comment
// in a footnote or endnote is also placed: that part, and the paragraph holding
// its first marker by kind, counted as revision locations count them.
function commentRanges(bytes: Uint8Array, listed: Set<string>) {
  const parts = unzipSync(bytes)
  if (!parts["word/document.xml"]) throw new Error("invalid_docx_package")
  const fail = () => { throw new Error("invalid_docx_xml") }
  const quotes = new Map<string, string>(), places = new Map<string, { scope: (typeof NOTES)[number][1]; paragraph: number; rank: number }>()
  for (const [file, scope] of [["word/document.xml", "body"], ...NOTES] as const) {
    if (!parts[file]) continue
    const doc = new DOMParser({ errorHandler: { warning: fail, error: fail, fatalError: fail } }).parseFromString(strFromU8(parts[file]), "application/xml")
    const open = new Set<string>()
    let paragraphs = 0, paragraph = 0
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
      // Separator and continuation notes have no paragraph a reader sees.
      if ((name === "footnote" || name === "endnote") && (node.getAttributeNS(W, "type") || "normal") !== "normal") return
      const rank = MARKERS.indexOf(name)
      if (rank >= 0) {
        const id = node.getAttributeNS(W, "id") ?? ""
        if (scope !== "body" && paragraph && listed.has(id) && rank < (places.get(id)?.rank ?? MARKERS.length)) places.set(id, { scope, paragraph, rank })
        if (name === "commentRangeEnd") open.delete(id)
        else if (name === "commentRangeStart" && listed.has(id) && (quotes.get(id) ?? "").length < QUOTE_COLLECTED) { open.add(id);quotes.set(id, quotes.get(id) ?? "") }
        return
      }
      if (name === "t" || name === "delText") return add(node.textContent ?? "")
      if (name === "tab") return add("\t")
      if (name === "br" || name === "cr") return add("\n")
      const outer = paragraph
      if (name === "p") paragraph = ++paragraphs
      for (const child of Array.from(node.childNodes)) if (child.nodeType === 1) visit(child as Element)
      if (name === "p") { add("\n");paragraph = outer }
    }
    visit(doc.documentElement)
  }
  return { quotes: new Map([...quotes].map(([id, text]) => [id, text.replace(/^\n+|\n+$/g, "")])), places }
}
