import { INGEST_URL } from "../../config"

// The other side's markup in an uploaded Word source (#200,
// services/ingest/src/document-markup.ts): its pending tracked changes and
// comment threads, read from one version's bytes (dochaus/lib/docx-markup.ts).
// Nothing is stored; deciding changes writes one new version.

// Which field of an item matched a screening rule, and what it found.
export type MarkupFinding = { field: "text" | "author" | "quote" | "initials"; rule: string; detail: string; snippet?: string }

// One affected paragraph: its 1-based place among its part's paragraphs as
// marked up (in scope `comment`, among that comment's own), null when unknown,
// and a short preview of it. `label` is the list number Word prints before it
// ("5.", "(a)"), not part of `context`; `container` is set for a table's,
// row's, cell's or section's own formatting, placed at its first paragraph;
// `comment` is the id of the comment whose text the paragraph is; `side` is
// the place a move leaves (`from`) or lands in (`to`), absent when not known.
export type MarkupLocation = {
  paragraph: number | null
  context: string
  label?: string
  container?: "table" | "row" | "cell" | "section"
  comment?: number
  side?: "from" | "to"
}

// One change, decided as a whole: a move with both its places, an inserted
// paragraph with its mark, adjacent insertions by one author together. `id` is
// opaque and belongs to these bytes; it is never shown.
export type MarkupRevision = {
  id: string
  kind: "insert" | "delete" | "move" | "format" | "structure"
  // The exact operation, e.g. content_insert, paragraph_mark, row_delete.
  family: string
  author: string
  date: string | null
  text: string
  // `comment`: an edit inside a comment's own text, not the document's.
  scope: "body" | "header" | "footer" | "footnote" | "endnote" | "comment"
  locations: MarkupLocation[]
  // Only supported changes can be decided here; the rest are resolved in Word.
  status: "supported" | "unsupported" | "malformed" | "ambiguous"
  diagnostic: { code: string; message: string } | null
  findings: MarkupFinding[]
}

// One comment; `parent` is the id of the comment it replies to and `resolved`
// is null when the file does not say. `paragraph` counts its part's paragraphs
// and `label` is the list number Word prints there, as a location's is;
// `scope` names that part when it is not the body.
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
  label?: string
  scope?: "footnote" | "endnote"
  findings: MarkupFinding[]
}

export type MarkupAuthor = { name: string; revisions: number; comments: number }

export type DocumentMarkup = {
  name: string
  // The version these items were read from; its sha256 is the base a
  // decision batch must still match.
  version: { id: string; number: number; sha256: string }
  current: boolean
  revisions: MarkupRevision[]
  comments: MarkupComment[]
  authors: MarkupAuthor[]
}

export type MarkupDecision = "accept" | "reject" | "hold"

// What one batch did. Every field is read tolerantly: the counts say what was
// applied, and a change resolved along with another one is `consumed`.
export type MarkupDecisionResult = {
  name?: string
  version?: { id: string; number: number } | null
  counts?: Partial<Record<"accepted" | "rejected" | "held" | "consumed", number>>
  notice?: string
}

// The service refused a request: the file changed since the list was read, a
// change can no longer be decided here, doc.haus proposals are pending (409),
// or nothing was accepted or rejected (422). `message` is the service's own
// words; `code` is whatever reason string it sent, known to this build or not;
// `ids` are the listed changes the refusal is about (`id`, then those it
// `merged_with`), empty when it names none.
export class MarkupRefusedError extends Error {
  constructor(readonly status: number, readonly code: string | undefined, message: string, readonly ids: string[] = []) {
    super(message)
  }
}

const route = (matterId: string, path: string, name: string) =>
  `${INGEST_URL}/matters/${encodeURIComponent(matterId)}/documents/markup${path}?name=${encodeURIComponent(name)}`

async function refusal(res: Response, fallback: string) {
  const body = (await res.json().catch(() => undefined)) as { error?: unknown; code?: unknown; reason?: unknown; id?: unknown; merged_with?: unknown } | undefined
  const code = typeof body?.code === "string" ? body.code : typeof body?.reason === "string" ? body.reason : undefined
  const message = typeof body?.error === "string" && body.error.trim() ? body.error : code ?? `${fallback} (${res.status})`
  const ids = [body?.id, ...(Array.isArray(body?.merged_with) ? body.merged_with : [])].filter((id): id is string => typeof id === "string" && id !== "")
  return new MarkupRefusedError(res.status, code, message, [...new Set(ids)])
}

// The working file's markup. Refused (4xx) for a file it does not cover: a
// generated deliverable, or anything but a registered Word source.
export async function fetchDocumentMarkup(matterId: string, name: string): Promise<DocumentMarkup> {
  const res = await fetch(route(matterId, "", name))
  if (!res.ok) throw await refusal(res, `Could not read the tracked changes in ${name}`)
  return res.json()
}

// Header values must be ISO-8859-1; a name outside plain ASCII is left to the
// service, which records the saved attorney name instead.
const HEADER_SAFE = /^[\x20-\x7e]+$/

// Whether `actor` travels with the decisions itself, needing no saved name.
export const actorSent = (actor: string) => HEADER_SAFE.test(actor.trim())

export async function applyMarkupDecisions(
  matterId: string,
  name: string,
  body: { base_sha256: string; decisions: { id: string; decision: MarkupDecision }[] },
  actor?: string,
): Promise<MarkupDecisionResult> {
  const headers: Record<string, string> = { "Content-Type": "application/json" }
  const who = actor?.trim()
  if (who && actorSent(who)) headers["x-dochaus-actor"] = who
  const res = await fetch(route(matterId, "/decisions", name), { method: "POST", headers, body: JSON.stringify(body) })
  if (!res.ok) throw await refusal(res, `Could not apply the decisions to ${name}`)
  return (await res.json().catch(() => ({}))) as MarkupDecisionResult
}
