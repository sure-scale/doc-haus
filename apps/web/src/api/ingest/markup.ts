import { INGEST_URL } from "../../config"

// The other side's markup in an uploaded Word source (#200,
// services/ingest/src/document-markup.ts): its pending tracked changes and
// comment threads, read from one version's bytes (dochaus/lib/docx-markup.ts).
// Nothing is stored; deciding changes, with any approved replies to the
// other side's comments, writes one new version.

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
// `scope` names that part when it is not the body. `detached`: its reference
// is no longer in the document (the text it was on was removed), so it takes
// no reply.
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
  detached?: true
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
  // How the matter's current drafting run answered each change and thread of
  // these bytes, by the change's or root comment's id; absent when it answered
  // none. `reason` is model text.
  responses?: Record<string, MarkupResponse>
}

export type MarkupResponse = {
  disposition: string
  reason?: string
  amendments?: { id: string; pid: string; outcome: string }[]
  question?: { id: string; status: string }
}

export type MarkupDecision = "accept" | "reject" | "hold"

// What one batch did. Every field is read tolerantly: the counts say what was
// applied, and a change resolved along with another one is `consumed`;
// `replies` are the comments added, by the thread they answer. `consistency`
// is the cross-references and defined terms (#198) the decisions broke in the
// new version, or that they could not be checked; it is kept nowhere.
export type MarkupDecisionResult = {
  name?: string
  version?: { id: string; number: number } | null
  counts?: Partial<Record<"accepted" | "rejected" | "held" | "consumed", number>>
  replies?: { thread: number; id: number }[]
  notice?: string
  consistency?: { status: "checked"; findings: { detail: string }[] } | { status: "not_checked" }
}

// The service refused a request: the file changed since the list was read, a
// change can no longer be decided here, doc.haus proposals are pending, a
// thread is gone, resolved or detached (409), nothing was accepted, rejected or
// replied, or a decision removes the text a reply's comment is on (422), or a
// reply could not be drafted (409 unreadable, 502). `message` is
// the service's own words; `code` is whatever reason string it sent, known to
// this build or not; `ids` are the listed changes the refusal is about (`id`,
// then those it `merged_with`), empty when it names none; `thread` is the
// root comment whose reply it is about.
export class MarkupRefusedError extends Error {
  constructor(readonly status: number, readonly code: string | undefined, message: string, readonly ids: string[] = [], readonly thread?: number) {
    super(message)
  }
}

const route = (matterId: string, path: string, name: string) =>
  `${INGEST_URL}/matters/${encodeURIComponent(matterId)}/documents/markup${path}?name=${encodeURIComponent(name)}`

async function refusal(res: Response, fallback: string) {
  const body = (await res.json().catch(() => undefined)) as { error?: unknown; code?: unknown; reason?: unknown; id?: unknown; merged_with?: unknown; thread?: unknown } | undefined
  const code = typeof body?.code === "string" ? body.code : typeof body?.reason === "string" ? body.reason : undefined
  const message = typeof body?.error === "string" && body.error.trim() ? body.error : code ?? `${fallback} (${res.status})`
  const ids = [body?.id, ...(Array.isArray(body?.merged_with) ? body.merged_with : [])].filter((id): id is string => typeof id === "string" && id !== "")
  const thread = typeof body?.thread === "number" ? body.thread : undefined
  return new MarkupRefusedError(res.status, code, message, [...new Set(ids)], thread)
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

// `replies` are the lawyer's approved replies, at most one per open thread,
// by its root comment's id; they are signed with the recorded name.
export async function applyMarkupDecisions(
  matterId: string,
  name: string,
  body: { base_sha256: string; decisions: { id: string; decision: MarkupDecision }[]; replies?: { thread: number; text: string }[] },
  actor?: string,
): Promise<MarkupDecisionResult> {
  const headers: Record<string, string> = { "Content-Type": "application/json" }
  const who = actor?.trim()
  if (who && actorSent(who)) headers["x-dochaus-actor"] = who
  const res = await fetch(route(matterId, "/decisions", name), { method: "POST", headers, body: JSON.stringify(body) })
  if (!res.ok) throw await refusal(res, `Could not apply the decisions to ${name}`)
  return (await res.json().catch(() => ({}))) as MarkupDecisionResult
}

// A draft reply to one open thread, for the lawyer to edit and approve: the
// service reads the thread, its clause and the changes there from
// `base_sha256` itself, with the panel's undecided `decisions` and an optional
// one-line `steer`. `open_items` are what the lawyer should check; a `[[…]]`
// in the reply is a part only the lawyer can fill. Nothing is stored. Aborting
// `signal` ends the model call.
export type MarkupReplyDraft = { reply: string; open_items: string[] }

export async function draftMarkupReply(
  matterId: string,
  name: string,
  body: { base_sha256: string; thread: number; decisions?: { id: string; decision: MarkupDecision }[]; steer?: string },
  signal?: AbortSignal,
): Promise<MarkupReplyDraft> {
  const res = await fetch(route(matterId, "/replies/draft", name), { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body), signal })
  if (!res.ok) throw await refusal(res, "Could not draft the reply")
  const draft = (await res.json().catch(() => ({}))) as { reply?: unknown; open_items?: unknown }
  if (typeof draft.reply !== "string" || !draft.reply.trim()) throw new MarkupRefusedError(res.status, "draft_unreadable", "The draft could not be read. Try again, or write the reply.")
  const items = Array.isArray(draft.open_items) ? draft.open_items.filter((item): item is string => typeof item === "string" && item.trim() !== "") : []
  return { reply: draft.reply, open_items: items }
}
