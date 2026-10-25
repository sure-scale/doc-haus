import { useEffect, useId, useMemo, useRef, useState, type KeyboardEvent } from "react"
import { humanizeError } from "../api/ingest/errors"
import {
  actorSent,
  applyMarkupDecisions,
  draftMarkupReply,
  MarkupRefusedError,
  type DocumentMarkup,
  type MarkupComment,
  type MarkupDecision,
  type MarkupDecisionResult,
  type MarkupFinding,
  type MarkupLocation,
  type MarkupRevision,
} from "../api/ingest/markup"
import { getPreferences, savePreferences } from "../api/ingest/preferences"
import { readBrowserJSON, removeBrowserValue, writeBrowserValue } from "../browser-storage"
import { formatDateTime } from "../lib/format"
import { useToast } from "./Toast"

// The other side's tracked changes and comments in an uploaded Word source
// (#200), beside the document. Every change starts on Hold; the lawyer marks
// Accept or Reject, and Apply sends every decision at once, which saves one new
// version. Held changes stay as tracked changes in it. Drafts of the decisions
// are kept in this browser per document, so closing the viewer loses nothing,
// and carry over to a new version of the file for the changes it still lists.
// Each open comment thread takes one reply: drafted by the selected model on
// request (never on opening), or written by hand, then edited and approved by
// the lawyer. Only approved replies go with Apply, in the same version, signed
// with the recorded name. Every name, change, quote and comment is the other
// side's text and renders as plain text.

type Thread = { root: MarkupComment; replies: MarkupComment[] }
type Item =
  | { type: "change"; revision: MarkupRevision; order: number[] }
  | { type: "thread"; thread: Thread; order: number[] }
type Kind = "changes" | "formatting" | "comments"
type Drafts = Record<string, Exclude<MarkupDecision, "hold">>
// The lawyer's reply to an open thread, by its root comment's id: the text,
// whether the lawyer approved it, what its draft said to check, and the root
// and number of comments it answers.
export type Reply = { text: string; approved: boolean; open_items: string[]; root: { author: string; text: string }; seen: number }
type Replies = Record<number, Reply>
// A refused Apply: the message, the service's code and the listed changes or
// the thread it names, whose cards carry the message.
type Refusal = { message: string; code?: string; ids: string[]; thread?: number }
// Where focus goes once the panel has rendered: a change's decision, a
// thread's reply (`reply`: its next control, not the refusal), the head's
// Draft replies row, or Apply.
type Focus = { card: string; reply?: boolean } | "apply" | "bulk"

const KINDS: { value: Kind; label: string }[] = [
  { value: "changes", label: "Changes" },
  { value: "formatting", label: "Formatting" },
  { value: "comments", label: "Comments" },
]
const DECISIONS: { value: MarkupDecision; label: string }[] = [
  { value: "accept", label: "Accept" },
  { value: "reject", label: "Reject" },
  { value: "hold", label: "Hold" },
]
const SCOPES = ["body", "header", "footer", "footnote", "endnote", "comment"] as const
// What a place outside the body is prefixed with.
const PART: Record<MarkupRevision["scope"], string> = { body: "", header: "header, ", footer: "footer, ", footnote: "footnote, ", endnote: "endnote, ", comment: "comment, " }
// A table's, row's, cell's or section's own formatting, which has no paragraph.
const CONTAINER: Record<NonNullable<MarkupLocation["container"]>, string> = {
  table: "Table formatting changed",
  row: "Table row formatting changed",
  cell: "Table cell formatting changed",
  section: "Page setup changed",
}
const STRUCTURE: Record<string, string> = {
  row_insert: "Table row inserted",
  row_delete: "Table row deleted",
  cell_insert: "Table cell inserted",
  cell_delete: "Table cell deleted",
  cell_merge: "Table cells merged",
  content_control_insert: "Form field inserted",
  content_control_delete: "Form field deleted",
  numbering_properties_insert: "Numbering added",
  numbering_change: "Numbering changed",
}
// Rules that only report characters stripped from the text; every other rule
// is text that reads like an instruction to an assistant.
const HIDDEN = new Set(["bidi-control", "invisible-tag", "zero-width"])
// Replies drafted at once, as the review grid fills cells.
const POOL = 3
// A part of a draft only the lawyer can fill.
const PLACEHOLDER = /\[\[[\s\S]*?\]\]/
const REPLY_MAX = 4000
const STEER_MAX = 1000

// Comments grouped into threads: each reply under the comment its chain starts
// from, in the file's order. A reply whose parent is missing stands alone.
export function commentThreads(comments: MarkupComment[]): Thread[] {
  const byId = new Map(comments.map((comment) => [comment.id, comment]))
  const rootOf = (comment: MarkupComment) => {
    let current = comment
    const seen = new Set<number>()
    while (current.parent !== null && byId.has(current.parent) && !seen.has(current.id)) {
      seen.add(current.id)
      current = byId.get(current.parent)!
    }
    return current
  }
  const threads = new Map<number, Thread>()
  for (const comment of comments) {
    const root = rootOf(comment)
    const thread = threads.get(root.id) ?? threads.set(root.id, { root, replies: [] }).get(root.id)!
    if (comment !== root) thread.replies.push(comment)
  }
  return [...threads.values()]
}

// What the toggle counts: one per card.
export function markupCount(markup: Pick<DocumentMarkup, "revisions" | "comments">) {
  return markup.revisions.length + commentThreads(markup.comments).length
}

// A thread takes a reply while it starts from a root comment (not a reply whose
// parent is gone) that is not resolved.
const isOpen = (thread: Thread) => thread.root.parent === null && thread.root.resolved !== true
const threadCard = (thread: Thread) => `comment-${thread.root.id}`

// Cards in document order: by part, then paragraph (unknown places last), a
// paragraph's changes before its comments, then as listed. An edit inside a
// comment's own text follows that comment's thread.
function documentOrder(markup: DocumentMarkup): Item[] {
  const last = Number.MAX_SAFE_INTEGER
  const threads = commentThreads(markup.comments).map((thread, index) => ({
    type: "thread" as const,
    thread,
    order: [SCOPES.indexOf(thread.root.scope ?? "body"), thread.root.paragraph ?? last, 1, index, 0],
  }))
  const items: Item[] = [
    ...markup.revisions.map((revision, index): Item => {
      const id = revision.scope === "comment" ? revision.locations.find((place) => place.comment !== undefined)?.comment : undefined
      const owner = id === undefined ? undefined : threads.find(({ thread }) => thread.root.id === id || thread.replies.some((reply) => reply.id === id))
      return {
        type: "change",
        revision,
        order: owner ? [...owner.order.slice(0, 4), 1 + index] : [SCOPES.indexOf(revision.scope), revision.locations.find((place) => place.paragraph !== null)?.paragraph ?? last, 0, 0, index],
      }
    }),
    ...threads,
  ]
  return items.sort((a, b) => {
    const at = a.order.findIndex((value, index) => value !== b.order[index])
    return at === -1 ? 0 : a.order[at]! - b.order[at]!
  })
}

const kindOf = (item: Item): Kind => (item.type === "thread" ? "comments" : item.revision.kind === "format" ? "formatting" : "changes")
const authorsOf = (item: Item) => (item.type === "change" ? [item.revision.author] : [item.thread.root.author, ...item.thread.replies.map((reply) => reply.author)])
const decidable = (revision: MarkupRevision) => revision.status === "supported"
const authorName = (name: string) => name.trim() || "No author named"

// A paragraph's mark travels with a whole inserted, deleted or moved paragraph;
// the reader shows it as a closing pilcrow, which the label says instead.
function wholeParagraph(revision: MarkupRevision) {
  return revision.family !== "paragraph_mark" && revision.kind !== "format" && revision.text.endsWith("¶")
}

export function changeLabel(revision: MarkupRevision) {
  if (revision.kind === "move") return wholeParagraph(revision) ? "Moved paragraph" : "Moved"
  if (revision.kind === "format") {
    const container = revision.locations.find((place) => place.container)?.container
    return container ? CONTAINER[container] : "Formatting changed"
  }
  if (revision.family === "paragraph_mark") return revision.kind === "insert" ? "Paragraph break inserted" : "Paragraph break deleted"
  if (STRUCTURE[revision.family]) return STRUCTURE[revision.family]!
  if (revision.kind === "insert") return wholeParagraph(revision) ? "Inserted paragraph" : "Inserted"
  if (revision.kind === "delete") return wholeParagraph(revision) ? "Deleted paragraph" : "Deleted"
  return "Layout changed"
}

function changeText(revision: MarkupRevision) {
  if (revision.family === "paragraph_mark") return ""
  return wholeParagraph(revision) ? revision.text.slice(0, -1) : revision.text
}

const clip = (text: string, max = 48) => (text.length > max ? `${text.slice(0, max - 1).trimEnd()}…` : text)
const capital = (text: string) => text.charAt(0).toUpperCase() + text.slice(1)
const joined = (names: (string | number)[]) => (names.length > 1 ? `${names.slice(0, -1).join(", ")} and ${names.at(-1)}` : `${names[0]}`)
// The list number Word prints before a paragraph, when it names one (a bullet does not).
const listLabel = (place: Pick<MarkupLocation, "label">) => (place.label && /[\p{L}\p{N}]/u.test(place.label) ? place.label : undefined)

// Where a place is, as the lawyer sees it in Word: by its list number
// ("clause 5", "Article 2"), else by its paragraph's ordinal. A table's, row's,
// cell's or section's change runs from its first paragraph.
function placeName(scope: MarkupRevision["scope"], place: Pick<MarkupLocation, "paragraph" | "label" | "container">) {
  const label = listLabel(place)
  const at = label ? (/^\p{L}{2,}\s/u.test(label) ? label : `clause ${label.replace(/\.$/, "")}`) : place.paragraph !== null ? `paragraph ${place.paragraph}` : undefined
  return at && `${PART[scope]}${place.container ? "from " : ""}${at}`
}

// An edit inside a comment's own text names that comment: who wrote it and,
// unless its first paragraph is quoted below, how it opens.
function commentPlace(place: MarkupLocation, comments: MarkupComment[], quoted: boolean) {
  const comment = comments.find((entry) => entry.id === place.comment)
  if (!comment) return placeName("comment", place) ?? "in a comment"
  const author = comment.author.trim()
  const opening = place.paragraph === 1 && quoted ? "" : clip(comment.text.trim().split("\n")[0]!.trim())
  const within = listLabel(place) || (place.paragraph ?? 0) > 1 ? `, ${placeName("body", place)}` : ""
  return `in a comment ${author ? `by ${author}` : "with no author named"}${opening ? `: “${opening}”` : ""}${within}`
}

// The name of a change's decision group, told apart from every other card's:
// what changed, who changed it and a short excerpt (its paragraph's when the
// change has no text of its own, such as a paragraph break).
export function decisionLabel(revision: MarkupRevision) {
  const author = revision.author.trim()
  const text = changeText(revision).replace(/\s*¶\s*/g, " ").trim() || revision.locations.find((place) => place.context.trim())?.context.trim() || ""
  return `${changeLabel(revision)} ${author ? `by ${author}` : "with no author named"}${text ? `: “${clip(text.replace(/…$/, ""))}”` : ""}`
}

// Plain words for what screening found, once per kind of finding.
export function flagNotes(findings: MarkupFinding[]) {
  const notes: string[] = []
  if (findings.some((finding) => !HIDDEN.has(finding.rule))) notes.push("Contains text that looks like an instruction — doc.haus did not follow it.")
  if (findings.some((finding) => HIDDEN.has(finding.rule))) notes.push("Hidden characters were removed from this text.")
  return notes
}

// Drafts are kept per matter and document with the bytes they were made on
// (`base`). A tracked change keeps its id in a newer version while it is
// unchanged, so a change still listed keeps its draft; the rest start on Hold.
const draftKey = (matterId: string, name: string) =>
  `dh.markup-decisions.v2/${encodeURIComponent(matterId)}/${encodeURIComponent(name)}`

export type StoredDrafts = { base: string; decisions: Record<string, unknown> }

function readStored(key: string): StoredDrafts | undefined {
  const stored = readBrowserJSON(key) as Partial<StoredDrafts> | null
  if (!stored || typeof stored !== "object" || typeof stored.base !== "string" || !stored.decisions || typeof stored.decisions !== "object") return
  return { base: stored.base, decisions: stored.decisions }
}

// The drafts a reading of the file takes from earlier ones: those for changes
// it still lists. When the earlier drafts were made on other bytes, `carried`
// names the ones kept and `dropped` counts decisions on changes now gone.
export function carryDrafts(earlier: StoredDrafts | undefined, markup: Pick<DocumentMarkup, "version" | "revisions">): {
  drafts: Drafts
  carried: string[]
  dropped: number
} {
  if (!earlier) return { drafts: {}, carried: [], dropped: 0 }
  const listed = new Set(markup.revisions.filter(decidable).map((revision) => revision.id))
  const decided = Object.entries(earlier.decisions).filter((entry): entry is [string, "accept" | "reject"] => entry[1] === "accept" || entry[1] === "reject")
  const drafts: Drafts = Object.fromEntries(decided.filter(([id]) => listed.has(id)))
  if (earlier.base === markup.version.sha256) return { drafts, carried: [], dropped: 0 }
  return { drafts, carried: Object.keys(drafts), dropped: decided.length - Object.keys(drafts).length }
}

// Replies are kept the same way, under their own key. A reply stays with its
// thread while the root reads the same and no comment came after the ones it
// answers; carried to other bytes, an approved reply needs approving again.
const replyKey = (matterId: string, name: string) =>
  `dh.markup-replies.v1/${encodeURIComponent(matterId)}/${encodeURIComponent(name)}`

export type StoredReplies = { base: string; replies: Record<string, unknown> }

function readStoredReplies(key: string): StoredReplies | undefined {
  const stored = readBrowserJSON(key) as Partial<StoredReplies> | null
  if (!stored || typeof stored !== "object" || typeof stored.base !== "string" || !stored.replies || typeof stored.replies !== "object") return
  return { base: stored.base, replies: stored.replies }
}

function isReply(value: unknown): value is Reply {
  const reply = value as Partial<Reply> | null
  return !!reply && typeof reply.text === "string" && typeof reply.approved === "boolean" && typeof reply.seen === "number" &&
    Array.isArray(reply.open_items) && reply.open_items.every((item) => typeof item === "string") &&
    typeof reply.root?.author === "string" && typeof reply.root.text === "string"
}

// The replies a reading of the file takes from earlier ones: those whose
// thread is still open, starts from the same comment and has nothing newer.
// From other bytes, `carried` names the ones kept and `dropped` counts the rest.
export function carryReplies(earlier: StoredReplies | undefined, markup: Pick<DocumentMarkup, "version" | "comments">): {
  replies: Replies
  carried: number[]
  dropped: number
} {
  if (!earlier) return { replies: {}, carried: [], dropped: 0 }
  const same = earlier.base === markup.version.sha256
  const open = new Map(commentThreads(markup.comments).filter(isOpen).map((thread) => [thread.root.id, thread]))
  const replies: Replies = {}
  const stored = Object.entries(earlier.replies)
  for (const [id, reply] of stored) {
    const thread = open.get(Number(id))
    if (!thread || !isReply(reply) || reply.root.author !== thread.root.author || reply.root.text !== thread.root.text || reply.seen !== 1 + thread.replies.length) continue
    replies[thread.root.id] = same ? reply : { ...reply, approved: false }
  }
  if (same) return { replies, carried: [], dropped: 0 }
  const carried = Object.keys(replies).map(Number)
  return { replies, carried, dropped: stored.length - carried.length }
}

function without<T>(record: Record<number, T>, ...ids: number[]) {
  const next = { ...record }
  for (const id of ids) delete next[id]
  return next
}

const plural = (count: number, one: string, many = `${one}s`) => `${count} ${count === 1 ? one : many}`

type Carried = { carried: number; dropped: number }

// Plain words for a new reading of a file the lawyer had decided changes or
// drafted replies in.
function carriedNote(name: string, decisions: Carried, replies: Carried, refused: boolean) {
  const decided = decisions.carried + decisions.dropped
  if (!refused && !decided && !replies.carried && !replies.dropped) return undefined
  return [
    refused ? `${name} changed after this list was read, so nothing was applied.` : `${name} changed since these ${decided ? "decisions were made" : "replies were drafted"}.`,
    decisions.carried && `${plural(decisions.carried, "decision")} carried over — check ${decisions.carried === 1 ? "it" : "them"} before applying.`,
    decisions.dropped && `${plural(decisions.dropped, "decision")} could not be carried over and ${decisions.dropped === 1 ? "is" : "are"} back on Hold.`,
    replies.carried && `${plural(replies.carried, "draft reply", "draft replies")} carried over — read ${replies.carried === 1 ? "it" : "them"} again before approving.`,
    replies.dropped && `${plural(replies.dropped, "draft reply", "draft replies")} could not be carried over, as ${replies.dropped === 1 ? "its thread" : "their threads"} changed.`,
    refused && decisions.dropped && !decisions.carried && "Decide the changes again.",
  ].filter(Boolean).join(" ")
}

export default function MarkupPanel({
  matterId,
  name,
  markup,
  loading = false,
  proposalsPending = 0,
  onApplied,
  onRefetch,
  onShowProposals,
}: {
  matterId: string
  name: string
  markup: DocumentMarkup
  // A newer reading is on its way; until then nothing can be applied.
  loading?: boolean
  // doc.haus's own proposals on this document, which are resolved first.
  proposalsPending?: number
  // A new version was saved: the viewer reloads the document and its versions,
  // and shows what the result says to check.
  onApplied: (result: MarkupDecisionResult) => void
  // The service refused the batch: read the file and its markup again.
  onRefetch: () => void
  // doc.haus's own proposals block applying: show them instead of this panel.
  onShowProposals?: () => void
}) {
  const sha256 = markup.version.sha256
  const key = draftKey(matterId, name)
  const repliesKey = replyKey(matterId, name)
  const [initial] = useState(() => ({ decisions: carryDrafts(readStored(key), markup), replies: carryReplies(readStoredReplies(repliesKey), markup) }))
  // The document and bytes the drafts belong to.
  const [reading, setReading] = useState({ key, base: sha256 })
  const [drafts, setDrafts] = useState<Drafts>(initial.decisions.drafts)
  // Decisions carried over from other bytes, until the lawyer looks at each.
  const [carried, setCarried] = useState<string[]>(initial.decisions.carried)
  const [notice, setNotice] = useState(() => carriedNote(name,
    { carried: initial.decisions.carried.length, dropped: initial.decisions.dropped }, { carried: initial.replies.carried.length, dropped: initial.replies.dropped }, false))
  const [replies, setReplies] = useState<Replies>(initial.replies.replies)
  // Replies carried over from other bytes, until the lawyer touches each.
  const [carriedReplies, setCarriedReplies] = useState<number[]>(initial.replies.carried)
  // Threads whose reply is being drafted (or waits its turn), each with its own Cancel.
  const [drafting, setDrafting] = useState<number[]>([])
  const controllers = useRef(new Map<number, AbortController>())
  // Per thread: the lawyer's steer for the next draft, why a draft failed, and
  // the reply Discard or Remove took, which Undo brings back.
  const [steers, setSteers] = useState<Record<number, string>>({})
  const [draftErrors, setDraftErrors] = useState<Record<number, string>>({})
  const [discarded, setDiscarded] = useState<Record<number, { reply: Reply; text: string }>>({})
  const [drafted, setDrafted] = useState<string>()
  // The bytes decisions were just applied to: their drafts are done with.
  const [appliedFrom, setAppliedFrom] = useState<string>()
  const [undo, setUndo] = useState<{ drafts: Drafts; carried: string[]; decision: MarkupDecision; text: string }>()
  const [authors, setAuthors] = useState<string[]>([])
  const [kinds, setKinds] = useState<Kind[]>([])
  // undefined while loading; "" when no name is saved, so the bar asks once.
  const [attorney, setAttorney] = useState<string>()
  const [nameDraft, setNameDraft] = useState("")
  const [confirming, setConfirming] = useState(false)
  const [busy, setBusy] = useState(false)
  const [refusal, setRefusal] = useState<Refusal>()
  const [focus, setFocus] = useState<Focus>()
  const heading = useRef<HTMLHeadingElement>(null)
  const applyButton = useRef<HTMLButtonElement>(null)
  const proposalsButton = useRef<HTMLButtonElement>(null)
  const refusalText = useRef<HTMLParagraphElement>(null)
  const nameInput = useRef<HTMLInputElement>(null)
  const bulkRow = useRef<HTMLDivElement>(null)
  const cards = useRef(new Map<string, HTMLLIElement>())
  const bulkButtons = useRef(new Map<MarkupDecision, HTMLButtonElement>())
  const ids = useId()
  const toast = useToast()

  // A new reading of other bytes keeps the drafts for changes it still lists,
  // and the replies for threads it still holds unchanged, and says so; after
  // this panel's own Apply the decisions start afresh and the replies it did
  // not send stay, without a note. Another document takes its own drafts.
  if (reading.key !== key || reading.base !== sha256) {
    const fresh = reading.key !== key
    const own = appliedFrom === reading.base
    const next = carryDrafts(fresh ? readStored(key) : own ? undefined : { base: reading.base, decisions: drafts }, markup)
    const kept = carryReplies(fresh ? readStoredReplies(repliesKey) : { base: reading.base, replies }, markup)
    setReading({ key, base: sha256 })
    setDrafts(next.drafts)
    setCarried(next.carried)
    setReplies(kept.replies)
    setCarriedReplies(own ? [] : kept.carried)
    setNotice(carriedNote(name, { carried: next.carried.length, dropped: next.dropped },
      own ? { carried: 0, dropped: 0 } : { carried: kept.carried.length, dropped: kept.dropped }, !fresh && refusal?.code === "base_changed"))
    if (fresh) {
      setDrafting([])
      setSteers({})
    }
    setDraftErrors({})
    setDiscarded({})
    setDrafted(undefined)
    setAppliedFrom(undefined)
    setUndo(undefined)
    setRefusal(undefined)
  }

  useEffect(() => {
    if (Object.keys(drafts).length) writeBrowserValue(key, JSON.stringify({ base: reading.base, decisions: drafts }))
    else removeBrowserValue(key)
  }, [key, reading.base, drafts])

  useEffect(() => {
    if (Object.keys(replies).length) writeBrowserValue(repliesKey, JSON.stringify({ base: reading.base, replies }))
    else removeBrowserValue(repliesKey)
  }, [repliesKey, reading.base, replies])

  // A closed panel, or another document, ends its drafts and their model calls.
  useEffect(() => {
    const running = controllers.current
    return () => {
      for (const controller of running.values()) controller.abort()
      running.clear()
    }
  }, [repliesKey])

  useEffect(() => {
    let cancelled = false
    getPreferences().then(
      (prefs) => !cancelled && setAttorney(prefs.attorney?.trim() ?? ""),
      () => !cancelled && setAttorney(""),
    )
    return () => {
      cancelled = true
    }
  }, [])

  // Focus waits for the render that shows its target: a refused change's
  // chosen decision or a refused reply's message, a reply's next control after
  // an action on it, the Draft replies row, or Apply once the file has been
  // read again (the proposals button or the message while Apply cannot be used).
  useEffect(() => {
    if (!focus || busy) return
    const card = typeof focus === "object" ? cards.current.get(focus.card) : undefined
    if (focus === "bulk") {
      bulkRow.current?.querySelector<HTMLElement>("button:not(:disabled)")?.focus()
    } else if (card && focus !== "apply" && focus.reply) {
      card.querySelector<HTMLElement>("[data-reply-focus]:not(:disabled)")?.focus()
    } else if (card) {
      card.scrollIntoView?.({ block: "center" })
      const control = card.querySelector<HTMLElement>('[role="radio"][aria-checked="true"]') ?? card.querySelector<HTMLElement>(".version-error")
      control?.focus({ preventScroll: true })
    } else {
      if (loading) return
      const target = [applyButton.current, proposalsButton.current, refusalText.current].find((element) => element && !(element as HTMLButtonElement).disabled) ?? heading.current
      target?.focus()
    }
    setFocus(undefined)
  })

  const items = useMemo(() => documentOrder(markup), [markup])
  const byAuthor = items.filter((item) => !authors.length || authorsOf(item).some((author) => authors.includes(author)))
  const shown = byAuthor.filter((item) => !kinds.length || kinds.includes(kindOf(item)))
  const shownChanges = shown.flatMap((item) => (item.type === "change" && decidable(item.revision) ? [item.revision] : []))
  const listed = markup.revisions.filter(decidable)
  const accepted = listed.filter((revision) => drafts[revision.id] === "accept").length
  const rejected = listed.filter((revision) => drafts[revision.id] === "reject").length
  const held = listed.length - accepted - rejected
  const elsewhere = markup.revisions.length - listed.length
  const savedName = attorney?.trim()
  const actor = savedName || nameDraft.trim()
  // Open threads in document order. One awaits a reply while its latest
  // comment is not the saved attorney's; only those are offered a draft.
  const open = items.flatMap((item) => (item.type === "thread" && isOpen(item.thread) ? [item.thread] : []))
  const awaiting = (thread: Thread) => !savedName || (thread.replies.at(-1) ?? thread.root).author.trim().toLowerCase() !== savedName.toLowerCase()
  const undrafted = open.filter((thread) => awaiting(thread) && !replies[thread.root.id] && !drafting.includes(thread.root.id))
  const approved = open.filter((thread) => replies[thread.root.id]?.approved)
  const unapproved = open.filter((thread) => replies[thread.root.id] && !replies[thread.root.id]!.approved).length
  const decided = accepted > 0 || rejected > 0
  const blocked = proposalsPending > 0
    ? `Accept or reject doc.haus's ${plural(proposalsPending, "proposed change")} on this document first.`
    : !decided && !approved.length ? (listed.length ? "Mark a change Accept or Reject, or approve a reply." : "Approve a reply first.") : undefined
  const canApply = !blocked && !busy && !loading && attorney !== undefined
  const proposalsFirst = proposalsPending > 0 || refusal?.code === "proposals_pending"
  // A refusal is shown on the cards it names; otherwise, or while they are
  // filtered out, beside Apply. The proposals line already says what a
  // proposals refusal says once the proposals are known.
  const shownIds = new Set(shown.flatMap((item) => (item.type === "change" ? [item.revision.id] : [])))
  const shownThreads = new Set(shown.flatMap((item) => (item.type === "thread" ? [item.thread.root.id] : [])))
  const onCards = !!refusal && (refusal.ids.some((id) => shownIds.has(id)) || (refusal.thread !== undefined && shownThreads.has(refusal.thread)))
  const refusalShown = refusal && !onCards && !(refusal.code === "proposals_pending" && proposalsPending > 0) ? refusal.message : undefined

  function decide(changes: MarkupRevision[], decision: MarkupDecision, bulk?: string) {
    setConfirming(false)
    const next = { ...drafts }
    for (const revision of changes) {
      if (decision === "hold") delete next[revision.id]
      else next[revision.id] = decision
    }
    const decided = new Set(changes.map((revision) => revision.id))
    const unchanged = changes.every((revision) => (drafts[revision.id] ?? "hold") === decision)
    setDrafts(next)
    setCarried(carried.filter((id) => !decided.has(id)))
    setUndo(bulk && !unchanged ? { drafts, carried, decision, text: bulk } : undefined)
    // A change the refusal named is settled once the lawyer decides it again.
    const left = refusal?.ids.filter((id) => !decided.has(id)) ?? []
    setRefusal(left.length ? { ...refusal!, ids: left } : undefined)
  }

  function undoBulk() {
    if (!undo) return
    setDrafts(undo.drafts)
    setCarried(undo.carried)
    setUndo(undefined)
    setConfirming(false)
    bulkButtons.current.get(undo.decision)?.focus()
  }

  // Sets, or with no reply removes, one thread's reply. Touching it settles a
  // carried-over mark, an earlier failure and the refusal that named it.
  function setReply(thread: Thread, reply: Reply | undefined) {
    const id = thread.root.id
    setReplies((current) => (reply ? { ...current, [id]: reply } : without(current, id)))
    setCarriedReplies((current) => current.filter((entry) => entry !== id))
    setDraftErrors((current) => without(current, id))
    setDiscarded((current) => without(current, id))
    setConfirming(false)
    if (refusal?.thread === id) setRefusal(undefined)
  }

  // A new reply to `thread`: as read now, with nothing approved.
  const replyTo = (thread: Thread, text: string, items: string[] = []): Reply =>
    ({ text, approved: false, open_items: items, root: { author: thread.root.author, text: thread.root.text }, seen: 1 + thread.replies.length })

  // Whether focus is nowhere, or still inside `within` (a control that went).
  const focusFree = (within: Element | null | undefined) => {
    const active = document.activeElement
    return !active || active === document.body || !!within?.contains(active)
  }

  // Drafts a reply for each thread, three at a time. Each draft has its own
  // Cancel, and one cancelled before its turn is skipped. The draft sees the
  // decisions as they stand now. A file that changed refuses every draft, so
  // the rest stop and the file is read again. Resolves with the threads that
  // got a draft.
  async function draftReplies(targets: Thread[]) {
    const queue = targets.filter((thread) => !controllers.current.has(thread.root.id))
    const ids = queue.map((thread) => thread.root.id)
    for (const id of ids) controllers.current.set(id, new AbortController())
    setDrafting((current) => [...current, ...ids])
    setDraftErrors((current) => without(current, ...ids))
    setDiscarded((current) => without(current, ...ids))
    setDrafted(undefined)
    const decisions = listed.flatMap((revision) => (drafts[revision.id] ? [{ id: revision.id, decision: drafts[revision.id]! }] : []))
    const done = new Set<Thread>()
    const draftOne = async (thread: Thread) => {
      const id = thread.root.id
      const controller = controllers.current.get(id)
      if (!controller || controller.signal.aborted) return
      const steer = steers[id]?.trim()
      try {
        const draft = await draftMarkupReply(matterId, name, { base_sha256: sha256, thread: id, decisions, ...(steer ? { steer } : {}) }, controller.signal)
        if (controller.signal.aborted) return
        setReplies((current) => ({ ...current, [id]: replyTo(thread, draft.reply, draft.open_items) }))
        setCarriedReplies((current) => current.filter((entry) => entry !== id))
        done.add(thread)
      } catch (cause) {
        if (controller.signal.aborted) return
        const stale = cause instanceof MarkupRefusedError && cause.code === "base_changed"
        setDraftErrors((current) => ({
          ...current,
          [id]: stale ? `${name} changed after this list was read. Draft the reply again once it is read again.`
            : cause instanceof MarkupRefusedError ? cause.message : humanizeError(cause, "The reply was not drafted. Try again, or write the reply."),
        }))
        if (stale) cancelDrafts()
        // The file or the thread moved on: read them again.
        if (cause instanceof MarkupRefusedError && cause.status === 409 && cause.code !== "draft_unreadable") onRefetch()
      } finally {
        if (controllers.current.get(id) === controller) {
          controllers.current.delete(id)
          setDrafting((current) => current.filter((entry) => entry !== id))
        }
      }
    }
    let next = 0
    const worker = async () => {
      while (next < queue.length) await draftOne(queue[next++]!)
    }
    await Promise.all(Array.from({ length: Math.min(POOL, queue.length) }, worker))
    if (done.size) setDrafted(`${plural(done.size, "draft reply", "draft replies")} ready — read ${done.size === 1 ? "it" : "each one"}, then approve it.`)
    return queue.filter((thread) => done.has(thread))
  }

  function cancelDrafts(ids = [...controllers.current.keys()]) {
    for (const id of ids) {
      controllers.current.get(id)?.abort()
      controllers.current.delete(id)
    }
    setDrafting((current) => current.filter((entry) => !ids.includes(entry)))
  }

  // Every thread awaiting a reply that has none: focus stays on the row (its
  // Cancel), then goes to the first draft if the lawyer is still there.
  async function draftAll() {
    setConfirming(false)
    setFocus("bulk")
    const done = await draftReplies(undrafted)
    if (done.length && focusFree(bulkRow.current)) setFocus({ card: threadCard(done[0]!), reply: true })
  }

  // One thread's actions; each leaves focus on the reply's next control.
  function replyActions(thread: Thread, reply: Reply | undefined): ReplyActions {
    const card = threadCard(thread)
    const next = () => setFocus({ card, reply: true })
    return {
      draft: async () => {
        setConfirming(false)
        next()
        const done = await draftReplies([thread])
        if (done.length && focusFree(cards.current.get(card))) next()
      },
      cancel: () => {
        cancelDrafts([thread.root.id])
        next()
      },
      write: () => {
        setReply(thread, replyTo(thread, ""))
        next()
      },
      change: (text: string) => reply && setReply(thread, { ...reply, text }),
      approve: () => {
        if (!reply) return
        setReply(thread, { ...reply, approved: true })
        next()
      },
      edit: () => {
        if (!reply) return
        setReply(thread, { ...reply, approved: false })
        next()
      },
      discard: () => {
        if (!reply) return
        setReply(thread, undefined)
        setDiscarded((current) => ({ ...current, [thread.root.id]: { reply, text: reply.approved ? "Reply removed." : "Draft discarded." } }))
        next()
      },
      undo: () => {
        const earlier = discarded[thread.root.id]
        if (!earlier) return
        setReply(thread, earlier.reply)
        next()
      },
      steer: (text: string) => setSteers((current) => ({ ...current, [thread.root.id]: text })),
    }
  }

  const toggle =<T,>(list: T[], value: T) => (list.includes(value) ? list.filter((entry) => entry !== value) : [...list, value])

  function review() {
    setRefusal(undefined)
    setUndo(undefined)
    if (!actor) {
      setRefusal({ message: "Add your name so it is recorded with these decisions.", ids: [] })
      nameInput.current?.focus()
      return
    }
    setConfirming(true)
  }

  function cancel() {
    setConfirming(false)
    // The Apply button is rendered again once the confirmation goes.
    requestAnimationFrame(() => applyButton.current?.focus())
  }

  // What a refused Apply shows, and where focus goes: the first change it
  // names, with the filters cleared if they hide any it names, or Apply.
  function refused(cause: unknown): Focus {
    if (!(cause instanceof MarkupRefusedError)) {
      setRefusal({ message: humanizeError(cause, "The decisions could not be applied. Try again."), ids: [] })
      return "apply"
    }
    // The file or its changes moved on: read them again before another try.
    if (cause.status === 409) onRefetch()
    const named = cause.ids.filter((id) => markup.revisions.some((revision) => revision.id === id))
    const thread = open.find((entry) => entry.root.id === cause.thread)
    // The panel reads the file again itself; there is nothing to refresh.
    const message = cause.code === "base_changed" ? `${name} changed after this list was read, so nothing was applied.` : cause.message
    setRefusal({ message, code: cause.code, ids: named, thread: thread?.root.id })
    if (!named.length && !thread) return "apply"
    if (!named.every((id) => shownIds.has(id)) || (thread && !shownThreads.has(thread.root.id))) {
      setAuthors([])
      setKinds([])
    }
    return { card: named.length ? named[0]! : threadCard(thread!) }
  }

  async function apply() {
    setBusy(true)
    setRefusal(undefined)
    let after: Focus | undefined
    try {
      let recorded = savedName
      if (!recorded) {
        try {
          recorded = (await savePreferences({ attorney: actor })).attorney?.trim() || actor
          setAttorney(recorded)
        } catch (cause) {
          // The name still travels with the decisions; only one that cannot
          // needs it saved first.
          if (!actorSent(actor)) {
            setConfirming(false)
            setRefusal({ message: `Your name could not be saved, so it cannot be recorded with these decisions. ${humanizeError(cause)}`, ids: [] })
            after = "apply"
            return
          }
          recorded = actor
          toast("warning", "Your name could not be saved in Settings; it is still recorded with these decisions.")
        }
      }
      // Only replies the lawyer approved are sent; drafts stay drafts.
      const sent = approved.map((thread) => ({ thread: thread.root.id, text: replies[thread.root.id]!.text.trim() }))
      const result = await applyMarkupDecisions(matterId, name, {
        base_sha256: sha256,
        decisions: listed.map((revision) => ({ id: revision.id, decision: drafts[revision.id] ?? "hold" })),
        ...(sent.length ? { replies: sent } : {}),
      }, recorded)
      removeBrowserValue(key)
      setAppliedFrom(sha256)
      setCarried([])
      setReplies((current) => without(current, ...sent.map((reply) => reply.thread)))
      setCarriedReplies([])
      setDrafted(undefined)
      setNotice(undefined)
      setUndo(undefined)
      setConfirming(false)
      const counts = result.counts ?? {}
      const added = result.replies?.length ?? sent.length
      const tally = [
        decided && [
          `${counts.accepted ?? accepted} accepted`,
          `${counts.rejected ?? rejected} rejected`,
          `${counts.held ?? held} held`,
        ].join(", "),
        added > 0 && `${plural(added, "reply", "replies")} added`,
      ].filter(Boolean).join("; ")
      // The tally names the held changes; the version's own notice about
      // tracked changes left in it would only repeat them.
      toast("success", `${result.version ? `Saved as version ${result.version.number}` : "Saved a new version"}: ${tally}.`)
      if (counts.consumed) toast("info", `${plural(counts.consumed, "change was", "changes were")} resolved along with another change.`)
      const broken = result.consistency?.status === "checked" ? result.consistency.findings.length : 0
      if (broken) toast("info", `${plural(broken, "reference or defined term", "references or defined terms")} to check.`)
      onApplied(result)
    } catch (cause) {
      setConfirming(false)
      after = refused(cause)
    } finally {
      setBusy(false)
      if (after) setFocus(after)
      else heading.current?.focus()
    }
  }

  const saved = [
    ...(decided ? [`${plural(accepted, "change")} accepted`, `${rejected} rejected`] : []),
    ...(approved.length ? [`${plural(approved.length, "reply", "replies")} added as ${actor}`] : []),
  ]
  const confirmText = `This saves a new version of ${name} with ${joined(saved)}.` +
    `${held ? ` ${plural(held, "held change stays", "held changes stay")} as tracked ${held === 1 ? "change" : "changes"}.` : ""}` +
    `${unapproved ? ` ${plural(unapproved, "draft reply is", "draft replies are")} not approved, so ${unapproved === 1 ? "it is" : "they are"} not added.` : ""}` +
    " The current version stays in the history."

  return (
    <aside id="document-markup" className="changes-panel markup-panel" aria-labelledby={`${ids}-title`}>
      <div className="changes-head markup-head">
        <h3 id={`${ids}-title`} ref={heading} tabIndex={-1}>
          Their changes <span className="compare-count">{items.length}</span>
        </h3>
        <p>
          Tracked changes and comments in this Word file. Choose Accept, Reject or Hold for each change{open.length > 0 && " and approve any replies to their comments"}, then
          apply them together. Held changes stay as tracked changes.
        </p>
        {markup.authors.length > 0 && (
          <div className="markup-chips" role="group" aria-label="Reviewers">
            {markup.authors.map((author) => (
              <button type="button" key={author.name} className="markup-chip" aria-pressed={authors.includes(author.name)} onClick={() => setAuthors(toggle(authors, author.name))}>
                {authorName(author.name)}
                <span className="markup-chip-count">{author.revisions + author.comments}</span>
              </button>
            ))}
          </div>
        )}
        <div className="markup-chips" role="group" aria-label="Kinds">
          {KINDS.map((kind) => {
            const count = byAuthor.filter((item) => kindOf(item) === kind.value).length
            return (
              <button type="button" key={kind.value} className="markup-chip" aria-pressed={kinds.includes(kind.value)} disabled={!count && !kinds.includes(kind.value)}
                onClick={() => setKinds(toggle(kinds, kind.value))}>
                {kind.label}
                <span className="markup-chip-count">{count}</span>
              </button>
            )
          })}
        </div>
        {shownChanges.length > 0 && (
          <div className="markup-bulk" role="group" aria-labelledby={`${ids}-bulk`}>
            <span id={`${ids}-bulk`}>
              Set {shownChanges.length === listed.length ? `all ${plural(shownChanges.length, "change")}` : `the ${plural(shownChanges.length, "shown change")}`} to
            </span>
            {DECISIONS.map((decision) => (
              <button type="button" key={decision.value} className="icon-btn" disabled={busy}
                ref={(node) => { if (node) bulkButtons.current.set(decision.value, node); else bulkButtons.current.delete(decision.value) }}
                onClick={() => decide(shownChanges, decision.value, `${plural(shownChanges.length, "change")} set to ${decision.label}.`)}>
                {decision.label}
              </button>
            ))}
          </div>
        )}
        {undo && (
          <p role="status">
            {undo.text}{" "}
            <button type="button" className="linklike" disabled={busy} onClick={undoBulk}>Undo</button>
          </p>
        )}
        {(undrafted.length > 0 || drafting.length > 0 || drafted) && (
          <div className="markup-bulk markup-draft-all" ref={bulkRow}>
            {drafting.length > 0 ? (
              <>
                <span role="status">Drafting {plural(drafting.length, "reply", "replies")}…</span>
                <button type="button" className="icon-btn" onClick={() => cancelDrafts()}>Cancel drafting</button>
              </>
            ) : (
              <>
                {undrafted.length > 0 && (
                  <button type="button" className="icon-btn is-primary" disabled={busy || loading} onClick={draftAll}>
                    Draft replies ({undrafted.length})
                  </button>
                )}
                {drafted && <span role="status">{drafted}</span>}
              </>
            )}
          </div>
        )}
      </div>

      {shown.length === 0 ? (
        <p className="muted markup-empty">
          Nothing matches these filters.{" "}
          <button type="button" className="linklike" onClick={() => { setAuthors([]); setKinds([]) }}>Show everything</button>
        </p>
      ) : (
        <ol className="markup-list" aria-label={`Changes and comments in ${name}, in document order`} aria-busy={loading || undefined}>
          {shown.map((item) =>
            item.type === "change" ? (
              <ChangeCard key={item.revision.id} revision={item.revision} comments={markup.comments} decision={drafts[item.revision.id] ?? "hold"} disabled={busy}
                carried={carried.includes(item.revision.id)}
                failure={refusal?.ids.includes(item.revision.id) ? refusal.message : undefined}
                cardRef={(node) => { if (node) cards.current.set(item.revision.id, node); else cards.current.delete(item.revision.id) }}
                onDecide={(decision) => decide([item.revision], decision)}
                onHold={() => {
                  decide([item.revision], "hold")
                  setFocus({ card: item.revision.id })
                }} />
            ) : (
              <ThreadCard key={threadCard(item.thread)} thread={item.thread} open={isOpen(item.thread)} awaiting={awaiting(item.thread)}
                reply={replies[item.thread.root.id]} drafting={drafting.includes(item.thread.root.id)} steer={steers[item.thread.root.id] ?? ""}
                carried={carriedReplies.includes(item.thread.root.id)} undo={discarded[item.thread.root.id]?.text} disabled={busy} canDraft={!busy && !loading}
                failure={refusal?.thread === item.thread.root.id ? refusal.message : draftErrors[item.thread.root.id]}
                cardRef={(node) => { if (node) cards.current.set(threadCard(item.thread), node); else cards.current.delete(threadCard(item.thread)) }}
                actions={replyActions(item.thread, replies[item.thread.root.id])} />
            ),
          )}
        </ol>
      )}

      {(listed.length > 0 || open.length > 0) && (
        <div className="markup-apply">
          {notice && <p className="markup-hint" role="status">{notice}</p>}
          {(listed.length > 0 || approved.length > 0) && (
            <p className="markup-tally">
              {listed.length > 0 && `${accepted} to accept · ${rejected} to reject · ${held} held`}
              {listed.length > 0 && approved.length > 0 && " · "}
              {approved.length > 0 && `${plural(approved.length, "reply", "replies")} to add`}
              {listed.length > 0 && elsewhere > 0 && <span> · {elsewhere} to resolve in Word</span>}
            </p>
          )}
          {attorney === undefined ? null : savedName ? (
            <p className="markup-actor">Recorded as <strong>{savedName}</strong></p>
          ) : (
            <div className="markup-name">
              <label htmlFor={`${ids}-name`}>Your name, recorded with these decisions</label>
              <input id={`${ids}-name`} ref={nameInput} type="text" autoComplete="name" value={nameDraft} maxLength={120} disabled={busy}
                aria-describedby={`${ids}-name-hint`} onChange={(event) => { setNameDraft(event.target.value); setConfirming(false) }} />
              <small id={`${ids}-name-hint`}>Saved as the attorney name in Settings → Drafting.</small>
            </div>
          )}
          {blocked && <p id={`${ids}-blocked`} className="markup-blocked">{blocked}</p>}
          {proposalsFirst && onShowProposals && (
            <button type="button" ref={proposalsButton} className="icon-btn" onClick={onShowProposals}>
              Show doc.haus's proposals
            </button>
          )}
          {confirming ? (
            <div className="version-confirm" role="group" aria-label="Apply decisions" onKeyDown={(event) => {
              if (event.key !== "Escape" || busy) return
              event.preventDefault()
              event.stopPropagation()
              cancel()
            }}>
              <p id={`${ids}-confirm`}>{confirmText}</p>
              {!savedName && <p className="markup-actor">Recorded as <strong>{actor}</strong></p>}
              <span className="upload-actions">
                <button type="button" className="icon-btn upload-choice-lead" autoFocus aria-describedby={`${ids}-confirm`} disabled={busy || loading} onClick={apply}>
                  {busy ? "Applying…" : "Apply decisions"}
                </button>
                <button type="button" className="icon-btn" disabled={busy} onClick={cancel}>
                  Cancel
                </button>
              </span>
            </div>
          ) : (
            <button type="button" ref={applyButton} className="primary markup-apply-button" disabled={!canApply} aria-describedby={blocked ? `${ids}-blocked` : undefined} onClick={review}>
              Apply decisions…
            </button>
          )}
          {refusalShown && <p ref={refusalText} tabIndex={-1} className="version-error" role="alert">{refusalShown}</p>}
        </div>
      )}
    </aside>
  )
}

// Above the page after an Apply (#198): the references and defined terms its
// decisions broke, each as the check worded it, or one line when they could
// not be checked; nothing for a clean check. It goes with the next action.
export function ConsistencyNotice({ consistency }: { consistency: NonNullable<MarkupDecisionResult["consistency"]> }) {
  const id = useId()
  if (consistency.status === "not_checked") return <p role="status" className="markup-consistency-unchecked">References and defined terms were not checked for this version.</p>
  if (!consistency.findings.length) return null
  return (
    <div role="status" className="markup-consistency">
      <strong id={id}>Check these references and defined terms after your decisions</strong>
      <ul aria-labelledby={id}>
        {consistency.findings.map((finding, index) => <li key={index}>{finding.detail}</li>)}
      </ul>
    </div>
  )
}

function Flags({ findings }: { findings: MarkupFinding[] }) {
  const notes = flagNotes(findings)
  if (!notes.length) return null
  return (
    <div className="markup-flag">
      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
        <path d="M12 3 2 21h20L12 3z" />
        <path d="M12 10v5M12 18h.01" />
      </svg>
      <span>{notes.map((note) => <span key={note}>{note}</span>)}</span>
    </div>
  )
}

function Places({ revision, comments }: { revision: MarkupRevision; comments: MarkupComment[] }) {
  const text = changeText(revision)
  const places = revision.locations.filter((place) => place.paragraph !== null || place.context)
  if (!places.length) return null
  // The preview repeats a whole inserted or moved paragraph; it is left out
  // then. A change with no text of its own (a paragraph break, a table row)
  // shows its paragraph's.
  const preview = (context: string) => {
    const head = context.replace(/…$/, "")
    return head && (!text || (!text.startsWith(head) && !head.startsWith(text))) ? context : ""
  }
  if (revision.kind === "move" && places.length > 1 && places.every((place) => place.paragraph !== null && !preview(place.context))) {
    // A move says where it left and where it landed when the reader knows each
    // place's side.
    const from = places.filter((place) => place.side === "from")
    const to = places.filter((place) => place.side === "to")
    const side = (list: MarkupLocation[]) =>
      list.length > 1 && !list.some(listLabel) ? `paragraphs ${joined(list.map((place) => place.paragraph!))}` : joined(list.map((place) => placeName("body", place)!))
    if (from.length && to.length && from.length + to.length === places.length)
      return <p className="markup-place">Moved from {PART[revision.scope]}{side(from)} to {side(to)}</p>
    // Otherwise its places come in document order, which does not say which
    // one it left; the document shows that.
    const labelled = places.some(listLabel)
    const where = joined(places.map((place) => (labelled ? placeName(revision.scope, place)! : place.paragraph!)))
    return <p className="markup-place">Old and new places: {labelled ? where : `${PART[revision.scope]}paragraphs ${where}`}</p>
  }
  return (
    <>
      {places.map((place, index) => {
        const context = preview(place.context)
        const named = revision.scope === "comment" ? commentPlace(place, comments, !!context) : placeName(revision.scope, place)
        // A move whose places are quoted says which one each is.
        const where = named && revision.scope !== "comment" && place.side ? `${place.side} ${named}` : named
        return (
          <p className="markup-place" key={index}>
            {where && <span>{capital(where)}</span>}
            {context && <q>{context}</q>}
          </p>
        )
      })}
    </>
  )
}

// One change: Accept, Reject and Hold are one radio group, a single Tab stop
// whose arrow keys move and choose. A refused change carries the refusal and a
// one-click Hold.
function ChangeCard({ revision, comments, decision, disabled, carried, failure, cardRef, onDecide, onHold }: {
  revision: MarkupRevision
  comments: MarkupComment[]
  decision: MarkupDecision
  disabled: boolean
  carried: boolean
  failure?: string
  cardRef: (node: HTMLLIElement | null) => void
  onDecide: (decision: MarkupDecision) => void
  onHold: () => void
}) {
  const id = useId()
  const label = changeLabel(revision)
  const text = changeText(revision)
  const removes = revision.kind === "delete"
  function step(event: KeyboardEvent<HTMLButtonElement>, index: number) {
    const last = DECISIONS.length - 1
    const moves: Record<string, number | undefined> = { ArrowRight: index + 1, ArrowDown: index + 1, ArrowLeft: index - 1, ArrowUp: index - 1, Home: 0, End: last }
    const to = moves[event.key]
    if (to === undefined) return
    event.preventDefault()
    const next = to > last ? 0 : to < 0 ? last : to
    onDecide(DECISIONS[next]!.value)
    event.currentTarget.parentElement?.querySelectorAll<HTMLElement>('[role="radio"]')[next]?.focus()
  }
  return (
    <li ref={cardRef} className={`change-card markup-card is-${decision}${failure ? " is-refused" : ""}`}>
      <div className="markup-card-head">
        <span className={`markup-kind is-${revision.kind}`}>{label}</span>
        <span className="change-author">{authorName(revision.author)}</span>
        {revision.date && <time dateTime={revision.date}>{formatDateTime(revision.date)}</time>}
        {carried && <span className="version-tag">Carried over</span>}
      </div>
      {text && <div className={removes ? "change-old" : revision.kind === "insert" ? "change-new" : "markup-text"}>{text}</div>}
      {revision.family === "paragraph_mark" && revision.kind === "delete" && <p className="markup-hint">Accepting joins this paragraph with the next one.</p>}
      {revision.kind === "format" && decidable(revision) && <p className="markup-hint">Accept keeps the new formatting; Reject restores the old.</p>}
      <Places revision={revision} comments={comments} />
      <Flags findings={revision.findings} />
      {decidable(revision) ? (
        <div className="markup-decision" role="radiogroup" aria-label={decisionLabel(revision)} aria-describedby={failure ? `${id}-refused` : undefined}>
          {DECISIONS.map((option, index) => (
            <button type="button" role="radio" key={option.value} className={`markup-choice is-${option.value}`} aria-checked={decision === option.value}
              tabIndex={decision === option.value ? 0 : -1} disabled={disabled} onClick={() => onDecide(option.value)} onKeyDown={(event) => step(event, index)}>
              {option.label}
            </button>
          ))}
        </div>
      ) : (
        <p className="markup-word"><strong>Resolve in Word.</strong> doc.haus cannot apply this change safely, so it stays as a tracked change.</p>
      )}
      {failure && (
        <p className="version-error" role="alert" tabIndex={-1}>
          <span id={`${id}-refused`}>{failure}</span>
          {decision !== "hold" && (
            <>
              {" "}
              <button type="button" className="linklike" disabled={disabled} onClick={onHold}>Hold this change</button>
            </>
          )}
        </p>
      )}
    </li>
  )
}

function CommentBody({ comment }: { comment: MarkupComment }) {
  return (
    <>
      <div className="markup-card-head">
        <span className="change-author">{authorName(comment.author)}</span>
        {comment.date && <time dateTime={comment.date}>{formatDateTime(comment.date)}</time>}
      </div>
      <p className="markup-comment-text">{comment.text}</p>
      <Flags findings={comment.findings} />
    </>
  )
}

// A thread, read-only, with the reply area while it is open.
function ThreadCard({ thread, open, cardRef, ...area }: { thread: Thread; open: boolean; cardRef: (node: HTMLLIElement | null) => void } & Omit<ReplyAreaProps, "thread">) {
  const { root, replies } = thread
  return (
    <li ref={cardRef} className={`change-card markup-card markup-thread${area.reply?.approved ? " is-approved" : ""}${area.failure ? " is-refused" : ""}`}>
      <div className="markup-card-head">
        <span className="markup-kind is-comment">Comment</span>
        {root.resolved === true && <span className="version-tag is-viewing">Resolved</span>}
      </div>
      {root.quote && <blockquote className="markup-quote">{root.quote}</blockquote>}
      <CommentBody comment={root} />
      {replies.length > 0 && (
        <ol className="markup-replies" aria-label={`${plural(replies.length, "reply", "replies")}`}>
          {replies.map((reply) => (
            <li key={reply.id}>
              <CommentBody comment={reply} />
            </li>
          ))}
        </ol>
      )}
      {root.paragraph !== null && <p className="markup-place"><span>{capital(placeName(root.scope ?? "body", root)!)}</span></p>}
      {open && <ReplyArea thread={thread} {...area} />}
    </li>
  )
}

// What a thread's reply area can do; the panel focuses its next control after each.
type ReplyActions = Record<"draft" | "cancel" | "write" | "approve" | "edit" | "discard" | "undo", () => void> & {
  change: (text: string) => void
  steer: (text: string) => void
}

type ReplyAreaProps = {
  thread: Thread
  // Offered a draft: its latest comment is not the lawyer's.
  awaiting: boolean
  reply?: Reply
  drafting: boolean
  steer: string
  carried: boolean
  // What Discard or Remove did, with Undo beside it.
  undo?: string
  failure?: string
  disabled: boolean
  canDraft: boolean
  actions: ReplyActions
}

// One open thread's reply: offered (Draft reply, Write a reply, an optional
// steer), drafting (Cancel), a draft to edit (Approve once nothing is left in
// brackets, Redraft, Discard) or approved (Edit, Remove). The control the
// panel focuses after an action carries data-reply-focus. A thread flagged for
// planted instructions repeats its flag above the draft.
function ReplyArea({ thread, awaiting, reply, drafting, steer, carried, undo, failure, disabled, canDraft, actions }: ReplyAreaProps) {
  const id = useId()
  const [steering, setSteering] = useState(steer !== "")
  const planted = [thread.root, ...thread.replies].flatMap((comment) => comment.findings).filter((finding) => !HIDDEN.has(finding.rule))
  const hold = !reply || reply.approved ? undefined : !reply.text.trim() ? "Write the reply first." : PLACEHOLDER.test(reply.text) ? "Fill the bracketed parts first." : undefined
  const steerToggle = awaiting && (
    <button type="button" className="linklike" aria-expanded={steering} onClick={() => setSteering(!steering)}>Steer…</button>
  )
  const steerField = awaiting && steering && (
    <div className="markup-steer">
      <label htmlFor={`${id}-steer`}>Direction for the draft</label>
      <input id={`${id}-steer`} type="text" value={steer} maxLength={STEER_MAX} disabled={disabled} placeholder="For example: agree, or hold firm at 30 days"
        onChange={(event) => actions.steer(event.target.value)}
        onKeyDown={(event) => {
          if (event.key !== "Enter" || !canDraft) return
          event.preventDefault()
          actions.draft()
        }} />
    </div>
  )
  return (
    <div className="markup-reply" role="group" aria-label={`Your reply to ${authorName(thread.root.author)}`}>
      {drafting ? (
        <div className="markup-reply-actions">
          <span className="markup-reply-status">Drafting a reply…</span>
          <button type="button" className="icon-btn" data-reply-focus onClick={actions.cancel}>Cancel</button>
        </div>
      ) : !reply ? (
        <>
          {undo && (
            <p className="markup-hint" role="status">
              {undo}{" "}
              <button type="button" className="linklike" disabled={disabled} onClick={actions.undo}>Undo</button>
            </p>
          )}
          <div className="markup-reply-actions">
            {awaiting && (
              <button type="button" className="icon-btn is-primary" data-reply-focus disabled={!canDraft} onClick={actions.draft}>Draft reply</button>
            )}
            <button type="button" className="linklike" data-reply-focus={awaiting ? undefined : true} disabled={disabled} onClick={actions.write}>Write a reply</button>
            {steerToggle}
          </div>
          {steerField}
        </>
      ) : reply.approved ? (
        <>
          <p className="markup-reply-approved">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
              <path d="M5 12.5 10 17 19 7" />
            </svg>
            Reply approved, added when you apply
          </p>
          <p className="markup-reply-final">{reply.text}</p>
          <div className="markup-reply-actions">
            <button type="button" className="icon-btn" data-reply-focus disabled={disabled} onClick={actions.edit}>Edit</button>
            <button type="button" className="linklike" disabled={disabled} onClick={actions.discard}>Remove</button>
          </div>
        </>
      ) : (
        <>
          {planted.length > 0 && <Flags findings={planted} />}
          <label className="markup-reply-label" htmlFor={`${id}-text`}>
            Your reply
            {carried && <span className="version-tag">Carried over</span>}
          </label>
          <textarea id={`${id}-text`} className="markup-reply-text" data-reply-focus value={reply.text} rows={4} maxLength={REPLY_MAX} disabled={disabled}
            aria-describedby={[reply.open_items.length ? `${id}-items` : "", hold ? `${id}-hold` : ""].filter(Boolean).join(" ") || undefined}
            onChange={(event) => actions.change(event.target.value)} />
          {reply.open_items.length > 0 && (
            <div className="markup-open-items" id={`${id}-items`}>
              <p>Check before sending</p>
              <ul>
                {reply.open_items.map((item, index) => <li key={index}>{item}</li>)}
              </ul>
            </div>
          )}
          {hold && <p id={`${id}-hold`} className="markup-hint">{hold}</p>}
          <div className="markup-reply-actions">
            <button type="button" className="icon-btn is-primary" disabled={disabled || !!hold} aria-describedby={hold ? `${id}-hold` : undefined} onClick={actions.approve}>
              Approve
            </button>
            {awaiting && <button type="button" className="icon-btn" disabled={!canDraft} onClick={actions.draft}>Redraft</button>}
            <button type="button" className="linklike" disabled={disabled} onClick={actions.discard}>Discard</button>
            {steerToggle}
          </div>
          {steerField}
        </>
      )}
      {failure && <p className="version-error" role="alert" tabIndex={-1}>{failure}</p>}
    </div>
  )
}
