import { createContext, useContext, useEffect, useId, useMemo, useRef, useState, type KeyboardEvent, type ReactNode } from "react"
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
  type MarkupResponse,
  type MarkupRevision,
} from "../api/ingest/markup"
import { getPreferences, savePreferences } from "../api/ingest/preferences"
import { readBrowserJSON, readBrowserValue, removeBrowserValue, writeBrowserValue } from "../browser-storage"
import { formatDateTime } from "../lib/format"
import { authorTones, initialsOf, nameKey, toneClass, toneOf } from "./comment-authors"
import { useToast } from "./Toast"

// The tracked changes and comments in an uploaded Word source (#200), beside
// the document, in two tabs: Changes, where the lawyer marks each change
// Accept, Reject or Leave tracked, and Comments, where each open comment
// thread takes one reply. The other side's changes start undecided; the
// lawyer's own edits (doc.haus proposals accepted into the file, under the
// lawyer's name) start as Leave tracked, so they go back as tracked changes.
// A change left undecided is kept as a tracked change, as one left tracked is. Save sends
// every decision and every approved reply at once, which saves one new
// version. Drafts of the decisions are kept in this browser per document, so
// closing the viewer loses nothing, and carry over to a new version of the
// file for the changes it still lists. A reply is drafted by the selected model
// on request (never on opening), or written by hand, then edited and approved
// by the lawyer. Only approved replies go with Save, in the same version,
// signed with the recorded name. Every name, change, quote and comment is the
// other side's text and renders as plain text. The cards follow Word: comment
// cards with each author's initials and colour, replies under them and a reply
// box at the foot; changes named as Word's Review pane names them.

type Thread = { root: MarkupComment; replies: MarkupComment[] }
type Item =
  | { type: "change"; revision: MarkupRevision; order: number[] }
  | { type: "thread"; thread: Thread; order: number[] }
type Tab = "changes" | "comments"
type Drafts = Record<string, MarkupDecision>
// What Word calls a change in its Review pane.
type ChangeType = "insert" | "delete" | "move" | "format" | "other"
// A change in the list: decided, not yet decided, or only Word can resolve it.
type ChangeState = MarkupDecision | "undecided" | "word"
// A thread's reply: awaited from the lawyer, drafted, approved, or none awaited.
type ReplyState = "needed" | "drafted" | "approved" | "none"
// Each tab's filters; "" shows everything. An author is `=` and the name.
type ChangeFilters = { author: string; type: string; decision: string }
type CommentFilters = { author: string; status: string; reply: string }
// The lawyer's reply to an open thread, by its root comment's id: the text,
// whether the lawyer approved it, what its draft said to check, and the root
// and number of comments it answers. `changed`: the thread gained or lost a
// comment since it was drafted, so it is checked again before approving.
export type Reply = { text: string; approved: boolean; open_items: string[]; root: { author: string; text: string }; seen: number; changed?: true }
type Replies = Record<number, Reply>
// A refused Save: the message, the service's code and the listed changes or
// the thread it names, whose cards carry the message.
type Refusal = { message: string; code?: string; ids: string[]; thread?: number }
// Where focus goes once the panel has rendered: a change's decision, a
// thread's reply (`reply`: its next control, not the refusal), the Comments
// tab's Draft replies row, the Changes tab's Decide all button, or Save.
// `filters`: the first dropdown of a tab's opened filters; `toggle`: its Filters button.
type Focus = { card: string; reply?: boolean } | { thread: string } | { filters: Tab } | { toggle: Tab } | "apply" | "bulk" | "decide-all"

const NO_CHANGE_FILTERS: ChangeFilters = { author: "", type: "", decision: "" }
const NO_COMMENT_FILTERS: CommentFilters = { author: "", status: "", reply: "" }
const DECISIONS: { value: MarkupDecision; label: string }[] = [
  { value: "accept", label: "Accept" },
  { value: "reject", label: "Reject" },
  { value: "hold", label: "Leave tracked" },
]
const TYPES: { value: ChangeType; label: string }[] = [
  { value: "insert", label: "Inserted" },
  { value: "delete", label: "Deleted" },
  { value: "move", label: "Moved" },
  { value: "format", label: "Formatted" },
  { value: "other", label: "Changed" },
]
const STATES: { value: ChangeState; label: string }[] = [
  { value: "undecided", label: "Not yet decided" },
  { value: "accept", label: "Accept" },
  { value: "reject", label: "Reject" },
  { value: "hold", label: "Leave tracked" },
  { value: "word", label: "Resolve in Word" },
]
const STATUSES = [
  { value: "open", label: "Open" },
  { value: "resolved", label: "Resolved" },
]
const REPLY_STATES: { value: ReplyState; label: string }[] = [
  { value: "needed", label: "Needs a reply" },
  { value: "drafted", label: "Draft to approve" },
  { value: "approved", label: "Approved" },
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
const CONTAINER_DETAIL: Record<NonNullable<MarkupLocation["container"]>, string> = { table: "Table", row: "Table row", cell: "Table cell", section: "Page setup" }
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
// What a structural change's badge adds to Word's word for it.
const STRUCTURE_DETAIL: Record<string, string> = {
  row_insert: "Table row",
  row_delete: "Table row",
  cell_insert: "Table cell",
  cell_delete: "Table cell",
  cell_merge: "Table cells merged",
  content_control_insert: "Form field",
  content_control_delete: "Form field",
  numbering_properties_insert: "Numbering added",
  numbering_change: "Numbering",
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
// Author colours (design-system.css --author-1 … --author-6), as Word gives
// each reviewer one.

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
// parent is gone) that is not resolved and still has its place in the text.
const isOpen = (thread: Thread) => thread.root.parent === null && thread.root.resolved !== true && !thread.root.detached
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

const threadAuthors = (thread: Thread) => [thread.root.author, ...thread.replies.map((reply) => reply.author)]
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

// Word's word for a change (Inserted, Deleted, Moved, Formatted), as its
// Review pane shows it; numbering is formatting there.
export function changeType(revision: MarkupRevision): ChangeType {
  if (revision.kind !== "structure") return revision.kind
  if (revision.family.startsWith("numbering")) return "format"
  if (revision.family.endsWith("_insert")) return "insert"
  if (revision.family.endsWith("_delete")) return "delete"
  return "other"
}

// What a change's badge adds to that word: a whole paragraph, a paragraph
// break, a table row, page setup and the like.
function changeDetail(revision: MarkupRevision) {
  if (revision.family === "paragraph_mark") return "Paragraph break"
  if (wholeParagraph(revision)) return "Whole paragraph"
  const container = revision.kind === "format" ? revision.locations.find((place) => place.container)?.container : undefined
  if (container) return CONTAINER_DETAIL[container]
  return revision.kind === "structure" ? STRUCTURE_DETAIL[revision.family] : undefined
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

// An author's colour anywhere in the panel; the lawyer's own is the brand's.
const ToneContext = createContext<(name: string) => number | "self">(() => 0)
// The number of the version on screen, which a run note from another version names.
const VersionContext = createContext<number | undefined>(undefined)

// Drafts are kept per matter and document with the bytes they were made on
// (`base`). A tracked change keeps its id in a newer version while it is
// unchanged, so a change still listed keeps its draft; the rest start undecided.
const draftKey = (matterId: string, name: string) =>
  `dh.markup-decisions.v2/${encodeURIComponent(matterId)}/${encodeURIComponent(name)}`

export type StoredDrafts = { base: string; decisions: Record<string, unknown> }

function readStored(key: string): StoredDrafts | undefined {
  const stored = readBrowserJSON(key) as Partial<StoredDrafts> | null
  if (!stored || typeof stored !== "object" || typeof stored.base !== "string" || !stored.decisions || typeof stored.decisions !== "object") return
  return { base: stored.base, decisions: stored.decisions }
}

// The drafts a reading of the file takes from earlier ones: those for changes
// it still lists, Leave tracked included. When the earlier drafts were made on
// other bytes, `carried` names the ones kept and `dropped` counts decisions on
// changes now gone.
export function carryDrafts(earlier: StoredDrafts | undefined, markup: Pick<DocumentMarkup, "version" | "revisions">): {
  drafts: Drafts
  carried: string[]
  dropped: number
} {
  if (!earlier) return { drafts: {}, carried: [], dropped: 0 }
  const listed = new Set(markup.revisions.filter(decidable).map((revision) => revision.id))
  const decided = Object.entries(earlier.decisions).filter((entry): entry is [string, MarkupDecision] => DECISIONS.some((option) => option.value === entry[1]))
  const drafts: Drafts = Object.fromEntries(decided.filter(([id]) => listed.has(id)))
  if (earlier.base === markup.version.sha256) return { drafts, carried: [], dropped: 0 }
  return { drafts, carried: Object.keys(drafts), dropped: decided.length - Object.keys(drafts).length }
}

// Replies are kept the same way, under their own key. A reply stays with its
// thread while the root reads the same and no comment came after the ones it
// answers; carried to other bytes, an approved reply needs approving again.
const replyKey = (matterId: string, name: string) =>
  `dh.markup-replies.v1/${encodeURIComponent(matterId)}/${encodeURIComponent(name)}`

// The tab last open on a document, for this browser session.
const tabKey = (matterId: string, name: string) =>
  `dh.markup-tab.v1/${encodeURIComponent(matterId)}/${encodeURIComponent(name)}`

// The tab a document opens on: the one last open while it still has items,
// else Changes unless the file has only comments.
function openingTab(key: string, markup: Pick<DocumentMarkup, "revisions" | "comments">): Tab {
  const stored = readBrowserValue(key, true)
  if (stored === "changes" && markup.revisions.length) return "changes"
  if (stored === "comments" && markup.comments.length) return "comments"
  return markup.revisions.length || !markup.comments.length ? "changes" : "comments"
}

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
    typeof reply.root?.author === "string" && typeof reply.root.text === "string" && (reply.changed === undefined || reply.changed === true)
}

// The replies a reading of the file takes from earlier ones: those whose
// thread is still open and starts from the same comment. A thread that gained
// or lost a comment keeps its reply as a draft marked `changed`, so the
// lawyer's words are never lost. From other bytes, `carried` names the ones
// kept and `dropped` counts the rest.
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
    if (!thread || !isReply(reply) || reply.root.author !== thread.root.author || reply.root.text !== thread.root.text) continue
    const seen = 1 + thread.replies.length
    replies[thread.root.id] = reply.seen !== seen ? { ...reply, approved: false, seen, changed: true } : same ? reply : { ...reply, approved: false }
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
    refused ? `${name} changed after this list was read, so nothing was saved.` : `${name} changed since these ${decided ? "decisions were made" : "replies were drafted"}.`,
    decisions.carried && `${plural(decisions.carried, "decision")} carried over — check ${decisions.carried === 1 ? "it" : "them"} before saving.`,
    decisions.dropped && `${plural(decisions.dropped, "decision")} could not be carried over and ${decisions.dropped === 1 ? "needs" : "need"} deciding again.`,
    replies.carried && `${plural(replies.carried, "draft reply", "draft replies")} carried over — read ${replies.carried === 1 ? "it" : "them"} again before approving.`,
    replies.dropped && `${plural(replies.dropped, "draft reply", "draft replies")} could not be carried over, as ${replies.dropped === 1 ? "its comment is" : "their comments are"} no longer open.`,
    refused && decisions.dropped && !decisions.carried && "Decide the changes again.",
  ].filter(Boolean).join(" ")
}

// A filter's options with how many items each would show, the rest of the
// tab's filters applied; an option that would show none is disabled unless
// chosen. The chosen one stays listed even when the file no longer has it.
type Option = { value: string; label: string; count: number }
function facets<T>(list: T[], tests: Record<string, (item: T) => boolean>, key: string, choices: { value: string; label: string; test: (item: T) => boolean }[]) {
  const rest = list.filter((item) => Object.entries(tests).every(([name, test]) => name === key || test(item)))
  return choices.map(({ value, label, test }): Option => ({ value, label, count: rest.filter(test).length }))
}

// The filters in use, each as "Reviewer: Avery Lindqvist" for its chip.
function activeFilters<K extends string>(filters: Record<K, string>, fields: [K, string, Option[]][]) {
  return fields.filter(([key]) => filters[key]).map(([key, label, options]) =>
    ({ key, label: `${label}: ${options.find((option) => option.value === filters[key])?.label ?? filters[key].replace(/^=/, "")}` }))
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
  onClose,
  reveal,
  onRevealed,
}: {
  matterId: string
  name: string
  markup: DocumentMarkup
  // A newer reading is on its way; until then nothing can be saved.
  loading?: boolean
  // doc.haus's own proposals on this document, which are resolved first.
  proposalsPending?: number
  // A new version was saved: the viewer reloads the document and its versions,
  // and shows what the result says to check.
  onApplied: (result: MarkupDecisionResult) => void
  // The service refused the batch: read the file and its markup again.
  onRefetch: () => void
  // doc.haus's own proposals block saving: show them instead of this panel.
  onShowProposals?: () => void
  // Close the panel, as the viewer's Tracked changes toggle does.
  onClose?: () => void
  // A comment chosen in the document (its Word id, a reply's or its root's;
  // `seq` repeats a request): the panel opens its thread on the Comments tab,
  // moves focus to it, and calls onRevealed so the viewer can clear the
  // request before the panel is next opened.
  reveal?: { thread: number; seq: number }
  onRevealed?: () => void
}) {
  const sha256 = markup.version.sha256
  const key = draftKey(matterId, name)
  const repliesKey = replyKey(matterId, name)
  const tabsKey = tabKey(matterId, name)
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
  const [undo, setUndo] = useState<{ drafts: Drafts; carried: string[]; text: string }>()
  const [tab, setTabState] = useState<Tab>(() => openingTab(tabsKey, markup))
  const [changeFilters, setChangeFilters] = useState<ChangeFilters>(NO_CHANGE_FILTERS)
  const [commentFilters, setCommentFilters] = useState<CommentFilters>(NO_COMMENT_FILTERS)
  // Cards acted on since the filters last changed stay in view, so a card does
  // not vanish from "Not yet decided" the moment it is decided.
  const [stay, setStay] = useState<string[]>([])
  const [bulkOpen, setBulkOpen] = useState(false)
  // Per tab, while the panel is open: whether its filters and its guide are shown.
  const [filtersOpen, setFiltersOpen] = useState<Record<Tab, boolean>>({ changes: false, comments: false })
  const [help, setHelp] = useState<Record<Tab, boolean>>({ changes: false, comments: false })
  // undefined while loading; "" when no name is saved, so the bar asks once.
  const [attorney, setAttorney] = useState<string>()
  // The settings could not be read: who "you" are is unknown.
  const [settingsFailed, setSettingsFailed] = useState(false)
  const [firm, setFirm] = useState("")
  const [nameDraft, setNameDraft] = useState("")
  // What the panel's live region says: drafting, drafts ready, a bulk
  // decision, a reply discarded, a new reading of the file.
  const [said, setSaid] = useState("")
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
  const decideAll = useRef<HTMLButtonElement>(null)
  const cards = useRef(new Map<string, HTMLLIElement>())
  const tabButtons = useRef(new Map<Tab, HTMLButtonElement>())
  const filterToggles = useRef(new Map<Tab, HTMLButtonElement>())
  const filterGroups = useRef(new Map<Tab, HTMLDivElement>())
  const ids = useId()
  const toast = useToast()

  // A new reading of other bytes keeps the drafts for changes it still lists,
  // and the replies for threads it still holds unchanged, and says so; after
  // this panel's own Save the decisions start afresh and the replies it did
  // not send stay, without a note. Another document takes its own drafts. A
  // tab left with nothing in it gives way to the other.
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
    const note = carriedNote(name, { carried: next.carried.length, dropped: next.dropped },
      own ? { carried: 0, dropped: 0 } : { carried: kept.carried.length, dropped: kept.dropped }, !fresh && refusal?.code === "base_changed")
    setNotice(note)
    setSaid(note ?? "")
    if (fresh) {
      setDrafting([])
      setSteers({})
      setChangeFilters(NO_CHANGE_FILTERS)
      setCommentFilters(NO_COMMENT_FILTERS)
      setTabState(openingTab(tabsKey, markup))
    } else if (tab === "changes" ? !markup.revisions.length && markup.comments.length : !markup.comments.length && markup.revisions.length) {
      setTabState(tab === "changes" ? "comments" : "changes")
    }
    setStay([])
    setBulkOpen(false)
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
      (prefs) => {
        if (cancelled) return
        setAttorney(prefs.attorney?.trim() ?? "")
        setFirm(prefs.firm?.trim() ?? "")
      },
      () => {
        if (cancelled) return
        setAttorney("")
        setSettingsFailed(true)
      },
    )
    return () => {
      cancelled = true
    }
  }, [])

  // Focus waits for the render that shows its target: a refused change's
  // chosen decision or a refused reply's message, a reply's next control after
  // an action on it, the Draft replies row, Decide all, or Save once the file
  // has been read again (the proposals button or the message while Save cannot
  // be used).
  useEffect(() => {
    if (!focus || busy) return
    const card = typeof focus === "object" && "card" in focus ? cards.current.get(focus.card) : undefined
    if (typeof focus === "object" && "filters" in focus) {
      filterGroups.current.get(focus.filters)?.querySelector<HTMLElement>("select")?.focus()
    } else if (typeof focus === "object" && "toggle" in focus) {
      filterToggles.current.get(focus.toggle)?.focus()
    } else if (focus === "bulk") {
      ;(bulkRow.current?.querySelector<HTMLElement>("button:not(:disabled)") ?? tabButtons.current.get("comments"))?.focus()
    } else if (focus === "decide-all") {
      ;(decideAll.current ?? tabButtons.current.get("changes"))?.focus()
    } else if (typeof focus === "object" && "thread" in focus) {
      const thread = cards.current.get(focus.thread)
      thread?.scrollIntoView?.({ block: "nearest" })
      thread?.focus({ preventScroll: true })
    } else if (card && typeof focus === "object" && "reply" in focus && focus.reply) {
      card.querySelector<HTMLElement>("[data-reply-focus]:not(:disabled)")?.focus()
    } else if (card) {
      card.scrollIntoView?.({ block: "center" })
      const control = card.querySelector<HTMLElement>('input[type="radio"]:checked') ?? card.querySelector<HTMLElement>(".version-error")
      control?.focus({ preventScroll: true })
    } else {
      if (loading) return
      const target = [applyButton.current, proposalsButton.current, refusalText.current].find((element) => element && !(element as HTMLButtonElement).disabled) ?? heading.current
      target?.focus()
    }
    setFocus(undefined)
  })

  const items = useMemo(() => documentOrder(markup), [markup])
  const changeList = items.flatMap((item) => (item.type === "change" ? [item.revision] : []))
  const threadList = items.flatMap((item) => (item.type === "thread" ? [item.thread] : []))
  const savedName = attorney?.trim()
  const actor = savedName || nameDraft.trim()
  // Every name doc.haus signs the lawyer's accepted edits with, as ingest reads
  // them: the saved attorney, the firm and doc.haus, so what it signed before a
  // name was saved stays the lawyer's. Until the settings are read, or when
  // they cannot be, no change is taken for the lawyer's own.
  const ours = attorney === undefined || settingsFailed ? undefined
    : new Set([savedName, firm, "doc.haus"].filter((entry): entry is string => Boolean(entry)).map(nameKey))
  const isYou = (author: string) => Boolean(ours?.has(nameKey(author)))
  const yours = (revision: MarkupRevision) => isYou(revision.author)
  // The lawyer's own edit stays tracked until the lawyer decides otherwise.
  const decisionOf = (revision: MarkupRevision): MarkupDecision | undefined => drafts[revision.id] ?? (yours(revision) ? "hold" : undefined)
  const listed = markup.revisions.filter(decidable)
  const accepted = listed.filter((revision) => decisionOf(revision) === "accept").length
  const rejected = listed.filter((revision) => decisionOf(revision) === "reject").length
  const held = listed.filter((revision) => decisionOf(revision) === "hold").length
  const undecided = listed.length - accepted - rejected - held
  const left = held + undecided
  const elsewhere = markup.revisions.length - listed.length
  // Open threads in document order. One awaits a reply while its latest
  // comment is not the saved attorney's; only those are offered a draft.
  const open = threadList.filter(isOpen)
  const awaiting = (thread: Thread) => !savedName || (thread.replies.at(-1) ?? thread.root).author.trim().toLowerCase() !== savedName.toLowerCase()
  const approved = open.filter((thread) => replies[thread.root.id]?.approved)
  const unapproved = open.filter((thread) => replies[thread.root.id] && !replies[thread.root.id]!.approved).length
  const decided = accepted > 0 || rejected > 0
  const blocked = proposalsPending > 0
    ? `Accept or reject doc.haus's ${plural(proposalsPending, "proposed change")} on this document first.`
    : !decided && !approved.length ? (listed.length ? "Accept or reject a change, or approve a reply, to save." : "Approve a reply to save a new version.") : undefined
  const canApply = !blocked && !busy && !loading && attorney !== undefined
  const proposalsFirst = proposalsPending > 0 || refusal?.code === "proposals_pending"

  // Each author's colour, told apart within this file; the lawyer's is the brand's.
  const tones = useMemo(() => authorTones([...markup.revisions.map((revision) => revision.author), ...markup.comments.map((comment) => comment.author)]), [markup])
  const tone = (author: string) => toneOf(tones, author, (name) => Boolean(savedName && nameKey(name) === nameKey(savedName)) || isYou(name))

  // What each card shows as its state, and the filters that read it.
  const stateOf = (revision: MarkupRevision): ChangeState => (!decidable(revision) ? "word" : decisionOf(revision) ?? "undecided")
  const replyStateOf = (thread: Thread): ReplyState => {
    if (!isOpen(thread)) return "none"
    const reply = replies[thread.root.id]
    if (reply) return reply.approved ? "approved" : "drafted"
    return awaiting(thread) ? "needed" : "none"
  }
  const changeTests: Record<keyof ChangeFilters, (revision: MarkupRevision) => boolean> = {
    author: (revision) => !changeFilters.author || changeFilters.author === `=${revision.author}`,
    type: (revision) => !changeFilters.type || changeType(revision) === changeFilters.type,
    decision: (revision) => !changeFilters.decision || stateOf(revision) === changeFilters.decision,
  }
  const commentTests: Record<keyof CommentFilters, (thread: Thread) => boolean> = {
    author: (thread) => !commentFilters.author || threadAuthors(thread).some((author) => commentFilters.author === `=${author}`),
    status: (thread) => !commentFilters.status || (commentFilters.status === "open" ? isOpen(thread) : thread.root.resolved === true),
    reply: (thread) => !commentFilters.reply || replyStateOf(thread) === commentFilters.reply,
  }
  const shownChangeList = changeList.filter((revision) => stay.includes(revision.id) || Object.values(changeTests).every((test) => test(revision)))
  const shownThreadList = threadList.filter((thread) => stay.includes(threadCard(thread)) || Object.values(commentTests).every((test) => test(thread)))
  // Decide all sets the other side's shown changes only; the lawyer's own
  // edits are decided one by one.
  const shownDecidable = shownChangeList.filter((revision) => decidable(revision) && !yours(revision))
  const ownShown = shownChangeList.some(yours)
  const changesFiltered = Object.values(changeFilters).some(Boolean)
  const undrafted = shownThreadList.filter((thread) => isOpen(thread) && awaiting(thread) && !replies[thread.root.id] && !drafting.includes(thread.root.id))

  // The reviewers each tab lists, in the file's order of authors.
  const authorChoices = (names: string[]) => {
    const present = new Set(names)
    const ordered = [...markup.authors.map((author) => author.name).filter((author) => present.has(author)), ...[...present].filter((author) => !markup.authors.some((entry) => entry.name === author))]
    return ordered.map((author) => ({ value: `=${author}`, label: isYou(author) ? `You (${authorName(author)})` : authorName(author) }))
  }
  const changeOptions = {
    author: facets(changeList, changeTests, "author", authorChoices(changeList.map((revision) => revision.author)).map((choice) => ({ ...choice, test: (revision: MarkupRevision) => `=${revision.author}` === choice.value }))),
    type: facets(changeList, changeTests, "type", TYPES.filter((type) => changeList.some((revision) => changeType(revision) === type.value)).map((type) => ({ ...type, test: (revision: MarkupRevision) => changeType(revision) === type.value }))),
    decision: facets(changeList, changeTests, "decision", STATES.filter((state) => state.value !== "word" || changeList.some((revision) => !decidable(revision))).map((state) => ({ ...state, test: (revision: MarkupRevision) => stateOf(revision) === state.value }))),
  }
  const commentOptions = {
    author: facets(threadList, commentTests, "author", authorChoices(threadList.flatMap(threadAuthors)).map((choice) => ({ ...choice, test: (thread: Thread) => threadAuthors(thread).some((author) => `=${author}` === choice.value) }))),
    status: facets(threadList, commentTests, "status", STATUSES.map((status) => ({ ...status, test: (thread: Thread) => (status.value === "open" ? isOpen(thread) : thread.root.resolved === true) }))),
    reply: facets(threadList, commentTests, "reply", REPLY_STATES.map((state) => ({ ...state, test: (thread: Thread) => replyStateOf(thread) === state.value }))),
  }

  const changeChips = activeFilters(changeFilters, [["author", "Reviewer", changeOptions.author], ["type", "Type", changeOptions.type], ["decision", "Decision", changeOptions.decision]])
  const commentChips = activeFilters(commentFilters, [["author", "Reviewer", commentOptions.author], ["status", "Status", commentOptions.status], ["reply", "Reply", commentOptions.reply]])

  // A refusal is shown on the cards it names; otherwise, or while they are
  // filtered out or on the other tab, beside Save. The proposals line already
  // says what a proposals refusal says once the proposals are known.
  const shownIds = new Set(tab === "changes" ? shownChangeList.map((revision) => revision.id) : [])
  const shownThreads = new Set(tab === "comments" ? shownThreadList.map((thread) => thread.root.id) : [])
  const onCards = !!refusal && (refusal.ids.some((id) => shownIds.has(id)) || (refusal.thread !== undefined && shownThreads.has(refusal.thread)))
  const refusalShown = refusal && !onCards && !(refusal.code === "proposals_pending" && proposalsPending > 0) ? refusal.message : undefined

  const keep = (...cardKeys: string[]) => setStay((current) => [...new Set([...current, ...cardKeys])])

  function setTab(next: Tab) {
    setTabState(next)
    setBulkOpen(false)
    writeBrowserValue(tabsKey, next, true)
  }

  // Left and Right (and Home and End) move between the tabs and open the one
  // they reach, as tabs do everywhere.
  function stepTab(event: KeyboardEvent<HTMLDivElement>) {
    const order: Tab[] = ["changes", "comments"]
    const at = order.indexOf(tab)
    const to = { ArrowRight: at + 1, ArrowLeft: at - 1, Home: 0, End: order.length - 1 }[event.key]
    if (to === undefined) return
    event.preventDefault()
    const next = order[(to + order.length) % order.length]!
    setTab(next)
    tabButtons.current.get(next)?.focus()
  }

  function filterChanges(next: Partial<ChangeFilters>) {
    setChangeFilters((current) => ({ ...current, ...next }))
    setStay([])
    setBulkOpen(false)
  }

  function filterComments(next: Partial<CommentFilters>) {
    setCommentFilters((current) => ({ ...current, ...next }))
    setStay([])
  }

  // Filters opens a tab's dropdowns and puts focus in the first; closing them,
  // with the button again or Escape, keeps what they filter.
  function toggleFilters(which: Tab) {
    const opening = !filtersOpen[which]
    setFiltersOpen((current) => ({ ...current, [which]: opening }))
    setBulkOpen(false)
    setFocus(opening ? { filters: which } : { toggle: which })
  }

  // A chip or Clear filters goes, and focus goes back to the Filters button.
  function unfilter(which: Tab, key?: string) {
    if (which === "changes") filterChanges(key ? { [key]: "" } : NO_CHANGE_FILTERS)
    else filterComments(key ? { [key]: "" } : NO_COMMENT_FILTERS)
    setFocus({ toggle: which })
  }

  // Still to do, one click away: the changes not yet decided, or the drafts
  // to approve, with every other filter cleared.
  function showUndecided() {
    setTab("changes")
    filterChanges({ ...NO_CHANGE_FILTERS, decision: "undecided" })
    setSaid(`Showing the ${plural(undecided, "change")} not yet decided.`)
  }

  // A comment chosen in the document opens its thread, even past the filters.
  useEffect(() => {
    if (!reveal) return
    onRevealed?.()
    const thread = threadList.find((entry) => [entry.root, ...entry.replies].some((comment) => comment.id === reveal.thread))
    setTab("comments")
    if (!thread) return
    keep(threadCard(thread))
    setFocus({ thread: threadCard(thread) })
    // Only a new request acts; the lists it reads are the current render's.
  }, [reveal?.thread, reveal?.seq])

  function showDrafts() {
    setTab("comments")
    filterComments({ ...NO_COMMENT_FILTERS, reply: "drafted" })
    setSaid(`Showing the ${plural(unapproved, "draft")} to approve.`)
  }

  function decide(changes: MarkupRevision[], decision: MarkupDecision, bulk?: string) {
    setConfirming(false)
    const next = { ...drafts }
    for (const revision of changes) next[revision.id] = decision
    const decided = new Set(changes.map((revision) => revision.id))
    const unchanged = changes.every((revision) => drafts[revision.id] === decision)
    setDrafts(next)
    setCarried(carried.filter((id) => !decided.has(id)))
    setUndo(bulk && !unchanged ? { drafts, carried, text: bulk } : undefined)
    if (bulk) setSaid(bulk)
    keep(...decided)
    // A change the refusal named is settled once the lawyer decides it again.
    const left = refusal?.ids.filter((id) => !decided.has(id)) ?? []
    setRefusal(left.length ? { ...refusal!, ids: left } : undefined)
  }

  // Decide all: every decidable change the filters show, then focus back on
  // the button that opened the choices.
  function decideShown(decision: MarkupDecision) {
    const verb = { accept: "accepted", reject: "rejected", hold: "left tracked" }[decision]
    decide(shownDecidable, decision, `${plural(shownDecidable.length, "change")} ${verb}.`)
    setBulkOpen(false)
    setFocus("decide-all")
  }

  function undoBulk() {
    if (!undo) return
    setDrafts(undo.drafts)
    setCarried(undo.carried)
    setUndo(undefined)
    setConfirming(false)
    setFocus("decide-all")
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
    keep(threadCard(thread))
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
    // A new draft replaces the lawyer's text, which Undo brings back.
    const earlier = new Map(queue.map((thread) => [thread.root.id, replies[thread.root.id] ?? discarded[thread.root.id]?.reply]))
    for (const id of ids) controllers.current.set(id, new AbortController())
    if (ids.length) setSaid(`Drafting ${plural(ids.length, "reply", "replies")}…`)
    setDrafting((current) => [...current, ...ids])
    setDraftErrors((current) => without(current, ...ids))
    setDiscarded((current) => without(current, ...ids))
    setDrafted(undefined)
    keep(...queue.map(threadCard))
    const decisions = listed.flatMap((revision) => (drafts[revision.id] && drafts[revision.id] !== "hold" ? [{ id: revision.id, decision: drafts[revision.id]! }] : []))
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
        const replaced = earlier.get(id)
        if (replaced?.text.trim()) setDiscarded((current) => ({ ...current, [id]: { reply: replaced, text: "Draft replaced." } }))
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
    if (done.size) {
      const ready = `${plural(done.size, "draft reply", "draft replies")} ready — read ${done.size === 1 ? "it" : "each one"}, then approve it.`
      setDrafted(ready)
      setSaid(ready)
    } else if (queue.length) setSaid(queue.length === 1 ? "The reply was not drafted." : "No replies were drafted.")
    return queue.filter((thread) => done.has(thread))
  }

  function cancelDrafts(ids = [...controllers.current.keys()]) {
    for (const id of ids) {
      controllers.current.get(id)?.abort()
      controllers.current.delete(id)
    }
    setDrafting((current) => current.filter((entry) => !ids.includes(entry)))
  }

  // Every shown thread awaiting a reply that has none: focus stays on the row
  // (its Cancel), then goes to the first draft if the lawyer is still there,
  // or back to the row when none came.
  async function draftAll() {
    setConfirming(false)
    setFocus("bulk")
    const done = await draftReplies(undrafted)
    if (focusFree(bulkRow.current)) setFocus(done.length ? { card: threadCard(done[0]!), reply: true } : "bulk")
  }

  function cancelAll() {
    cancelDrafts()
    setSaid("Drafting cancelled.")
    setFocus("bulk")
  }

  // One thread's actions; each leaves focus on the reply's next control.
  function replyActions(thread: Thread, reply: Reply | undefined): ReplyActions {
    const card = threadCard(thread)
    const next = () => setFocus({ card, reply: true })
    return {
      draft: async () => {
        setConfirming(false)
        next()
        await draftReplies([thread])
        if (focusFree(cards.current.get(card))) next()
      },
      cancel: () => {
        cancelDrafts([thread.root.id])
        setSaid("Drafting cancelled.")
        next()
      },
      write: () => {
        setReply(thread, replyTo(thread, ""))
        next()
      },
      change: (text: string) => reply && setReply(thread, { ...reply, text }),
      approve: () => {
        if (!reply) return
        // Approving is the check a changed thread asked for.
        const { changed: _changed, ...checked } = reply
        setReply(thread, { ...checked, approved: true })
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
        const text = reply.approved ? "Reply removed." : "Draft discarded."
        setDiscarded((current) => ({ ...current, [thread.root.id]: { reply, text } }))
        setSaid(text)
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
    // The Save button is rendered again once the confirmation goes.
    requestAnimationFrame(() => applyButton.current?.focus())
  }

  // What a refused Save shows, and where focus goes: the first change or the
  // thread it names, on its tab, with that tab's filters cleared if they hide
  // it, or Save.
  function refused(cause: unknown): Focus {
    if (!(cause instanceof MarkupRefusedError)) {
      setRefusal({ message: humanizeError(cause, "The decisions could not be saved. Try again."), ids: [] })
      return "apply"
    }
    // The file or its changes moved on: read them again before another try.
    if (cause.status === 409) onRefetch()
    const named = cause.ids.filter((id) => markup.revisions.some((revision) => revision.id === id))
    const thread = open.find((entry) => entry.root.id === cause.thread)
    // The panel reads the file again itself; there is nothing to refresh.
    const message = cause.code === "base_changed" ? `${name} changed after this list was read, so nothing was saved.` : cause.message
    setRefusal({ message, code: cause.code, ids: named, thread: thread?.root.id })
    if (!named.length && !thread) return "apply"
    setTab(named.length ? "changes" : "comments")
    const shown = new Set([...shownChangeList.map((revision) => revision.id), ...shownThreadList.map(threadCard)])
    if (named.length ? !named.every((id) => shown.has(id)) : !shown.has(threadCard(thread!))) {
      setChangeFilters(NO_CHANGE_FILTERS)
      setCommentFilters(NO_COMMENT_FILTERS)
      setStay([])
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
      // Only replies the lawyer approved are sent; drafts stay drafts. A change
      // left undecided is left tracked, as Leave tracked leaves it.
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
          `${counts.held ?? left + elsewhere} left tracked`,
        ].join(", "),
        added > 0 && `${plural(added, "reply", "replies")} added`,
      ].filter(Boolean).join("; ")
      // The tally names the changes left tracked; the version's own notice
      // about tracked changes left in it would only repeat them.
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
  // Every change left tracked, as the toast and the version's label count
  // them: those left tracked, those not decided and those to resolve in Word.
  const kept = left + elsewhere
  const asTracked = kept === 1 ? "as a tracked change" : "as tracked changes"
  const including = [undecided > 0 && `${undecided} you have not decided`, elsewhere > 0 && `${elsewhere} to resolve in Word`].filter((part): part is string => Boolean(part))
  const stays = !kept ? ""
    : undecided === kept ? ` ${kept === 1 ? "1 change you have not decided stays" : `${kept} changes you have not decided stay`} ${asTracked}.`
    : elsewhere === kept ? ` ${kept === 1 ? "1 change to resolve in Word stays" : `${kept} changes to resolve in Word stay`} ${asTracked}.`
    : ` ${plural(kept, "change stays", "changes stay")} ${asTracked}${including.length ? `, including ${joined(including)}` : ""}.`
  const confirmText = `This saves a new version of ${name} with ${joined(saved)}.${stays}` +
    `${unapproved ? ` ${plural(unapproved, "draft reply is", "draft replies are")} not approved, so ${unapproved === 1 ? "it is" : "they are"} not added.` : ""}` +
    " The current version stays in the history."
  // The other side's reviewers; the lawyer, when the file holds the lawyer's
  // own edits or comments, is named apart.
  const reviewers = markup.authors.filter((author) => !isYou(author.name) && !(savedName && nameKey(author.name) === nameKey(savedName))).length
  const withYou = markup.authors.some((author) => isYou(author.name) || (savedName && nameKey(author.name) === nameKey(savedName)))
  const theirListed = listed.filter((revision) => !yours(revision)).length
  const decideLabel = ownShown ? `Decide all ${shownDecidable.length} of their changes shown` : `Decide all ${shownDecidable.length} shown`
  const tabs: { value: Tab; label: string; count: number }[] = [
    { value: "changes", label: "Changes", count: changeList.length },
    { value: "comments", label: "Comments", count: threadList.length },
  ]

  return (
    <ToneContext.Provider value={tone}>
    <VersionContext.Provider value={markup.version.number}>
    <aside id="document-markup" className="changes-panel markup-panel" aria-labelledby={`${ids}-title`}>
      <p className="sr-only" role="status">{said}</p>
      <div className="changes-head markup-head">
        <div className="markup-title">
          <div className="markup-title-text">
            <h3 id={`${ids}-title`} ref={heading} tabIndex={-1}>Tracked changes</h3>
            <p className="markup-summary">
              {plural(items.length, "item")}{reviewers > 0 && ` from ${plural(reviewers, "reviewer")}`}{withYou && (reviewers > 0 ? " and you" : " from you")}
            </p>
          </div>
          {onClose && (
            <button type="button" className="markup-close" aria-label="Close tracked changes" title="Close" onClick={onClose}>
              <Icon name="close" />
            </button>
          )}
        </div>
        <div className="markup-tabs" role="tablist" aria-label="Tracked changes and comments" onKeyDown={stepTab}>
          {tabs.map((entry) => (
            <button type="button" role="tab" key={entry.value} id={`${ids}-${entry.value}-tab`} aria-selected={tab === entry.value} aria-controls={`${ids}-${entry.value}-panel`}
              tabIndex={tab === entry.value ? 0 : -1} className="markup-tab" onClick={() => setTab(entry.value)}
              ref={(node) => { if (node) tabButtons.current.set(entry.value, node); else tabButtons.current.delete(entry.value) }}>
              {entry.label} <span className="markup-tab-count">{entry.count}</span>
            </button>
          ))}
        </div>
      </div>

      <div className="markup-tabpanel" role="tabpanel" id={`${ids}-changes-panel`} aria-labelledby={`${ids}-changes-tab`} hidden={tab !== "changes"}>
        {changeList.length === 0 ? (
          <p className="muted markup-empty">There are no tracked changes in this file.</p>
        ) : (
          <>
            <Intro id={`${ids}-changes-help`} open={help.changes} lead="Accept or reject each change, as in Word."
              onToggle={() => setHelp((current) => ({ ...current, changes: !current.changes }))}>
              <p id={`${ids}-hold-hint`}><strong>Leave tracked</strong> keeps a change as a tracked change, with its author and date, in the new version.</p>
              <p>doc.haus edits you accept appear here as your own, under your name.</p>
            </Intro>
            <FilterBar id={`${ids}-change-filters`} label="Filter changes" open={filtersOpen.changes} chips={changeChips}
              shown={shownChangeList.length} total={changeList.length} noun="change"
              toggleRef={(node) => { if (node) filterToggles.current.set("changes", node); else filterToggles.current.delete("changes") }}
              groupRef={(node) => { if (node) filterGroups.current.set("changes", node); else filterGroups.current.delete("changes") }}
              onToggle={() => toggleFilters("changes")} onRemove={(key) => unfilter("changes", key)} onClear={() => unfilter("changes")}
              action={shownDecidable.length > 0 && (!ours ? settingsFailed && (
                <p className="markup-bulk-unavailable">Decide all is unavailable: your settings could not be read.</p>
              ) : (
                <button type="button" ref={decideAll} className="markup-bulk-toggle" aria-expanded={bulkOpen} aria-controls={bulkOpen ? `${ids}-bulk` : undefined} disabled={busy}
                  onClick={() => setBulkOpen(!bulkOpen)} onKeyDown={(event) => {
                    // Escape closes the open choices, not the viewer.
                    if (event.key !== "Escape" || !bulkOpen) return
                    event.preventDefault()
                    event.stopPropagation()
                    setBulkOpen(false)
                  }}>
                  {decideLabel}
                  <Icon name="chevron" />
                </button>
              ))}>
              <FilterSelect id={`${ids}-change-author`} label="Reviewer" all="All reviewers" value={changeFilters.author} options={changeOptions.author} onChange={(author) => filterChanges({ author })} />
              <FilterSelect id={`${ids}-change-type`} label="Type" all="All change types" value={changeFilters.type} options={changeOptions.type} onChange={(type) => filterChanges({ type })} />
              <FilterSelect id={`${ids}-change-decision`} label="Decision" all="All decisions" value={changeFilters.decision} options={changeOptions.decision} onChange={(decision) => filterChanges({ decision })} />
            </FilterBar>
            {bulkOpen && ours && shownDecidable.length > 0 && (
              <div id={`${ids}-bulk`} className="markup-bulk-menu" role="group" aria-label={decideLabel} onKeyDown={(event) => {
                if (event.key !== "Escape") return
                event.preventDefault()
                event.stopPropagation()
                setBulkOpen(false)
                decideAll.current?.focus()
              }}>
                <p className="markup-bulk-scope">
                  {changesFiltered || shownDecidable.length < theirListed
                    ? `Sets the ${plural(shownDecidable.length, "change")} of theirs your filters show. The others keep their decisions.`
                    : `Sets every change of theirs in this file that doc.haus can apply.`}
                  {ownShown && " Your own edits keep their decisions."}
                </p>
                <div className="markup-bulk-options">
                  {DECISIONS.map((option) => (
                    <button type="button" key={option.value} className={`markup-bulk-option is-${option.value}`} disabled={busy} onClick={() => decideShown(option.value)}>
                      <Icon name={option.value} />
                      {option.value === "hold" ? `Leave all ${shownDecidable.length} tracked` : `${option.label} all ${shownDecidable.length}`}
                    </button>
                  ))}
                </div>
              </div>
            )}
            {undo && (
              <p className="markup-undo">
                {undo.text}{" "}
                <button type="button" className="linklike" disabled={busy} onClick={undoBulk}>Undo</button>
              </p>
            )}
            {shownChangeList.length === 0 ? (
              <p className="muted markup-empty">
                No changes match these filters.{" "}
                <button type="button" className="linklike" onClick={() => unfilter("changes")}>Clear filters</button>
              </p>
            ) : (
              <ol className="markup-list" aria-label={`Changes in ${name}, in document order`} aria-busy={loading || undefined}>
                {shownChangeList.map((revision) => (
                  <ChangeCard key={revision.id} revision={revision} comments={markup.comments} state={stateOf(revision)} own={yours(revision)} disabled={busy} holdHint={`${ids}-hold-hint`}
                    carried={carried.includes(revision.id)} response={markup.responses?.[revision.id]}
                    failure={refusal?.ids.includes(revision.id) ? refusal.message : undefined}
                    cardRef={(node) => { if (node) cards.current.set(revision.id, node); else cards.current.delete(revision.id) }}
                    onDecide={(decision) => decide([revision], decision)}
                    onHold={() => {
                      decide([revision], "hold")
                      setFocus({ card: revision.id })
                    }} />
                ))}
              </ol>
            )}
          </>
        )}
      </div>

      <div className="markup-tabpanel" role="tabpanel" id={`${ids}-comments-panel`} aria-labelledby={`${ids}-comments-tab`} hidden={tab !== "comments"}>
        {threadList.length === 0 ? (
          <p className="muted markup-empty">There are no comments in this file.</p>
        ) : (
          <>
            <Intro id={`${ids}-comments-help`} open={help.comments} lead="Reply to comments as you would in Word."
              onToggle={() => setHelp((current) => ({ ...current, comments: !current.comments }))}>
              <p>doc.haus can draft each reply for you to edit. Only the replies you approve are added when you save.</p>
            </Intro>
            <FilterBar id={`${ids}-comment-filters`} label="Filter comments" open={filtersOpen.comments} chips={commentChips}
              shown={shownThreadList.length} total={threadList.length} noun="comment"
              toggleRef={(node) => { if (node) filterToggles.current.set("comments", node); else filterToggles.current.delete("comments") }}
              groupRef={(node) => { if (node) filterGroups.current.set("comments", node); else filterGroups.current.delete("comments") }}
              onToggle={() => toggleFilters("comments")} onRemove={(key) => unfilter("comments", key)} onClear={() => unfilter("comments")}>
              <FilterSelect id={`${ids}-comment-author`} label="Reviewer" all="All reviewers" value={commentFilters.author} options={commentOptions.author} onChange={(author) => filterComments({ author })} />
              <FilterSelect id={`${ids}-comment-status`} label="Status" all="Open and resolved" value={commentFilters.status} options={commentOptions.status} onChange={(status) => filterComments({ status })} />
              <FilterSelect id={`${ids}-comment-reply`} label="Reply" all="All replies" value={commentFilters.reply} options={commentOptions.reply} onChange={(reply) => filterComments({ reply })} />
            </FilterBar>
            {(undrafted.length > 0 || drafting.length > 0 || drafted) && (
              <div className="markup-bulk markup-draft-all" ref={bulkRow}>
                {drafting.length > 0 ? (
                  <>
                    <span className="markup-reply-status"><span className="markup-spinner" aria-hidden="true" />Drafting {plural(drafting.length, "reply", "replies")}…</span>
                    <button type="button" className="icon-btn" onClick={cancelAll}>Cancel drafting</button>
                  </>
                ) : (
                  <>
                    {undrafted.length > 0 && (
                      <button type="button" className="primary btn-sm markup-draft-button" disabled={busy || loading} onClick={draftAll}>
                        <Icon name="spark" />
                        Draft replies to {plural(undrafted.length, "comment")}
                      </button>
                    )}
                    {drafted && <span className="markup-drafted"><Icon name="accept" />{drafted}</span>}
                  </>
                )}
              </div>
            )}
            {shownThreadList.length === 0 ? (
              <p className="muted markup-empty">
                No comments match these filters.{" "}
                <button type="button" className="linklike" onClick={() => unfilter("comments")}>Clear filters</button>
              </p>
            ) : (
              <ol className="markup-list" aria-label={`Comments in ${name}, in document order`} aria-busy={loading || undefined}>
                {shownThreadList.map((thread) => (
                  <ThreadCard key={threadCard(thread)} thread={thread} open={isOpen(thread)} awaiting={awaiting(thread)} response={markup.responses?.[String(thread.root.id)]}
                    reply={replies[thread.root.id]} drafting={drafting.includes(thread.root.id)} steer={steers[thread.root.id] ?? ""} you={actor || "You"}
                    carried={carriedReplies.includes(thread.root.id)} undo={discarded[thread.root.id]?.text} disabled={busy} canDraft={!busy && !loading}
                    failure={refusal?.thread === thread.root.id ? refusal.message : draftErrors[thread.root.id]}
                    cardRef={(node) => { if (node) cards.current.set(threadCard(thread), node); else cards.current.delete(threadCard(thread)) }}
                    actions={replyActions(thread, replies[thread.root.id])} />
                ))}
              </ol>
            )}
          </>
        )}
      </div>

      {(listed.length > 0 || open.length > 0) && (
        <div className="markup-apply">
          {notice && <p className="markup-notice"><Icon name="warning" />{notice}</p>}
          {(listed.length > 0 || unapproved > 0) && (
            <div className="markup-progress">
              {listed.length > 0 && (
                <span className="markup-progress-bar" aria-hidden="true">
                  <span style={{ width: `${Math.round(((listed.length - undecided) / listed.length) * 100)}%` }} />
                </span>
              )}
              <p>
                {listed.length > 0 && (undecided > 0
                  ? <><strong>{listed.length - undecided} of {listed.length}</strong> decided</>
                  : <span className="markup-done"><Icon name="accept" />{listed.length === 1 ? "The change is decided" : `All ${listed.length} changes decided`}</span>)}
                {listed.length > 0 && elsewhere > 0 && <span className="markup-tally-word"> · {elsewhere} to resolve in Word</span>}
                {(undecided > 0 || unapproved > 0) && (
                  <span className="markup-todo">
                    {listed.length > 0 && " · "}Still to do:{" "}
                    {undecided > 0 && <button type="button" className="linklike" aria-label={`Show the ${plural(undecided, "change")} to decide`} onClick={showUndecided}>{plural(undecided, "change")}</button>}
                    {undecided > 0 && unapproved > 0 && " · "}
                    {unapproved > 0 && <button type="button" className="linklike" aria-label={`Show the ${plural(unapproved, "draft")} to approve`} onClick={showDrafts}>{plural(unapproved, "draft")} to approve</button>}
                  </span>
                )}
              </p>
            </div>
          )}
          {attorney === undefined || savedName ? null : (
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
            <div className="version-confirm" role="group" aria-label="Save a new version" onKeyDown={(event) => {
              if (event.key !== "Escape" || busy) return
              event.preventDefault()
              event.stopPropagation()
              cancel()
            }}>
              <p id={`${ids}-confirm`}>{confirmText}</p>
              {!savedName && <p className="markup-actor">Saved as: <strong>{actor}</strong></p>}
              <span className="upload-actions">
                <button type="button" className="primary btn-sm" autoFocus aria-describedby={`${ids}-confirm`} disabled={busy || loading} onClick={apply}>
                  {busy ? "Saving…" : "Save new version"}
                </button>
                <button type="button" className="icon-btn" disabled={busy} onClick={cancel}>
                  Cancel
                </button>
              </span>
            </div>
          ) : (
            <div className="markup-save">
              <button type="button" ref={applyButton} className="primary markup-apply-button" disabled={!canApply} aria-describedby={blocked ? `${ids}-blocked` : undefined} onClick={review}>
                Save decisions as a new version…
              </button>
              {savedName && <p className="markup-actor">Saved as <strong>{savedName}</strong></p>}
            </div>
          )}
          {refusalShown && <p ref={refusalText} tabIndex={-1} className="version-error markup-error" role="alert"><Icon name="warning" /><span>{refusalShown}</span></p>}
        </div>
      )}
    </aside>
    </VersionContext.Provider>
    </ToneContext.Provider>
  )
}

// Above the page after a Save (#198): the references and defined terms its
// decisions broke, each as the check worded it, or one line when they could
// not be checked; nothing for a clean check. It goes with the next action.
export function ConsistencyNotice({ consistency }: { consistency: NonNullable<MarkupDecisionResult["consistency"]> }) {
  const id = useId()
  if (consistency.status === "not_checked") return <p role="status" className="markup-consistency-unchecked">References and defined terms were not checked for this version — check them yourself before sending.</p>
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

type IconName = "accept" | "reject" | "hold" | "insert" | "delete" | "move" | "format" | "other" | "chevron" | "close" | "spark" | "warning" | "pending" | "info" | "resolved" | "filter"
const ICONS: Record<IconName, ReactNode> = {
  accept: <path d="M5 12.5 10 17 19 7" />,
  reject: <path d="M6 6l12 12M18 6 6 18" />,
  hold: <><path d="M4 7h16M4 12h16" /><path d="M4 17h9" strokeDasharray="2 3" /></>,
  insert: <path d="M12 5v14M5 12h14" />,
  delete: <path d="M5 12h14" />,
  move: <path d="M12 4v16M8 8l4-4 4 4M8 16l4 4 4-4" />,
  format: <path d="M5 19 12 5l7 14M8 14h8" />,
  other: <path d="M4 12h4l2-5 4 10 2-5h4" />,
  chevron: <path d="m6 9 6 6 6-6" />,
  close: <path d="M6 6l12 12M18 6 6 18" />,
  spark: <path d="M12 3v4M12 17v4M3 12h4M17 12h4M6 6l2.5 2.5M15.5 15.5 18 18M6 18l2.5-2.5M15.5 8.5 18 6" />,
  warning: <><path d="M12 3 2 21h20L12 3z" /><path d="M12 10v5M12 18h.01" /></>,
  pending: <><circle cx="12" cy="12" r="9" /><path d="M12 7v5l3 2" /></>,
  info: <><circle cx="12" cy="12" r="9" /><path d="M12 11v6M12 7.5h.01" /></>,
  resolved: <><circle cx="12" cy="12" r="9" /><path d="m8 12.5 3 3 5-6" /></>,
  filter: <path d="M4 5h16l-6 7.5V18l-4 2v-7.5L4 5z" />,
}

function Icon({ name }: { name: IconName }) {
  return (
    <svg className="markup-icon" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
      {ICONS[name]}
    </svg>
  )
}

// One labelled filter: "All …" first, then each choice with how many it shows.
function FilterSelect({ id, label, all, value, options, onChange }: {
  id: string
  label: string
  all: string
  value: string
  options: Option[]
  onChange: (value: string) => void
}) {
  return (
    <div className="markup-filter">
      <label htmlFor={id}>{label}</label>
      <select id={id} value={value} className={value ? "is-active" : undefined} onChange={(event) => onChange(event.target.value)}>
        <option value="">{all}</option>
        {options.map((option) => (
          <option key={option.value} value={option.value} disabled={!option.count && option.value !== value}>
            {option.label} ({option.count})
          </option>
        ))}
      </select>
    </div>
  )
}

// A tab's one-line guide, with the rest a click away.
function Intro({ id, lead, open, onToggle, children }: { id: string; lead: string; open: boolean; onToggle: () => void; children: ReactNode }) {
  return (
    <div className="markup-intro">
      <p>
        {lead}{" "}
        <button type="button" className="linklike" aria-expanded={open} aria-controls={id} onClick={onToggle}>How this works</button>
      </p>
      <div id={id} className="markup-help" hidden={!open}>{children}</div>
    </div>
  )
}

// A tab's line: Filters, which opens the tab's dropdowns as Word's Show Markup
// opens Specific People; how many cards show, said again whenever it changes;
// the tab's bulk action. Closed, the filters in use stay in view as chips,
// each removing its own. Escape closes the dropdowns.
function FilterBar({ id, label, open, chips, shown, total, noun, action, toggleRef, groupRef, onToggle, onRemove, onClear, children }: {
  id: string
  label: string
  open: boolean
  chips: { key: string; label: string }[]
  shown: number
  total: number
  noun: string
  action?: ReactNode
  toggleRef: (node: HTMLButtonElement | null) => void
  groupRef: (node: HTMLDivElement | null) => void
  onToggle: () => void
  onRemove: (key: string) => void
  onClear: () => void
  children: ReactNode
}) {
  return (
    <>
      <div className="markup-toolbar">
        <button type="button" ref={toggleRef} className={`markup-filters-toggle${chips.length ? " is-active" : ""}`} aria-expanded={open} aria-controls={id} onClick={onToggle}
          onKeyDown={(event) => {
            // Escape closes the open dropdowns, not the viewer.
            if (event.key !== "Escape" || !open) return
            event.preventDefault()
            event.stopPropagation()
            onToggle()
          }}>
          <Icon name="filter" />
          Filters
          {chips.length > 0 && <span className="markup-filters-count"><span aria-hidden="true">·</span> {chips.length}<span className="sr-only"> in use</span></span>}
        </button>
        <p className="markup-showing" role="status">
          {chips.length ? <>Showing <strong>{shown}</strong> of {total}</> : <>Showing all {total}</>}
          <span className="sr-only"> {total === 1 ? noun : `${noun}s`}</span>
        </p>
        {action}
      </div>
      <div id={id} ref={groupRef} className="markup-filters" role="group" aria-label={label} hidden={!open} onKeyDown={(event) => {
        if (event.key !== "Escape") return
        event.preventDefault()
        event.stopPropagation()
        onToggle()
      }}>
        {children}
        {chips.length > 0 && <button type="button" className="linklike markup-filters-clear" onClick={onClear}>Clear filters</button>}
      </div>
      {!open && chips.length > 0 && (
        <ul className="markup-filter-chips" aria-label="Filters in use">
          {chips.map((chip) => (
            <li key={chip.key}>
              <button type="button" className="markup-filter-chip" aria-label={`Remove the filter ${chip.label}`} onClick={() => onRemove(chip.key)}>
                <span>{chip.label}</span>
                <Icon name="close" />
              </button>
            </li>
          ))}
          <li><button type="button" className="linklike" onClick={onClear}>Clear filters</button></li>
        </ul>
      )}
    </>
  )
}

// One line, cut with an ellipsis where it runs long, which More opens in full.
// A screen reader reads it all either way.
function OneLine({ className, children }: { className: string; children: ReactNode }) {
  const text = useRef<HTMLSpanElement>(null)
  const [long, setLong] = useState(false)
  const [open, setOpen] = useState(false)
  useEffect(() => {
    const node = text.current
    if (!node || typeof ResizeObserver === "undefined") return
    const measure = () => setLong(node.scrollWidth > node.clientWidth + 1)
    measure()
    const observer = new ResizeObserver(measure)
    observer.observe(node)
    return () => observer.disconnect()
  }, [])
  return (
    <p className={`${className} markup-line${open ? " is-open" : ""}`}>
      <span ref={text} className="markup-line-text">{children}</span>
      {(long || open) && (
        <button type="button" className="linklike markup-line-toggle" aria-expanded={open} aria-label={open ? "Show less" : "Show in full"} onClick={() => setOpen(!open)}>
          {open ? "Less" : "More"}
        </button>
      )}
    </p>
  )
}

// An author's initials in their colour, as Word marks a comment's author.
function Avatar({ name, initials, small = false }: { name: string; initials?: string | null; small?: boolean }) {
  const tone = useContext(ToneContext)(name)
  return (
    <span className={`markup-avatar ${toneClass(tone)}${small ? " is-small" : ""}`} aria-hidden="true">
      {initialsOf(name, initials)}
    </span>
  )
}

// Who and when: the avatar, the name in bold and the date beneath it, or
// beside it on a change card.
function Byline({ name, initials, date, small, inline, children }: { name: string; initials?: string | null; date: string | null; small?: boolean; inline?: boolean; children?: ReactNode }) {
  return (
    <div className={`markup-byline${inline ? " is-inline" : ""}`}>
      <Avatar name={name} initials={initials} small={small} />
      <span className="markup-byline-text">
        <span className="change-author">{authorName(name)}</span>
        {date && <time dateTime={date}>{formatDateTime(date)}</time>}
      </span>
      {children}
    </div>
  )
}

// A card's state at its top right while something is left to do on it.
function Chip({ tone, icon, children }: { tone: "todo" | "done" | "quiet"; icon: IconName; children: ReactNode }) {
  return <span className={`markup-chip is-${tone}`}><Icon name={icon} />{children}</span>
}

// What the drafting run did with this change or thread, in one line that
// says it is doc.haus's; nothing for an item it left open. A run that read an
// earlier version says which. The reason is model text, shown as text.
function RunResponse({ response }: { response?: MarkupResponse }) {
  const current = useContext(VersionContext)
  const reason = response?.reason?.trim()
  const answered = response?.question?.status === "answered"
  const lines: Record<string, [string, string]> = {
    amended: ["doc.haus draft:", "changed in response"],
    rejected: ["doc.haus draft:", reason ? `not changed — ${reason}` : "not changed"],
    deferred: ["doc.haus draft:", "left for you to decide"],
    question: ["doc.haus draft:", answered ? "you answered its question on this" : "waiting for your answer to its question"],
    recommendation: ["doc.haus suggests:", reason || "a change here"],
  }
  const line = response && lines[response.disposition]
  if (!line) return null
  const from = typeof response.from_version === "number" && response.from_version !== current ? response.from_version : undefined
  return (
    <p className="markup-response">
      <Icon name="spark" />
      <span>
        <span className="markup-response-label">{line[0]}</span> {line[1]}
        {from !== undefined && <span className="markup-response-from"> (from the run on version {from})</span>}
      </span>
    </p>
  )
}

function Flags({ findings }: { findings: MarkupFinding[] }) {
  const notes = flagNotes(findings)
  if (!notes.length) return null
  return (
    <div className="markup-flag">
      <Icon name="warning" />
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
      return <OneLine className="markup-place">Moved from {PART[revision.scope]}{side(from)} to {side(to)}</OneLine>
    // Otherwise its places come in document order, which does not say which
    // one it left; the document shows that.
    const labelled = places.some(listLabel)
    const where = joined(places.map((place) => (labelled ? placeName(revision.scope, place)! : place.paragraph!)))
    return <OneLine className="markup-place">Old and new places: {labelled ? where : `${PART[revision.scope]}paragraphs ${where}`}</OneLine>
  }
  return (
    <>
      {places.map((place, index) => {
        const context = preview(place.context)
        const named = revision.scope === "comment" ? commentPlace(place, comments, !!context) : placeName(revision.scope, place)
        // A move whose places are quoted says which one each is.
        const where = named && revision.scope !== "comment" && place.side ? `${place.side} ${named}` : named
        return (
          <OneLine className="markup-place" key={index}>
            {where && <span className="markup-place-name">{capital(where)}</span>}
            {where && context && " · "}
            {context && <q>{context}</q>}
          </OneLine>
        )
      })}
    </>
  )
}

// One change, as Word's Review pane shows it: what kind, who and when, the
// text. Accept, Reject and Leave tracked are one radio group (one Tab stop,
// arrow keys move and choose), with nothing chosen until the lawyer decides;
// the lawyer's own edit starts on Leave tracked and says what Accept does to
// it. A refused change carries the refusal and a one-click Leave tracked.
function ChangeCard({ revision, comments, state, own, disabled, carried, response, failure, holdHint, cardRef, onDecide, onHold }: {
  revision: MarkupRevision
  comments: MarkupComment[]
  state: ChangeState
  // The lawyer's own edit: a doc.haus proposal accepted into the file.
  own: boolean
  disabled: boolean
  carried: boolean
  response?: MarkupResponse
  failure?: string
  // The tab's line that says what Leave tracked does.
  holdHint: string
  cardRef: (node: HTMLLIElement | null) => void
  onDecide: (decision: MarkupDecision) => void
  onHold: () => void
}) {
  const id = useId()
  const type = changeType(revision)
  const detail = changeDetail(revision)
  const text = changeText(revision)
  const removes = revision.kind === "delete"
  return (
    <li ref={cardRef} className={`change-card markup-card is-${state}${own ? " is-own" : ""}${failure ? " is-refused" : ""}`}>
      <div className="markup-card-top">
        <span className={`markup-kind is-${type}`}>
          <Icon name={type} />
          {TYPES.find((entry) => entry.value === type)!.label}
        </span>
        {detail && <span className="markup-kind-detail">{detail}</span>}
        {own && <span className="markup-own">Your edit</span>}
        {carried && <span className="version-tag">Carried over</span>}
        {state === "undecided" && <Chip tone="todo" icon="pending">Needs your decision</Chip>}
        {state === "word" && <Chip tone="quiet" icon="info">Resolve in Word</Chip>}
      </div>
      <Byline name={revision.author} date={revision.date} small inline />
      {text && <div className={removes ? "change-old" : revision.kind === "insert" ? "change-new" : "markup-text"}>{text}</div>}
      {revision.family === "paragraph_mark" && revision.kind === "delete" && <p className="markup-hint">Accepting joins this paragraph with the next one.</p>}
      {revision.kind === "format" && decidable(revision) && <p className="markup-hint">Accept keeps the new formatting; Reject restores the old.</p>}
      {own && decidable(revision) && (
        <p className="markup-hint" id={`${id}-own`}>Leave tracked to send it to the other side as a tracked change. Accept makes it final, without tracking.</p>
      )}
      <Places revision={revision} comments={comments} />
      <Flags findings={revision.findings} />
      <RunResponse response={response} />
      {state === "word" ? (
        <p className="markup-word"><Icon name="info" /><span><strong>Resolve in Word.</strong> doc.haus cannot apply this change safely, so it stays as a tracked change.</span></p>
      ) : (
        <fieldset className="markup-decision" aria-label={`Your decision: ${decisionLabel(revision)}`} aria-describedby={failure ? `${id}-refused` : undefined} disabled={disabled}>
          <legend>Your decision</legend>
          <div className="markup-choices">
            {DECISIONS.map((option) => (
              <label key={option.value} className={`markup-choice is-${option.value}`}>
                <input type="radio" name={`${id}-decision`} value={option.value} checked={state === option.value}
                  aria-describedby={option.value === "hold" ? (own ? `${id}-own` : holdHint) : undefined} onChange={() => onDecide(option.value)} />
                <Icon name={option.value} />
                <span>{option.label}</span>
              </label>
            ))}
          </div>
        </fieldset>
      )}
      {failure && (
        <p className="version-error markup-error" role="alert" tabIndex={-1}>
          <Icon name="warning" />
          <span>
            <span id={`${id}-refused`}>{failure}</span>
            {state !== "hold" && (
              <>
                {" "}
                <button type="button" className="linklike" disabled={disabled} onClick={onHold}>Leave this change tracked</button>
              </>
            )}
          </span>
        </p>
      )}
    </li>
  )
}

// One comment in a thread, as Word shows it: avatar, name and date, the text.
// The thread's first comment carries the card's state beside the author and,
// under the author, the text the comment is on.
function CommentBody({ comment, reply = false, chip, anchor }: { comment: MarkupComment; reply?: boolean; chip?: ReactNode; anchor?: ReactNode }) {
  return (
    <div className={`markup-comment${reply ? " is-reply" : ""}`}>
      <Byline name={comment.author} initials={comment.initials} date={comment.date} small={reply}>{chip}</Byline>
      {anchor}
      <p className="markup-comment-text">{comment.text}</p>
      <Flags findings={comment.findings} />
    </div>
  )
}

// A thread as a Word comment card: the text it is anchored to, quietly, the
// comment and its replies under it, and the reply box while it is open. A
// resolved thread is greyed, as Word greys it, and so is a detached one, whose
// text is no longer in the document; it says so and takes no reply.
function ThreadCard({ thread, open, response, cardRef, ...area }: { thread: Thread; open: boolean; response?: MarkupResponse; cardRef: (node: HTMLLIElement | null) => void } & Omit<ReplyAreaProps, "thread">) {
  const { root, replies } = thread
  const place = root.paragraph !== null ? placeName(root.scope ?? "body", root) : undefined
  const state = !open ? (root.resolved === true ? "resolved" : root.detached ? "detached" : "closed") : area.reply?.approved ? "approved" : area.reply ? "drafted" : area.drafting ? "drafting" : area.awaiting ? "needed" : "answered"
  return (
    <li ref={cardRef} tabIndex={-1} data-comment-id={root.id} className={`change-card markup-card markup-thread is-${state}${area.failure ? " is-refused" : ""}`}>
      <CommentBody comment={root}
        chip={state === "resolved" ? <Chip tone="quiet" icon="resolved">Resolved</Chip>
          : state === "detached" ? <Chip tone="quiet" icon="info">Text removed</Chip>
          : state === "needed" ? <Chip tone="todo" icon="pending">Needs a reply</Chip>
          : state === "drafted" ? <Chip tone="todo" icon="pending">Reply to approve</Chip>
          : state === "approved" ? <Chip tone="done" icon="accept">Reply approved</Chip>
          : undefined}
        anchor={(root.quote || place) && (
          <OneLine className="markup-anchor">
            {root.quote && <><span className="sr-only">On the text </span><q className="markup-quote">{root.quote}</q></>}
            {root.quote && place && " · "}
            {place && <span className="markup-place-name">{capital(place)}</span>}
          </OneLine>
        )} />
      {replies.length > 0 && (
        <ol className="markup-replies" aria-label={`${plural(replies.length, "reply", "replies")}`}>
          {replies.map((reply) => (
            <li key={reply.id}>
              <CommentBody comment={reply} reply />
            </li>
          ))}
        </ol>
      )}
      <RunResponse response={response} />
      {state === "detached" && <p className="markup-hint">The text this comment was on is no longer in the document, so it takes no reply here.</p>}
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
  // Who an approved reply is signed as.
  you: string
  // What Discard or Remove did, with Undo beside it.
  undo?: string
  failure?: string
  disabled: boolean
  canDraft: boolean
  actions: ReplyActions
}

// One open thread's reply, in Word's reply box at the foot of the card:
// offered (Write a reply in the box, Draft reply, an optional steer), drafting
// (Cancel), a draft to edit (Approve reply once nothing is left in brackets,
// Redraft, Discard) or approved (shown as the next reply in the thread, with
// Edit and Remove). The control the panel focuses after an action carries
// data-reply-focus. A thread flagged for planted instructions repeats its flag
// above the draft.
function ReplyArea({ thread, awaiting, reply, drafting, steer, carried, you, undo, failure, disabled, canDraft, actions }: ReplyAreaProps) {
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
    <div className={`markup-reply${reply?.approved ? " is-approved" : ""}`} role="group" aria-label={`Your reply to ${authorName(thread.root.author)}`}>
      {drafting ? (
        <div className="markup-replybox is-busy">
          <span className="markup-reply-status"><span className="markup-spinner" aria-hidden="true" />Drafting a reply…</span>
          <button type="button" className="icon-btn" data-reply-focus onClick={actions.cancel}>Cancel</button>
        </div>
      ) : !reply ? (
        <>
          {undo && (
            <p className="markup-hint">
              {undo}{" "}
              <button type="button" className="linklike" disabled={disabled} onClick={actions.undo}>Undo</button>
            </p>
          )}
          <button type="button" className="markup-replybox" data-reply-focus={awaiting ? undefined : true} disabled={disabled} onClick={actions.write}>Write a reply</button>
          {awaiting && (
            <div className="markup-reply-actions">
              <button type="button" className="icon-btn is-primary" data-reply-focus disabled={!canDraft} onClick={actions.draft}><Icon name="spark" />Draft reply</button>
              {steerToggle}
            </div>
          )}
          {steerField}
        </>
      ) : reply.approved ? (
        <div className="markup-comment is-reply is-pending">
          <Byline name={you} date={null} small>
            <span className="markup-chip is-done"><Icon name="accept" />Approved</span>
          </Byline>
          <p className="markup-reply-final">{reply.text}</p>
          <p className="markup-reply-approved">Reply approved — added when you save</p>
          <div className="markup-reply-actions">
            <button type="button" className="icon-btn" data-reply-focus disabled={disabled} onClick={actions.edit}>Edit</button>
            <button type="button" className="linklike" disabled={disabled} onClick={actions.discard}>Remove</button>
          </div>
        </div>
      ) : (
        <>
          {planted.length > 0 && <Flags findings={planted} />}
          {reply.changed && (
            <p className="markup-reply-changed"><Icon name="warning" /><span>The thread changed since you drafted this — check before approving.</span></p>
          )}
          {undo && (
            <p className="markup-hint">
              {undo}{" "}
              <button type="button" className="linklike" disabled={disabled} onClick={actions.undo}>Undo</button>
            </p>
          )}
          <label className="markup-reply-label" htmlFor={`${id}-text`}>
            Your reply
            {carried && <span className="version-tag">Carried over</span>}
          </label>
          <textarea id={`${id}-text`} className="markup-reply-text" data-reply-focus value={reply.text} rows={4} maxLength={REPLY_MAX} disabled={disabled}
            aria-describedby={[reply.open_items.length ? `${id}-items` : "", hold ? `${id}-hold` : ""].filter(Boolean).join(" ") || undefined}
            onChange={(event) => actions.change(event.target.value)} />
          {reply.open_items.length > 0 && (
            <div className="markup-open-items" id={`${id}-items`}>
              <p><Icon name="info" />Check before sending</p>
              <ul>
                {reply.open_items.map((item, index) => <li key={index}>{item}</li>)}
              </ul>
            </div>
          )}
          {hold && <p id={`${id}-hold`} className="markup-hint markup-hold">{hold}</p>}
          <div className="markup-reply-actions">
            <button type="button" className="primary btn-sm" disabled={disabled || !!hold} aria-describedby={hold ? `${id}-hold` : undefined} onClick={actions.approve}>
              <Icon name="accept" />Approve reply
            </button>
            {awaiting && <button type="button" className="icon-btn" disabled={!canDraft} onClick={actions.draft}>Redraft</button>}
            <button type="button" className="linklike" disabled={disabled} onClick={actions.discard}>Discard</button>
            {steerToggle}
          </div>
          {steerField}
        </>
      )}
      {failure && <p className="version-error markup-error" role="alert" tabIndex={-1}><Icon name="warning" /><span>{failure}</span></p>}
    </div>
  )
}
