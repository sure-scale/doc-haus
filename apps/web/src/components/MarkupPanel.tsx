import { useEffect, useId, useMemo, useRef, useState, type KeyboardEvent } from "react"
import { humanizeError } from "../api/ingest/errors"
import {
  actorSent,
  applyMarkupDecisions,
  MarkupRefusedError,
  type DocumentMarkup,
  type MarkupComment,
  type MarkupDecision,
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
// Every name, change, quote and comment is the other side's text and renders as
// plain text.

type Thread = { root: MarkupComment; replies: MarkupComment[] }
type Item =
  | { type: "change"; revision: MarkupRevision; order: number[] }
  | { type: "thread"; thread: Thread; order: number[] }
type Kind = "changes" | "formatting" | "comments"
type Drafts = Record<string, Exclude<MarkupDecision, "hold">>
// A refused Apply: the message, the service's code and the listed changes it
// names, whose cards carry the message.
type Refusal = { message: string; code?: string; ids: string[] }
// Where focus goes once the panel has rendered: a change's decision, or Apply.
type Focus = { card: string } | "apply"

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

const plural = (count: number, one: string, many = `${one}s`) => `${count} ${count === 1 ? one : many}`

// Plain words for a new reading of a file the lawyer had decided changes in.
function carriedNote(name: string, carried: number, dropped: number, refused: boolean) {
  if (!refused && !carried && !dropped) return undefined
  return [
    refused ? `${name} changed after this list was read, so nothing was applied.` : `${name} changed since these decisions were made.`,
    carried && `${plural(carried, "decision")} carried over — check ${carried === 1 ? "it" : "them"} before applying.`,
    dropped && `${plural(dropped, "decision")} could not be carried over and ${dropped === 1 ? "is" : "are"} back on Hold.`,
    refused && !carried && "Decide the changes again.",
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
  // A new version was saved: the viewer reloads the document and its versions.
  onApplied: () => void
  // The service refused the batch: read the file and its markup again.
  onRefetch: () => void
  // doc.haus's own proposals block applying: show them instead of this panel.
  onShowProposals?: () => void
}) {
  const sha256 = markup.version.sha256
  const key = draftKey(matterId, name)
  const [initial] = useState(() => carryDrafts(readStored(key), markup))
  // The document and bytes the drafts belong to.
  const [reading, setReading] = useState({ key, base: sha256 })
  const [drafts, setDrafts] = useState<Drafts>(initial.drafts)
  // Decisions carried over from other bytes, until the lawyer looks at each.
  const [carried, setCarried] = useState<string[]>(initial.carried)
  const [notice, setNotice] = useState(() => carriedNote(name, initial.carried.length, initial.dropped, false))
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
  const cards = useRef(new Map<string, HTMLLIElement>())
  const bulkButtons = useRef(new Map<MarkupDecision, HTMLButtonElement>())
  const ids = useId()
  const toast = useToast()

  // A new reading of other bytes keeps the drafts for changes it still lists,
  // and says so; one after this panel's own Apply starts afresh. Another
  // document takes its own drafts.
  if (reading.key !== key || reading.base !== sha256) {
    const fresh = reading.key !== key
    const next = carryDrafts(fresh ? readStored(key) : appliedFrom === reading.base ? undefined : { base: reading.base, decisions: drafts }, markup)
    setReading({ key, base: sha256 })
    setDrafts(next.drafts)
    setCarried(next.carried)
    setNotice(carriedNote(name, next.carried.length, next.dropped, !fresh && refusal?.code === "base_changed"))
    setAppliedFrom(undefined)
    setUndo(undefined)
    setRefusal(undefined)
  }

  useEffect(() => {
    if (Object.keys(drafts).length) writeBrowserValue(key, JSON.stringify({ base: reading.base, decisions: drafts }))
    else removeBrowserValue(key)
  }, [key, reading.base, drafts])

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
  // chosen decision, or Apply once the file has been read again (the
  // proposals button or the message while Apply cannot be used).
  useEffect(() => {
    if (!focus || busy) return
    const card = typeof focus === "object" ? cards.current.get(focus.card) : undefined
    if (card) {
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
  const blocked = proposalsPending > 0
    ? `Accept or reject doc.haus's ${plural(proposalsPending, "proposed change")} on this document first.`
    : !accepted && !rejected ? "Mark at least one change Accept or Reject to apply." : undefined
  const canApply = !blocked && !busy && !loading && attorney !== undefined
  const proposalsFirst = proposalsPending > 0 || refusal?.code === "proposals_pending"
  // A refusal is shown on the cards it names; otherwise, or while they are
  // filtered out, beside Apply. The proposals line already says what a
  // proposals refusal says once the proposals are known.
  const shownIds = new Set(shown.flatMap((item) => (item.type === "change" ? [item.revision.id] : [])))
  const onCards = !!refusal && refusal.ids.some((id) => shownIds.has(id))
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

  const toggle = <T,>(list: T[], value: T) => (list.includes(value) ? list.filter((entry) => entry !== value) : [...list, value])

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
    // The panel reads the file again itself; there is nothing to refresh.
    const message = cause.code === "base_changed" ? `${name} changed after this list was read, so nothing was applied.` : cause.message
    setRefusal({ message, code: cause.code, ids: named })
    if (!named.length) return "apply"
    if (!named.every((id) => shownIds.has(id))) {
      setAuthors([])
      setKinds([])
    }
    return { card: named[0]! }
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
      const result = await applyMarkupDecisions(matterId, name, {
        base_sha256: sha256,
        decisions: listed.map((revision) => ({ id: revision.id, decision: drafts[revision.id] ?? "hold" })),
      }, recorded)
      removeBrowserValue(key)
      setAppliedFrom(sha256)
      setCarried([])
      setNotice(undefined)
      setUndo(undefined)
      setConfirming(false)
      const counts = result.counts ?? {}
      const tally = [
        `${counts.accepted ?? accepted} accepted`,
        `${counts.rejected ?? rejected} rejected`,
        `${counts.held ?? held} held`,
      ].join(", ")
      // The tally names the held changes; the version's own notice about
      // tracked changes left in it would only repeat them.
      toast("success", `${result.version ? `Saved as version ${result.version.number}` : "Saved a new version"}: ${tally}.`)
      if (counts.consumed) toast("info", `${plural(counts.consumed, "change was", "changes were")} resolved along with another change.`)
      onApplied()
    } catch (cause) {
      setConfirming(false)
      after = refused(cause)
    } finally {
      setBusy(false)
      if (after) setFocus(after)
      else heading.current?.focus()
    }
  }

  const confirmText = `This saves a new version of ${name} with ${plural(accepted, "change")} accepted and ${rejected} rejected.` +
    `${held ? ` ${plural(held, "held change stays", "held changes stay")} as tracked ${held === 1 ? "change" : "changes"}.` : ""} The current version stays in the history.`

  return (
    <aside id="document-markup" className="changes-panel markup-panel" aria-labelledby={`${ids}-title`}>
      <div className="changes-head markup-head">
        <h3 id={`${ids}-title`} ref={heading} tabIndex={-1}>
          Their changes <span className="compare-count">{items.length}</span>
        </h3>
        <p>Tracked changes and comments in this Word file. Choose Accept, Reject or Hold for each change, then apply them together. Held changes stay as tracked changes.</p>
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
              <ThreadCard key={`comment-${item.thread.root.id}`} thread={item.thread} />
            ),
          )}
        </ol>
      )}

      {listed.length > 0 && (
        <div className="markup-apply">
          {notice && <p className="markup-hint" role="status">{notice}</p>}
          <p className="markup-tally">
            {accepted} to accept · {rejected} to reject · {held} held
            {elsewhere > 0 && <span> · {elsewhere} to resolve in Word</span>}
          </p>
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
  // A move's places come in document order, which does not say which one it
  // left; the document shows that.
  if (revision.kind === "move" && places.length > 1 && places.every((place) => place.paragraph !== null && !preview(place.context))) {
    const labelled = places.some(listLabel)
    const names = places.map((place) => (labelled ? placeName(revision.scope, place)! : place.paragraph!))
    const where = `${names.slice(0, -1).join(", ")} and ${names.at(-1)}`
    return <p className="markup-place">Old and new places: {labelled ? where : `${PART[revision.scope]}paragraphs ${where}`}</p>
  }
  return (
    <>
      {places.map((place, index) => {
        const context = preview(place.context)
        const where = revision.scope === "comment" ? commentPlace(place, comments, !!context) : placeName(revision.scope, place)
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

function ThreadCard({ thread }: { thread: Thread }) {
  const { root, replies } = thread
  return (
    <li className="change-card markup-card markup-thread">
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
      {root.paragraph !== null && <p className="markup-place"><span>{capital(placeName(root.scope ?? "body", { paragraph: root.paragraph })!)}</span></p>}
    </li>
  )
}
