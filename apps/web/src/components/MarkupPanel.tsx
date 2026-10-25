import { useEffect, useId, useMemo, useRef, useState } from "react"
import { humanizeError } from "../api/ingest/errors"
import {
  applyMarkupDecisions,
  MarkupRefusedError,
  type DocumentMarkup,
  type MarkupComment,
  type MarkupDecision,
  type MarkupFinding,
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
// are kept in this browser per version of the file, so closing the viewer
// loses nothing. Every name, change, quote and comment is the other side's
// text and renders as plain text.

type Thread = { root: MarkupComment; replies: MarkupComment[] }
type Item =
  | { type: "change"; revision: MarkupRevision; order: [number, number, number, number] }
  | { type: "thread"; thread: Thread; order: [number, number, number, number] }
type Kind = "changes" | "formatting" | "comments"
type Drafts = Record<string, Exclude<MarkupDecision, "hold">>

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
const SCOPES = ["body", "header", "footer", "footnote", "endnote"] as const
const PART: Record<MarkupRevision["scope"], string> = { body: "Paragraph", header: "Header, paragraph", footer: "Footer, paragraph", footnote: "Footnote, paragraph", endnote: "Endnote, paragraph" }
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
// paragraph's changes before its comments, then as listed.
function documentOrder(markup: DocumentMarkup): Item[] {
  const last = Number.MAX_SAFE_INTEGER
  const items: Item[] = [
    ...markup.revisions.map((revision, index): Item => ({
      type: "change",
      revision,
      order: [SCOPES.indexOf(revision.scope), revision.locations.find((place) => place.paragraph !== null)?.paragraph ?? last, 0, index],
    })),
    ...commentThreads(markup.comments).map((thread, index): Item => ({ type: "thread", thread, order: [0, thread.root.paragraph ?? last, 1, index] })),
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
  if (revision.kind === "format") return "Formatting changed"
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

// Plain words for what screening found, once per kind of finding.
export function flagNotes(findings: MarkupFinding[]) {
  const notes: string[] = []
  if (findings.some((finding) => !HIDDEN.has(finding.rule))) notes.push("Contains text that looks like an instruction — doc.haus did not follow it.")
  if (findings.some((finding) => HIDDEN.has(finding.rule))) notes.push("Hidden characters were removed from this text.")
  return notes
}

// Drafts are kept per matter, document and version bytes: the ids belong to
// one version's bytes, so a new version starts with every change on Hold.
const draftKey = (matterId: string, name: string, sha256: string) =>
  `dh.markup-decisions.v1/${encodeURIComponent(matterId)}/${encodeURIComponent(name)}/${sha256}`

function readDrafts(key: string, revisions: MarkupRevision[]): Drafts {
  const stored = readBrowserJSON(key)
  if (!stored || typeof stored !== "object") return {}
  const listed = new Set(revisions.filter(decidable).map((revision) => revision.id))
  return Object.fromEntries(
    Object.entries(stored as Record<string, unknown>).filter(
      (entry): entry is [string, "accept" | "reject"] => listed.has(entry[0]) && (entry[1] === "accept" || entry[1] === "reject"),
    ),
  )
}

const plural = (count: number, one: string, many = `${one}s`) => `${count} ${count === 1 ? one : many}`

export default function MarkupPanel({
  matterId,
  name,
  markup,
  loading = false,
  proposalsPending = 0,
  onApplied,
  onRefetch,
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
}) {
  const sha256 = markup.version.sha256
  const key = draftKey(matterId, name, sha256)
  const [drafts, setDrafts] = useState<Drafts>(() => readDrafts(key, markup.revisions))
  const [draftsKey, setDraftsKey] = useState(key)
  const [authors, setAuthors] = useState<string[]>([])
  const [kinds, setKinds] = useState<Kind[]>([])
  // undefined while loading; "" when no name is saved, so the bar asks once.
  const [attorney, setAttorney] = useState<string>()
  const [nameDraft, setNameDraft] = useState("")
  const [confirming, setConfirming] = useState(false)
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState<string>()
  const heading = useRef<HTMLHeadingElement>(null)
  const applyButton = useRef<HTMLButtonElement>(null)
  const nameInput = useRef<HTMLInputElement>(null)
  const ids = useId()
  const toast = useToast()

  // A new reading of the file (another version, or the same bytes again) takes
  // that version's drafts.
  if (draftsKey !== key) {
    setDraftsKey(key)
    setDrafts(readDrafts(key, markup.revisions))
  }

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

  function decide(changes: MarkupRevision[], decision: MarkupDecision) {
    setError(undefined)
    setConfirming(false)
    const next = { ...drafts }
    for (const revision of changes) {
      if (decision === "hold") delete next[revision.id]
      else next[revision.id] = decision
    }
    setDrafts(next)
    if (Object.keys(next).length) writeBrowserValue(key, JSON.stringify(next))
    else removeBrowserValue(key)
  }

  const toggle = <T,>(list: T[], value: T) => (list.includes(value) ? list.filter((entry) => entry !== value) : [...list, value])

  function review() {
    setError(undefined)
    if (!actor) {
      setError("Add your name so it is recorded with these decisions.")
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

  async function apply() {
    setBusy(true)
    setError(undefined)
    try {
      let recorded = savedName
      if (!recorded) {
        try {
          recorded = (await savePreferences({ attorney: actor })).attorney?.trim() || actor
          setAttorney(recorded)
        } catch (cause) {
          setError(`Your name could not be saved. ${humanizeError(cause)}`)
          return
        }
      }
      const result = await applyMarkupDecisions(matterId, name, {
        base_sha256: sha256,
        decisions: listed.map((revision) => ({ id: revision.id, decision: drafts[revision.id] ?? "hold" })),
      }, recorded)
      removeBrowserValue(key)
      setConfirming(false)
      const counts = result.counts ?? {}
      const tally = [
        `${counts.accepted ?? accepted} accepted`,
        `${counts.rejected ?? rejected} rejected`,
        `${counts.held ?? held} held`,
      ].join(", ")
      toast("success", `${result.version ? `Saved as version ${result.version.number}` : "Saved a new version"}: ${tally}.`)
      if (counts.consumed) toast("info", `${plural(counts.consumed, "change was", "changes were")} resolved along with another change.`)
      if (result.notice) toast("warning", result.notice)
      onApplied()
    } catch (cause) {
      setConfirming(false)
      setError(cause instanceof MarkupRefusedError ? cause.message : humanizeError(cause, "The decisions could not be applied. Try again."))
      // The file or its changes moved on: read them again before another try.
      if (cause instanceof MarkupRefusedError && cause.status === 409) onRefetch()
    } finally {
      setBusy(false)
      heading.current?.focus()
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
              <button type="button" key={decision.value} className="icon-btn" disabled={busy} onClick={() => decide(shownChanges, decision.value)}>
                {decision.label}
              </button>
            ))}
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
              <ChangeCard key={item.revision.id} revision={item.revision} decision={drafts[item.revision.id] ?? "hold"} disabled={busy}
                onDecide={(decision) => decide([item.revision], decision)} />
            ) : (
              <ThreadCard key={`comment-${item.thread.root.id}`} thread={item.thread} />
            ),
          )}
        </ol>
      )}

      {listed.length > 0 && (
        <div className="markup-apply">
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
          {error && <p className="version-error" role="alert">{error}</p>}
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

function Places({ revision }: { revision: MarkupRevision }) {
  const text = changeText(revision)
  const places = revision.locations.filter((place) => place.paragraph !== null || place.context)
  if (!places.length) return null
  if (revision.kind === "move" && places.length > 1 && places.every((place) => place.paragraph !== null)) {
    const numbers = places.map((place) => place.paragraph!)
    return <p className="markup-place">Moved between {PART[revision.scope].toLowerCase()} {numbers.slice(0, -1).join(", ")} and {numbers.at(-1)}</p>
  }
  // The preview repeats a whole inserted or moved paragraph; it is left out then.
  const preview = (context: string) => {
    const head = context.replace(/…$/, "")
    return head && !text.startsWith(head) && !head.startsWith(text) ? context : ""
  }
  return (
    <>
      {places.map((place, index) => {
        const context = preview(place.context)
        return (
          <p className="markup-place" key={index}>
            {place.paragraph !== null && <span>{PART[revision.scope]} {place.paragraph}</span>}
            {context && <q>{context}</q>}
          </p>
        )
      })}
    </>
  )
}

function ChangeCard({ revision, decision, disabled, onDecide }: {
  revision: MarkupRevision
  decision: MarkupDecision
  disabled: boolean
  onDecide: (decision: MarkupDecision) => void
}) {
  const label = changeLabel(revision)
  const text = changeText(revision)
  const removes = revision.kind === "delete"
  return (
    <li className={`change-card markup-card is-${decision}`}>
      <div className="markup-card-head">
        <span className={`markup-kind is-${revision.kind}`}>{label}</span>
        <span className="change-author">{authorName(revision.author)}</span>
        {revision.date && <time dateTime={revision.date}>{formatDateTime(revision.date)}</time>}
      </div>
      {text && <div className={removes ? "change-old" : revision.kind === "insert" ? "change-new" : "markup-text"}>{text}</div>}
      {revision.family === "paragraph_mark" && revision.kind === "delete" && <p className="markup-hint">Accepting joins this paragraph with the next one.</p>}
      <Places revision={revision} />
      <Flags findings={revision.findings} />
      {decidable(revision) ? (
        <div className="markup-decision" role="group" aria-label={`Decision on this change: ${label}`}>
          {DECISIONS.map((option) => (
            <button type="button" key={option.value} className={`markup-choice is-${option.value}`} aria-pressed={decision === option.value} disabled={disabled}
              onClick={() => onDecide(option.value)}>
              {option.label}
            </button>
          ))}
        </div>
      ) : (
        <p className="markup-word"><strong>Resolve in Word.</strong> doc.haus cannot apply this change safely, so it stays as a tracked change.</p>
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
      {root.paragraph !== null && <p className="markup-place"><span>Paragraph {root.paragraph}</span></p>}
    </li>
  )
}
