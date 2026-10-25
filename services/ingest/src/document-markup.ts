import { Hono, type Context } from "hono"
import { existsSync, readFileSync } from "node:fs"
import path from "node:path"
import { withArtifactMutationLocksAsync } from "../../../dochaus/capabilities/artifacts/mutation-lock"
import { diagnosticErrorMetadata } from "../../../dochaus/lib/diagnostic-metadata"
import { addCommentReplies, CommentReplyError, replyInitials, type CommentReply } from "../../../dochaus/lib/docx-comment-replies"
import { readDocxMarkup, type DocxMarkup } from "../../../dochaus/lib/docx-markup"
import { getDocumentRole, listPendingRedlines, openDb } from "./db"
import { findVersionBySha, oneLine, readDocumentVersions, readVersion, readVersionBytes, sha256Hex, type VersionWrite } from "./document-versions"
import { matterDocumentName, SourceDocumentConflictError } from "./ingest"
import { applyMarkupDecisions, MarkupDecisionError, type MarkupDecision } from "./markup-decisions"
import { readDraftingPreferences } from "./preferences"
import { readDocumentTarget } from "./registered-document"

// A counterparty's tracked changes and comments in an uploaded Word source
// (#200): the list the viewer's triage panel shows, and the lawyer's accept,
// reject and hold decisions applied as the document's next version. The list is
// a computed view of one version's bytes (readDocxMarkup); nothing is stored.
// Decisions write through the revision writer as an accept, so the source
// re-ingests as any revision does: held changes stay labelled in search, the
// others become plain text or go. Their event names each change's kind, author,
// date, decision and outcome and holds none of its text, so a later redaction
// leaves no copy of removed wording behind. The same Apply adds the replies the
// lawyer approved to the other side's open comment threads (scope 5), as Word
// replies signed with the lawyer's name, in the same version; replies alone are
// a revision with no event, as the event kinds have none for a reply. Only the
// lawyer's request reaches the decisions route; no agent tool calls it.
// doc.haus deliverables keep their own review path.
type Options = {
  directory: (id: string) => string
  /** The proposal accept route's lock set for the matter file `file`. */
  locks: (directory: string, file: string) => string[]
  /** The revision writer: `bytes` become the next version of source `name`. */
  write: (directory: string, name: string, bytes: Buffer, version: VersionWrite) => Promise<object>
}

const DECISIONS = new Set<string>(["accept", "reject", "hold"])
const REPLY_LIMIT = 4_000

export function documentMarkupRoutes(options: Options) {
  const app = new Hono()

  // The changes and comments of the source's current version (the bytes its
  // working file holds), or of the kept version `version`. `current`: these are
  // the working file's bytes, the only ones decisions apply to.
  app.get("/:id/documents/markup", async (c) => {
    const directory = options.directory(c.req.param("id"))
    const name = wordName(c.req.query("name"))
    if (!name) return c.json({ error: "Changes and comments are listed for Word (.docx) documents only." }, 400)
    const target = sourceTarget(directory, name)
    if (!target) return notSource(c, name)
    const working = await withArtifactMutationLocksAsync(directory, [target.name], async () =>
      existsSync(target.docPath) ? readFileSync(target.docPath) : undefined)
    if (!working) return notSource(c, name)
    const requested = c.req.query("version")
    let bytes: Uint8Array = working
    let version: { id: string | null; number: number | null; sha256: string }
    if (requested) {
      const found = readVersion(directory, requested)
      if (!found || found.document.removed_at != null || found.document.id !== readDocumentVersions(directory, target.name).document?.id)
        return c.json({ error: "That version is not in this matter." }, 404)
      if (found.version.removed)
        return c.json({ error: `Version ${found.version.number} was removed by a redaction; its text is no longer kept.`, code: "removed_by_redaction" }, 410)
      const kept = readVersionBytes(directory, found.version)
      if (!kept) return c.json({ error: "The kept bytes of this version are missing or damaged.", code: "version_damaged" }, 409)
      bytes = kept
      version = { id: found.version.id, number: found.version.number, sha256: found.version.sha256 }
    } else {
      const sha256 = sha256Hex(working)
      const kept = findVersionBySha(directory, target.name, sha256)
      version = { id: kept?.id ?? null, number: kept?.number ?? null, sha256 }
    }
    const markup = await readMarkup(target.name, bytes)
    if (!markup) return unreadable(c, target.name)
    return c.json({ name: target.name, version, current: version.sha256 === sha256Hex(working), ...markup })
  })

  // Applies one decision per listed change and adds the approved replies, as
  // one new version. `base_sha256` is the version the list was read from; a
  // listed change left out is held. `replies` answer open threads by their
  // root's id, one per thread. Decisions go first, so a reply sits on the text
  // as decided; any refusal writes nothing.
  app.post("/:id/documents/markup/decisions", async (c) => {
    const directory = options.directory(c.req.param("id"))
    const name = wordName(c.req.query("name"))
    if (!name) return c.json({ error: "Tracked changes are decided in Word (.docx) documents only." }, 400)
    const body = decisionsBody(await c.req.json().catch(() => undefined))
    if (!body)
      return c.json({ error: "Provide base_sha256, one accept, reject or hold decision for each listed change, each change once, and at most one reply of up to 4,000 characters per thread." }, 400)
    const decided = [...body.decisions.values()].some((decision) => decision !== "hold")
    if (!decided && !body.replies.length)
      return c.json({ error: "Accept or reject at least one change, or approve a reply; held changes stay as they are.", code: "nothing_decided" }, 422)
    const actor = oneLine(c.req.header("x-dochaus-actor")) || oneLine(readDraftingPreferences().attorney) || undefined
    if (body.replies.length && !actor)
      return c.json({ error: "Replies are signed with your name. Add your name, then apply again.", code: "author_required" }, 422)
    const initial = sourceTarget(directory, name)
    if (!initial) return notSource(c, name)
    const result = await withArtifactMutationLocksAsync(directory, options.locks(directory, initial.docPath), async () => {
      const target = sourceTarget(directory, name)
      if (!target || !existsSync(target.docPath)) return notSource(c, name)
      const bytes = readFileSync(target.docPath)
      if (sha256Hex(bytes) !== body.base)
        return c.json({ error: `${target.name} changed after its changes were listed. Refresh the list and decide again.`, code: "base_changed" }, 409)
      const db = openDb(directory)
      const pending = (() => { try { return listPendingRedlines(db, target.docPath).length } finally { db.close() } })()
      if (pending)
        return c.json({ error: `Accept or reject doc.haus's proposed changes to ${target.name} before deciding the tracked changes.`, code: "proposals_pending" }, 409)
      const markup = await readMarkup(target.name, bytes)
      if (!markup) return unreadable(c, target.name)
      const closed = closedThread(markup, body.replies)
      if (closed) return c.json(closed, 409)
      // Replies alone leave every change as it is.
      let applied: Awaited<ReturnType<typeof applyMarkupDecisions>> = {
        bytes, outcomes: markup.revisions.map(({ id }) => ({ id, decision: "hold", outcome: "held" })) }
      if (decided) try {
        applied = await applyMarkupDecisions(bytes, body.decisions)
      } catch (error) {
        if (error instanceof MarkupDecisionError)
          return c.json({ error: error.message, code: error.code, id: error.revision, ...(error.mergedWith.length ? { merged_with: error.mergedWith } : {}) }, 409)
        console.warn("[markup] decisions refused", diagnosticErrorMetadata(error))
        return c.json({ error: "The decisions could not be applied. Nothing was changed; refresh the list and try again.", code: "decision_failed" }, 409)
      }
      let replied: Awaited<ReturnType<typeof addCommentReplies>> = { bytes: applied.bytes, added: [] }
      if (body.replies.length) try {
        replied = await addCommentReplies(applied.bytes, body.replies, { author: actor!, initials: replyInitials(actor!), date: new Date().toISOString() })
      } catch (error) {
        if (error instanceof CommentReplyError)
          return c.json({ error: error.message, code: error.code, thread: error.thread }, error.code === "empty_reply" || error.code === "unattributed_reply" ? 422 : 409)
        console.warn("[markup] replies refused", diagnosticErrorMetadata(error))
        return c.json({ error: "The replies could not be added. Nothing was changed; refresh the list and try again.", code: "reply_failed" }, 409)
      }
      // Who made the change and how it was decided; never its text. Replies by id only.
      const rows = applied.outcomes.map(({ id, decision, outcome }) => {
        const { kind, family, author, date } = markup.revisions.find((revision) => revision.id === id)!
        return { id, kind, family, author, date, decision, outcome }
      })
      const count = (decision: MarkupDecision, outcome = "applied") => rows.filter((row) => row.decision === decision && row.outcome === outcome).length
      const counts = { accepted: count("accept"), rejected: count("reject"), held: count("hold", "held"), consumed: rows.filter((row) => row.outcome === "consumed").length }
      const added = replied.added
      const answered = added.length === 1 ? "1 reply added" : `${added.length} replies added`
      const version: VersionWrite = decided
        ? { origin: "accept", actor, event: { kind: "accept", actor, detail: { markup: rows, ...(added.length ? { replies: added } : {}) } },
          label: `Tracked changes resolved: ${counts.accepted} accepted, ${counts.rejected} rejected, ${counts.held} held` +
            (counts.consumed ? `, ${counts.consumed} resolved with another change` : "") + (added.length ? `; ${answered}` : "") }
        : { origin: "revision", actor, label: answered }
      try {
        return { written: await options.write(directory, target.name, Buffer.from(replied.bytes), version), counts, rows, added }
      } catch (error) {
        if (error instanceof SourceDocumentConflictError) return c.json({ error: error.message }, 409)
        throw error
      }
    })
    if (result instanceof Response) return result
    return c.json({ ...result.written, counts: result.counts, decisions: result.rows, replies: result.added })
  })

  return app
}

function wordName(requested: string | undefined) {
  const name = matterDocumentName(requested ?? "")
  return name && path.extname(name).toLowerCase() === ".docx" ? name : undefined
}

/** The registered uploaded source `name`, read from legal.db; undefined for a
 * deliverable, an unregistered file or a link. */
function sourceTarget(directory: string, name: string) {
  const target = readDocumentTarget(directory, name)
  if (!target || !existsSync(path.join(directory, ".dochaus", "legal.db"))) return
  const db = openDb(directory)
  try {
    return getDocumentRole(db, target.docPath) === "source" ? target : undefined
  } finally {
    db.close()
  }
}

function notSource(c: Context, name: string) {
  return c.json({ error: `${name} is not an uploaded Word source in this matter.`, code: "not_source" }, 404)
}

async function readMarkup(name: string, bytes: Uint8Array): Promise<DocxMarkup | undefined> {
  try {
    return await readDocxMarkup(bytes)
  } catch (error) {
    console.warn("[markup] tracked changes and comments not read", { name, ...diagnosticErrorMetadata(error) })
  }
}

function unreadable(c: Context, name: string) {
  return c.json({ error: `The tracked changes and comments in ${name} could not be read. Open it in Word, save it again and upload it, or resolve them in Word.`, code: "markup_unreadable" }, 422)
}

function decisionsBody(input: unknown) {
  if (!input || typeof input !== "object" || Array.isArray(input)) return
  const { base_sha256: base, decisions, replies = [] } = input as Record<string, unknown>
  if (typeof base !== "string" || !/^[a-f0-9]{64}$/.test(base) || !Array.isArray(decisions) || !Array.isArray(replies)) return
  const map = new Map<string, MarkupDecision>()
  for (const row of decisions) {
    const { id, decision } = (row && typeof row === "object" ? row : {}) as Record<string, unknown>
    if (typeof id !== "string" || !id || map.has(id) || typeof decision !== "string" || !DECISIONS.has(decision)) return
    map.set(id, decision as MarkupDecision)
  }
  const answers: CommentReply[] = []
  for (const row of replies) {
    const { thread, text } = (row && typeof row === "object" ? row : {}) as Record<string, unknown>
    if (typeof thread !== "number" || !Number.isSafeInteger(thread) || thread < 0 || answers.some((reply) => reply.thread === thread)
      || typeof text !== "string" || text.length > REPLY_LIMIT) return
    answers.push({ thread, text })
  }
  return { base, decisions: map, replies: answers }
}

/** The first reply this listing cannot take: replies answer a thread's root, and
 * resolved threads stay as the other side left them. */
function closedThread(markup: DocxMarkup, replies: readonly CommentReply[]) {
  for (const { thread } of replies) {
    const root = markup.comments.find((comment) => comment.id === thread)
    if (!root) return { error: "This comment is no longer in the document. Refresh the list and reply again.", code: "thread_not_found", thread }
    if (root.parent !== null) return { error: "Reply to the comment that starts this thread.", code: "thread_not_root", thread }
    if (root.resolved) return { error: "This thread is resolved, so it takes no reply here. Reopen it in Word to answer it.", code: "thread_resolved", thread }
  }
}
