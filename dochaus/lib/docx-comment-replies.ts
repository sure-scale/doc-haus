import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import type { CommentListEntry } from "docxodus"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { docxodus } from "./docxodus"
import { readMarkupComments } from "./docx-markup"
import { normalizeUntrustedText } from "./untrusted-text"

// The lawyer's replies to a counterparty's comment threads (#200), written as
// native Word replies into one version's bytes in one Docxodus session. A
// thread is named by its root's native w:id, as the reader lists it
// (dochaus/lib/docx-markup.ts).
//
// - A reply answers the thread's root: Word links every reply to its root, and
//   a reply linked to another reply is shown loose in LibreOffice. Naming a
//   reply refuses.
// - Text is plain: Docxodus reads Markdown, so its punctuation is escaped and
//   comes back literally. A line break stays one; a blank line starts a new
//   paragraph. Text and author are written as a reader sees them (invisible
//   characters stripped) and trimmed; empty text refuses.
// - Every reply carries the given author, initials and date. An empty author
//   or a date that does not parse refuses; the date is written in UTC to the
//   second, as Word writes it.
// - Docxodus writes the threading records (commentsExtended and commentsIds,
//   created with their relationship and override when the parent was flat) and
//   places the reply's reference straight after its parent's, ahead of the
//   thread's earlier replies. LibreOffice lists a thread in reference order, so
//   the new reference moves after the thread's last one when both share a
//   paragraph. The reply has a reference and no range of its own.
// - The output is read back before it is returned: each reply must be listed
//   under its thread, open, with its author and text, and every tracked change
//   must be listed as before. Otherwise `reply_not_threaded` refuses and nothing
//   is returned. The session writes no anchor ids into the file.
export type CommentReply = { thread: number; text: string }
export type ReplyAttribution = { author: string; initials: string; date: string }
export type CommentReplyErrorCode = "thread_not_found" | "thread_not_root" | "empty_reply" | "unattributed_reply" | "reply_failed" | "reply_not_threaded"

export class CommentReplyError extends Error {
  constructor(readonly code: CommentReplyErrorCode, message: string, readonly thread?: number) {
    super(message)
  }
}

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
// Backslash-escapable Markdown punctuation.
const MARKDOWN = /[\\`*_{}[\]()#+\-.!<>~|]/g

/** Word's initials for a name: the first letter of each word, at most three, upper case. */
export function replyInitials(name: string) {
  return name.trim().split(/\s+/).map(word => /\p{L}/u.exec(word)?.[0] ?? "").join("").slice(0, 3).toUpperCase()
}

export async function addCommentReplies(bytes: Uint8Array, replies: readonly CommentReply[], by: ReplyAttribution) {
  const author = visible(by.author), time = Date.parse(by.date)
  if (!author || Number.isNaN(time)) throw new CommentReplyError("unattributed_reply", "A reply needs the lawyer's name and the time it was written.")
  const date = new Date(time).toISOString().replace(/\.\d{3}Z$/, "Z")
  const texts = replies.map(reply => {
    const text = visible(reply.text)
    if (!text) throw new CommentReplyError("empty_reply", "Write the reply before adding it.", reply.thread)
    return text
  })
  const dx = await docxodus()
  const revisions = revisionIds(dx, bytes)
  const session = dx.openDocxSession(bytes, { persistAnchorIds: false })
  const added: { thread: number; id: number }[] = []
  let saved: Uint8Array, listed: CommentListEntry[]
  try {
    listed = session.listComments()
    for (const [at, reply] of replies.entries()) {
      const root = listed.find(entry => entry.id === reply.thread)
      if (!root) throw new CommentReplyError("thread_not_found", "This comment is no longer in the document.", reply.thread)
      if (root.parentAnchorId !== undefined) throw new CommentReplyError("thread_not_root", "Reply to the comment that starts this thread.", reply.thread)
      const known = new Set(listed.map(entry => entry.id))
      if (!session.addCommentReply(root.anchorId, author, texts[at]!.replace(MARKDOWN, "\\$&"), { initials: by.initials, date }).success)
        throw new CommentReplyError("reply_failed", "This reply could not be added here. Nothing was changed.", reply.thread)
      listed = session.listComments()
      const created = listed.find(entry => !known.has(entry.id))
      if (!created) throw new CommentReplyError("reply_failed", "This reply could not be added here. Nothing was changed.", reply.thread)
      added.push({ thread: reply.thread, id: created.id })
    }
    saved = session.save()
  } finally {
    session.close()
  }
  const output = afterThreads(saved, added, listed)
  const comments = new Map((await readMarkupComments(output)).map(comment => [comment.id, comment]))
  for (const [at, reply] of added.entries()) {
    const read = comments.get(reply.id)
    if (!read || read.parent !== reply.thread || read.author !== author || read.resolved !== false || lines(read.text) !== lines(texts[at]!))
      throw new CommentReplyError("reply_not_threaded", "This reply could not be added to its thread. Nothing was changed.", reply.thread)
  }
  if (revisionIds(dx, output).join("\n") !== revisions.join("\n"))
    throw new CommentReplyError("reply_not_threaded", "Adding the replies would change the tracked changes. Nothing was changed.")
  return { bytes: output, added }
}

const visible = (text: string) => normalizeUntrustedText(text).text.trim()
// Breaks compared as the reader sees them: blank lines and the spaces around a break are not kept.
const lines = (text: string) => text.replace(/\s*\n\s*/g, "\n")

function revisionIds(dx: Awaited<ReturnType<typeof docxodus>>, bytes: Uint8Array) {
  const session = dx.openDocxSession(bytes, { persistAnchorIds: false, trackedChanges: "render_inline" })
  try {
    return session.listRevisions().map(row => row.id)
  } finally {
    session.close()
  }
}

// Moves each new reply's reference run to directly after its thread's last
// reference, in the part that holds it, when both share a paragraph. Replies
// are placed in the order added, each after the ones before it.
function afterThreads(bytes: Uint8Array, added: readonly { thread: number; id: number }[], listed: readonly CommentListEntry[]) {
  const ids = new Map(listed.map(entry => [entry.anchorId, entry.id]))
  const parents = new Map(listed.map(entry => [entry.id, entry.parentAnchorId === undefined ? undefined : ids.get(entry.parentAnchorId)]))
  const rootOf = (id: number) => {
    const seen = new Set<number>()
    let root = id
    while (parents.get(root) !== undefined && !seen.has(root)) {
      seen.add(root)
      root = parents.get(root)!
    }
    return root
  }
  const parts = unzipSync(bytes), changed = new Set<string>()
  const docs = Object.entries(parts).filter(([name, content]) => name.startsWith("word/") && name.endsWith(".xml") && strFromU8(content).includes("commentReference"))
    .map(([name, content]) => [name, new DOMParser().parseFromString(strFromU8(content), "application/xml")] as const)
  const waiting = new Set(added.map(reply => reply.id))
  for (const reply of added) {
    waiting.delete(reply.id)
    for (const [name, doc] of docs) {
      const references = Array.from(doc.getElementsByTagNameNS(W, "commentReference")), id = (node: Element) => Number(node.getAttributeNS(W, "id"))
      const own = references.find(node => id(node) === reply.id)
      if (!own) continue
      const last = references.filter(node => node !== own && !waiting.has(id(node)) && rootOf(id(node)) === reply.thread).at(-1)
      const run = own.parentNode as Element, paragraph = owningParagraph(run)
      if (last && paragraph && owningParagraph(last) === paragraph) {
        let after: Node = last
        while (after.parentNode !== paragraph) after = after.parentNode!
        if (after.nextSibling !== run) {
          paragraph.insertBefore(run, after.nextSibling)
          changed.add(name)
        }
      }
      break
    }
  }
  if (!changed.size) return bytes
  for (const [name, doc] of docs) if (changed.has(name)) parts[name] = strToU8(new XMLSerializer().serializeToString(doc))
  return zipSync(parts, { mtime: new Date(1980, 0, 1) })
}

function owningParagraph(node: Node) {
  for (let parent = node.parentNode; parent; parent = parent.parentNode)
    if (parent.nodeType === 1 && (parent as Element).namespaceURI === W && (parent as Element).localName === "p") return parent as Element
  return null
}
