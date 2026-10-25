import { createHash } from "node:crypto"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { z } from "zod"
import { REVISION_ELEMENTS } from "./docx-comments"
import { paragraphFields, refInstruction, runText, unlink } from "./docx-fields"
import { docxodus } from "./docxodus"
import { shownText } from "./list-label"
import { normalizeUntrustedText } from "./untrusted-text"

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const PT = "http://powertools.codeplex.com/2011"
const schema = "docx-amendments/1" as const
const amendmentSchema = z.object({ id: z.string().min(1), before_quote: z.string().min(1), replacement: z.string() }).strict()
const selectionSchema = z.array(z.union([
  z.object({ amendment_id: z.string(), target_handle: z.string().min(1) }).strict(),
  z.object({ amendment_id: z.string(), unresolved_reason: z.string().min(1) }).strict(),
]))
const paragraphEditSchema = z.object({ id: z.string().min(1), paragraph_handle: z.string().min(1),
  occurrence: z.number().int().min(0), before_quote: z.string().min(1), replacement: z.string() }).strict()
export type DocxAmendment = z.infer<typeof amendmentSchema>
export type DocxLocalizationSelection = z.infer<typeof selectionSchema>[number]
export type DocxParagraphEdit = z.infer<typeof paragraphEditSchema>
type Session = ReturnType<Awaited<ReturnType<typeof docxodus>>["openDocxSession"]>
type Match = ReturnType<Session["grep"]>[number]
type Opened = Awaited<ReturnType<typeof openRevision>>
type Binding = { handle: string; amendment: DocxAmendment; match: Match; ordinal: number }
type Limits = { maxCandidates?: number; maxContextCharacters?: number }
export type DocxLocalizationLimits = Limits
// How the reader treats pending tracked revisions. "refuse" (the default) throws
// DocxTrackedRevisionsError. "layer" projects their accepted text and leaves them
// pending: an edit in a paragraph with a pending revision is refused, and other
// edits are written beside them. With "layer", a revisionAuthor writes edits as
// native tracked changes under that name, and pending revisions by ownAuthor
// (doc.haus's own name: edits an earlier Accept kept tracked) hold no paragraph:
// an edit beside them is written, one inside their text is refused.
export type DocxRevisionOptions = { pendingRevisions?: "refuse" | "layer"; revisionAuthor?: string; ownAuthor?: string }

// A working-document projection is a view, not legal source authority. No native
// UNIDs or offsets leave this module. Handles are valid only for this exact base.
export async function createDocxRevision(bytes: Uint8Array, options: DocxRevisionOptions = {}) {
  const opened = await openRevision(new Uint8Array(bytes), options)
  try { return opened.revision } finally { opened.session.close() }
}
export type DocxRevisionPacket = Awaited<ReturnType<typeof createDocxRevision>>

export async function createDocxLocalization(bytes: Uint8Array, amendments: DocxAmendment[], limits: Limits = {}, options: DocxRevisionOptions = {}) {
  const opened = await openRevision(new Uint8Array(bytes), options)
  try { return localize(opened, amendments, limits).packet } finally { opened.session.close() }
}
export type DocxLocalizationPacket = Awaited<ReturnType<typeof createDocxLocalization>>

// Ephemeral code-only occurrence proof. Rebuild the complete packet before
// exposing spans; native offsets never become model or persisted authority.
export async function docxLocalizationCandidateOccurrences(bytes: Uint8Array, packet: DocxLocalizationPacket) {
  const opened = await openRevision(new Uint8Array(bytes))
  try {
    const rebuilt = localize(opened, packet.amendments, packet.limits)
    if (JSON.stringify(rebuilt.packet) !== JSON.stringify(packet)) throw new Error("docx_packet_binding_changed")
    return [...rebuilt.bindings.values()].map(binding => ({ amendment_id: binding.amendment.id,
      target_handle: binding.handle, paragraph_handle: opened.revision.paragraphs[binding.ordinal]!.paragraph_handle,
      start: binding.match.span.start, end: binding.match.span.start + binding.match.span.length }))
  } finally { opened.session.close() }
}

// Call under the caller's existing document mutation lock, with freshly read
// canonical bytes. Admission is all-or-nothing; this function never writes files.
// The locator can select a handle or refuse. It cannot supply replacement text.
export async function applyDocxLocalization(bytes: Uint8Array, packet: DocxLocalizationPacket, selections: unknown, options: DocxRevisionOptions = {}) {
  return applyDocxParagraphEdits(bytes, [], { packet, selections }, options)
}

// Paragraph-addressed amendments for this exact base: each names one projection
// paragraph by its handle and the occurrence of its exact quote inside that
// paragraph's text. Only that paragraph is searched, so a localization packet's
// whole-document candidate limits do not apply. `localized` carries any
// handle-selected amendments for the same base under applyDocxLocalization's
// rules. Admission is all-or-nothing across both; this never writes files.
export async function applyDocxParagraphEdits(bytes: Uint8Array, edits: DocxParagraphEdit[],
  localized?: { packet: DocxLocalizationPacket; selections: unknown }, options: DocxRevisionOptions = {}) {
  const source = new Uint8Array(bytes)
  if (localized && digest(source) !== localized.packet.revision.source_sha256) throw new Error("stale_docx_revision")
  const selected = localized ? selectionSchema.parse(localized.selections) : []
  const addressed = z.array(paragraphEditSchema).parse(edits)
  const opened = await openRevision(source, options)
  try {
    const bindings = [
      ...localized ? selectedBindings(opened, localized.packet, selected) : [],
      ...addressed.map(edit => paragraphBinding(opened, edit)),
    ]
    if (!bindings.length) throw new Error("missing_docx_amendments")
    if (new Set(bindings.map(binding => binding.amendment.id)).size !== bindings.length) throw new Error("duplicate_amendment_id")
    return await writeEdits(opened, bindings.map(binding => admitted(opened, binding)))
  } finally { opened.session.close() }
}

function selectedBindings(opened: Opened, packet: DocxLocalizationPacket, selected: DocxLocalizationSelection[]) {
  const rebuilt = localize(opened, packet.amendments, packet.limits)
  if (JSON.stringify(rebuilt.packet) !== JSON.stringify(packet)) throw new Error("docx_packet_binding_changed")
  if (selected.length !== packet.amendments.length || new Set(selected.map(s => s.amendment_id)).size !== selected.length)
    throw new Error("missing_or_duplicate_amendment_selection")
  return selected.map(selection => {
    const amendment = packet.amendments.find(a => a.id === selection.amendment_id)
    if (!amendment) throw new Error("unknown_amendment_selection")
    if ("unresolved_reason" in selection) throw new Error(`unresolved_docx_localization:${selection.amendment_id}`)
    const binding = rebuilt.bindings.get(`${selection.amendment_id}:${selection.target_handle}`)
    if (!binding || binding.amendment.id !== selection.amendment_id) throw new Error("unknown_or_conflicting_target_handle")
    if (binding.match.text !== amendment.before_quote) throw new Error("before_quote_changed")
    return binding
  })
}

// The occurrence-th match of the quote in the named paragraph, counted left to
// right without overlap as grep counts candidates; its handle is the one a
// localization packet would offer for the same span.
function paragraphBinding(opened: Opened, edit: DocxParagraphEdit): Binding {
  if (/[\r\n\t]/.test(edit.before_quote)) throw new Error("unsupported_cross_block_quote")
  const ordinal = opened.revision.paragraphs.findIndex(paragraph => paragraph.paragraph_handle === edit.paragraph_handle)
  if (ordinal < 0) throw new Error("unknown_paragraph_handle")
  const paragraph = opened.paragraphs[ordinal]!.match
  let start = -edit.before_quote.length
  for (let found = 0; found <= edit.occurrence; found++) {
    start = paragraph.text.indexOf(edit.before_quote, start + edit.before_quote.length)
    if (start < 0) throw new Error(`before_quote_not_found:${edit.id}`)
  }
  const match = paragraphSpan(paragraph, start, edit.before_quote.length)
  return { handle: targetHandle(opened.revision.source_sha256, ordinal, match), ordinal, match,
    amendment: { id: edit.id, before_quote: edit.before_quote, replacement: edit.replacement } }
}

// The grep match for one span of a whole-paragraph match: the same anchor,
// paragraph-relative span and per-run fragments.
function paragraphSpan(paragraph: Match, start: number, length: number): Match {
  const end = start + length
  let at = 0
  const fragments = paragraph.fragments.flatMap(fragment => {
    const from = at
    at += fragment.text.length
    const first = Math.max(start, from)
    const last = Math.min(end, at)
    return first < last ? [{ ...fragment, text: fragment.text.slice(first - from, last - from),
      spanInElement: { start: fragment.spanInElement.start + first - from, length: last - first } }] : []
  })
  const text = paragraph.text.slice(start, end)
  return { ...paragraph, text, span: { start: paragraph.span.start + start, length }, fragments, groups: [text],
    contextBefore: paragraph.text.slice(Math.max(0, start - 160), start), contextAfter: paragraph.text.slice(end, end + 160) }
}

function admitted(opened: Opened, binding: Binding) {
  const { amendment, match, ordinal } = binding
  // The touch rule (see pendingRevisionParagraphs); only a layer-mode reader marks paragraphs.
  const pending = opened.paragraphs[ordinal]!.pending
  if (othersPending(pending, opened.options.ownAuthor)) throw new Error(`docx_pending_revision:${amendment.id}`)
  if (/[\r\n\t]/.test(amendment.replacement)) throw new Error("unsupported_structural_replacement")
  if (!amendment.replacement && match.text === opened.paragraphs[ordinal].match.text)
    throw new Error("unsupported_empty_paragraph_replacement")
  const change = changedDocxSpan(match, amendment.replacement, Boolean(opened.options.revisionAuthor))
  // Beside doc.haus's own pending change, never inside it: Docxodus would nest
  // the new revision in that one.
  if (pending && insideRevision(opened, paragraphSpan(match, change.match.span.start - match.span.start, change.match.text.length)))
    throw new Error(`docx_own_revision:${amendment.id}`)
  const xml = parse(opened.session.raw.getXml(match.enclosingAnchor.id))
  if (!elements(xml, "fldChar").length && !elements(xml, "instrText").length) return { ...binding, change, unlinks: [] as number[] }
  // A live cross-reference (a REF field) is the only field an edit may stand
  // beside. One the edit's changed text touches becomes the text it shows.
  const paragraph = opened.paragraphs[ordinal]!
  const fields = paragraphFields(paragraph.node)
  const owned = new Set(fields.flatMap(field => field.controls))
  if (runText(paragraph.node) !== paragraph.match.text || fields.some(field => !refInstruction(field))
    || [...elements(paragraph.node, "fldChar"), ...elements(paragraph.node, "instrText")].some(node => !owned.has(node)))
    throw new Error("unsupported_field_replacement")
  const start = change.match.span.start - paragraph.match.span.start
  const end = start + change.match.text.length
  return { ...binding, change, unlinks: fields.flatMap((field, index) => start < field.to && field.from < end ? [index] : []) }
}

// Whether a run of the changed span lies inside a tracked insertion, deletion or move.
function insideRevision(opened: Opened, match: Match) {
  const runs = new Map(elements(parse(opened.session.raw.getXml(match.enclosingAnchor.id)), "r").map(run => [run.getAttributeNS(PT, "Unid"), run]))
  return match.fragments.some(fragment => {
    for (let node = runs.get(fragment.unid)?.parentNode; node && node.nodeType === 1; node = node.parentNode)
      if ((node as Element).namespaceURI === W && ["ins", "del", "moveFrom", "moveTo"].includes((node as Element).localName)) return true
    return false
  })
}

/** Each selected amendment the writer would refuse for landing inside
 * doc.haus's own pending change (ownAuthor), by amendment id; nothing is
 * written. Prepare reports these as open items before qualifying the batch. */
export async function ownRevisionRefusals(bytes: Uint8Array, packet: DocxLocalizationPacket, selections: unknown, options: DocxRevisionOptions) {
  const opened = await openRevision(new Uint8Array(bytes), options)
  try {
    return new Set(selectedBindings(opened, packet, selectionSchema.parse(selections)).flatMap(binding => {
      try { admitted(opened, binding); return [] }
      catch (error) { return error instanceof Error && error.message === `docx_own_revision:${binding.amendment.id}` ? [binding.amendment.id] : [] }
    }))
  } finally { opened.session.close() }
}

// The source with the REF fields the edits touch unwrapped to their text, for
// the edits to be written to; undefined when they touch none.
function unlinkedSource(opened: Opened, edits: ReturnType<typeof admitted>[]) {
  const touched = Map.groupBy(edits.filter(edit => edit.unlinks.length), edit => edit.ordinal)
  if (!touched.size) return
  const documents = new Map<string, Document>()
  for (const [ordinal, rows] of touched) {
    const { part, node } = opened.paragraphs[ordinal]!
    const doc = documents.get(part) ?? parse(strFromU8(opened.parts[part]!))
    documents.set(part, doc)
    const fields = paragraphFields(elements(doc, "p")[elements(opened.documents[part]!, "p").indexOf(node)]!)
    for (const index of new Set(rows.flatMap(row => row.unlinks))) unlink(fields[index]!)
  }
  const parts = { ...opened.parts }
  for (const [part, doc] of documents) parts[part] = strToU8(new XMLSerializer().serializeToString(doc))
  return zipSync(parts, { mtime: new Date(1980, 0, 1) })
}

// The report's line for the live cross-references unlinkedSource made text:
// how many, and the paragraphs they are in, by list number or opening words.
function unlinkedFields(opened: Opened, edits: ReturnType<typeof admitted>[]) {
  const counts = [...Map.groupBy(edits, edit => edit.ordinal)].map(([ordinal, rows]) => [ordinal, new Set(rows.flatMap(row => row.unlinks)).size] as const)
    .filter(([, count]) => count)
  if (!counts.length) return []
  const total = counts.reduce((sum, [, count]) => sum + count, 0)
  const where = counts.map(([ordinal]) => {
    const view = opened.revision.paragraphs[ordinal]!
    return view.label ?? `"${view.text.length > 40 ? `${view.text.slice(0, 40).trimEnd()}…` : view.text}"`
  })
  return [`${total} cross-reference field${total === 1 ? "" : "s"} converted to text in paragraph${where.length === 1 ? "" : "s"} ${where.join(", ")}`]
}

async function writeEdits(opened: Opened, edits: ReturnType<typeof admitted>[]) {
  if (edits.some((a, i) => edits.slice(i + 1).some(b => a.ordinal === b.ordinal
    && a.match.span.start < b.match.span.start + b.match.span.length
    && b.match.span.start < a.match.span.start + a.match.span.length))) throw new Error("overlapping_docx_amendments")
  const expected = opened.paragraphs.map(p => p.match.text)
  const ordered = edits.toSorted((a, b) => a.ordinal - b.ordinal || b.match.span.start - a.match.span.start)
  // Edits touching a live cross-reference are written to the source with that
  // field unwrapped: the same paragraphs and text, so the same spans.
  const unlinked = unlinkedSource(opened, edits)
  const target = unlinked ? await openRevision(unlinked, opened.options) : opened
  try {
    for (const edit of ordered) {
      const text = expected[edit.ordinal]
      expected[edit.ordinal] = text.slice(0, edit.match.span.start) + edit.amendment.replacement
        + text.slice(edit.match.span.start + edit.match.span.length)
      // The complete authenticated quote localizes the occurrence. Unchanged
      // context keeps its own runs; only the changed text is mutated.
      const paragraph = target.paragraphs[edit.ordinal]!.match
      const change = target === opened ? edit.change : changedDocxSpan(paragraphSpan(paragraph,
        edit.match.span.start - opened.paragraphs[edit.ordinal]!.match.span.start, edit.match.text.length), edit.amendment.replacement,
        Boolean(opened.options.revisionAuthor))
      if (!change.match.text && !change.replacement) continue
      const result = target.session.replaceMatch(change.match, change.replacement)
      if (!result.success) throw new Error(`docx_mutation_failed:${result.error?.code}`)
    }
    const output = target.session.save()
    const reopened = await openRevision(output, opened.options)
    try {
      if (JSON.stringify(reopened.paragraphs.map(p => p.match.text)) !== JSON.stringify(expected))
        throw new Error("docx_saved_text_mismatch")
      const changed = new Set(edits.map(edit => edit.ordinal))
      // The edited runs are those of the session the edits were written in,
      // which stays open until its own check has compared their formatting.
      verifyPreservation(opened, reopened, changed)
      if (target !== opened) verifyPreservation(target, reopened, changed)
      return {
        bytes: output,
        // The projection these edits were admitted against.
        revision: opened.revision,
        report: {
          source_sha256: opened.revision.source_sha256,
          output_sha256: reopened.revision.source_sha256,
          amendments: edits.map(edit => ({ amendment_id: edit.amendment.id, target_handle: edit.handle,
            paragraph_handle: opened.revision.paragraphs[edit.ordinal].paragraph_handle,
            before_quote: edit.amendment.before_quote, replacement: edit.amendment.replacement,
            paragraph_before: opened.revision.paragraphs[edit.ordinal].text,
            paragraph_after: expected[edit.ordinal],
            replacement_formatting: edit.change.formatting,
          })),
          unchanged_paragraphs_verified: opened.paragraphs.length - changed.size,
          package_structure_verified: true,
          limitation: [`changed text inherits first touched run formatting; structural edits and ${opened.options.pendingRevisions === "layer"
            ? "edits in paragraphs with pending tracked revisions" : "existing tracked revisions"} unsupported`,
            ...opened.options.revisionAuthor ? [`changes tracked as ${opened.options.revisionAuthor}`] : [],
            ...unlinkedFields(opened, edits)].join("; "),
        },
      }
    } finally { reopened.session.close() }
  } finally { if (target !== opened) target.session.close() }
}

export function changedDocxSpan(match: Match, replacement: string, words = false) {
  const { offset, replaced, changed } = changedRange(match.text, replacement, words)
  const formatting = match.fragments.find((fragment, index) => match.fragments.slice(0, index + 1)
    .reduce((length, row) => length + row.text.length, 0) > offset)?.formatting ?? match.fragments.at(-1)?.formatting
  return { match: { ...match, text: replaced, span: { start: match.span.start + offset, length: replaced.length } },
    replacement: changed, formatting }
}

/** The part of `text` a replacement by `replacement` rewrites: from `offset`
 * (UTF-16 units), `replaced` becomes `changed`. With `words`, whole words. */
export function changedRange(text: string, replacement: string, words = false) {
  // Native offsets count UTF-16 units, but trim whole code points so a shared
  // surrogate never becomes an unchanged half of a changed character.
  const before = Array.from(text)
  const after = Array.from(replacement)
  let first = 0
  let lastBefore = before.length
  let lastAfter = after.length
  while (first < lastBefore && first < lastAfter && before[first] === after[first]) first++
  while (lastBefore > first && lastAfter > first && before[lastBefore - 1] === after[lastAfter - 1]) {
    lastBefore--
    lastAfter--
  }
  if (words && (first < lastBefore || first < lastAfter)) {
    // A tracked change reads as Word records overtyped words: whole words, and
    // a pure insertion takes the adjacent word as its anchor.
    const word = (char?: string) => char !== undefined && /[\p{L}\p{N}_]/u.test(char)
    if (first === lastBefore) if (lastBefore < before.length) { lastBefore++; lastAfter++ } else first--
    while (first > 0 && word(before[first - 1]) && (word(before[first]) || word(after[first]))) first--
    while (lastBefore < before.length && word(before[lastBefore]) && (word(before[lastBefore - 1]) || word(after[lastAfter - 1]))) {
      lastBefore++
      lastAfter++
    }
  }
  const removed = before.slice(first, lastBefore).join("")
  const inserted = after.slice(first, lastAfter).join("")
  // Native replacement needs a nonempty span. A pure insertion retains one
  // adjacent code point inside the authenticated quote as its formatting anchor.
  const anchor = !removed && inserted ? Math.min(first, before.length - 1) : first
  const offset = before.slice(0, anchor).join("").length
  const replaced = !removed && inserted ? before[anchor]! : removed
  const changed = !removed && inserted ? first < before.length ? inserted + replaced : replaced + inserted : inserted
  return { offset, replaced, changed }
}

function targetHandle(source_sha256: string, ordinal: number, match: Match) {
  return `target-${digest(JSON.stringify([source_sha256, ordinal, match.span.start, match.span.length, match.text])).slice(0, 24)}`
}

function localize(opened: Opened, proposed: DocxAmendment[], limits: Limits) {
  const amendments = z.array(amendmentSchema).min(1).parse(proposed)
  if (new Set(amendments.map(a => a.id)).size !== amendments.length) throw new Error("duplicate_amendment_id")
  const bounded = { maxCandidates: limits.maxCandidates ?? 512, maxContextCharacters: limits.maxContextCharacters ?? 200_000 }
  if (!Number.isSafeInteger(bounded.maxCandidates) || bounded.maxCandidates < 1
    || !Number.isSafeInteger(bounded.maxContextCharacters) || bounded.maxContextCharacters < 1) throw new Error("invalid_localization_limits")
  const bindings = new Map<string, Binding>()
  const views = opened.revision.paragraphs
  // Context stops at the paragraph's edge, so a whole-paragraph quote has none;
  // the adjacent paragraphs in the same part give its position.
  const neighbour = (ordinal: number, index: number) => views[index]?.scope === views[ordinal]!.scope ? views[index]!.text : null
  const candidates = amendments.flatMap(amendment => {
    if (/[\r\n\t]/.test(amendment.before_quote)) throw new Error("unsupported_cross_block_quote")
    const matches = opened.session.grep(literal(amendment.before_quote), { scope: 63, whitespace: 0, contextChars: 160 })
    if (!matches.length) throw new Error(`before_quote_not_found:${amendment.id}`)
    return matches.map(match => {
      const ordinal = opened.paragraphs.findIndex(p => p.match.enclosingAnchor.id === match.enclosingAnchor.id)
      if (ordinal < 0 || match.text !== amendment.before_quote) throw new Error("unowned_docx_match")
      const handle = targetHandle(opened.revision.source_sha256, ordinal, match)
      bindings.set(`${amendment.id}:${handle}`, { handle, amendment, match, ordinal })
      return { amendment_id: amendment.id, target_handle: handle, exact_text: match.text,
        context_before: match.contextBefore, context_after: match.contextAfter,
        previous_paragraph: neighbour(ordinal, ordinal - 1)?.slice(-160) ?? null,
        next_paragraph: neighbour(ordinal, ordinal + 1)?.slice(0, 160) ?? null,
        paragraph: views[ordinal],
        touched_formatting: match.fragments.map(fragment => ({ text: fragment.text, formatting: fragment.formatting })),
      }
    })
  })
  if (candidates.length > bounded.maxCandidates) throw new Error(`docx_candidate_limit_exceeded:${candidates.length}`)
  if (JSON.stringify(candidates).length > bounded.maxContextCharacters) throw new Error("docx_candidate_context_limit_exceeded")
  return { packet: { schema, revision: opened.revision, amendments, limits: bounded, candidates }, bindings }
}

// Word writes these elements when Track Changes is on. Such a revision has no
// single accepted text to project; the user resolves them in Word.
export class DocxTrackedRevisionsError extends Error {
  constructor() { super("unsupported_existing_tracked_revisions") }
}

// The touch rule. Its unit is one projection paragraph in any story, which is
// pending when the native revision listing names it (both ends of a move
// included), or when its own markup, paragraph mark included, or an enclosing
// table, row or cell's properties carry a revision element. Any author counts,
// but doc.haus's own (ownAuthor; see othersPending): an edit there would merge
// with their change in Word, or untracked become their text. doc.haus's own are
// edits an Accept wrote tracked into their markup; an edit may stand beside
// those, never inside them.

/** Whether a paragraph's pending revisions, by their authors as pendingAuthors
 * lists them ("" for none named), include one by someone other than `own`. */
export function othersPending(authors: readonly string[] | undefined, own?: string) {
  return Boolean(authors?.some(author => author !== own))
}

/** Each paragraph of these bytes with a pending revision, by paragraph_handle,
 * with the authors of its revisions in document order ("" for none named). */
export async function pendingRevisionParagraphs(bytes: Uint8Array) {
  const opened = await openRevision(new Uint8Array(bytes), { pendingRevisions: "layer" })
  try {
    return new Map(opened.paragraphs.flatMap((paragraph, ordinal) =>
      paragraph.pending ? [[opened.revision.paragraphs[ordinal]!.paragraph_handle, paragraph.pending] as const] : []))
  } finally { opened.session.close() }
}

/** The same paragraphs in a caller's own session, by the anchor id grep and
 * findAllByText name them, with the authors of their revisions. */
export function pendingRevisionAnchors(session: Session) {
  const listed = listedRevisions(session)
  // A table paragraph is read inside its outermost table, for the table, row
  // and cell properties.
  const inTables = new Map<string, Element>()
  for (const [token, anchor] of Object.entries(session.project().anchorIndex)) if (anchor.kind === "tbl")
    for (const node of elements(parse(session.raw.getXml(token)), "p")) if (!inTables.has(node.getAttributeNS(PT, "Unid") ?? ""))
      inTables.set(node.getAttributeNS(PT, "Unid") ?? "", node)
  const pending = new Map<string, string[]>()
  for (const { enclosingAnchor: { id, unid } } of session.grep("^[\\s\\S]*$", { scope: 63, whitespace: 0 })) {
    const node = inTables.get(unid) ?? parse(session.raw.getXml(id)).documentElement
    const authors = pendingAuthors(listed.get(id), node)
    if (authors) pending.set(id, authors)
  }
  return pending
}

/** The touch rule's refusal of an older chat-tool edit of `text` in these
 * paragraphs (anchor ids, as findAllByText names them), given their session's
 * pendingRevisionAnchors, in the words of Prepare's open item; undefined when
 * none has a pending revision. `verb` says what did not happen to the change. */
export function pendingRevisionRefusal(pending: Map<string, string[]>, anchors: string[], text: string, verb: "proposed" | "applied" | "made") {
  const touched = anchors.filter(id => pending.has(id))
  if (!touched.length) return
  const names = [...new Set(touched.flatMap(id => pending.get(id)!).map(author => normalizeUntrustedText(author).text.trim()).filter(Boolean))]
  const many = names.length > 1
  // As Prepare lists them: three names, then a count.
  const shown = names.length > 3 ? [...names.slice(0, 3), `${names.length - 3} other${names.length === 4 ? "" : "s"}`] : names
  const by = shown.length ? ` by ${shown.length > 1 ? `${shown.slice(0, -1).join(", ")} and ${shown.at(-1)}` : shown[0]}` : ""
  return `The change to ${JSON.stringify(text)} was not ${verb}: that paragraph has ${many ? "pending changes" : "a pending change"}${by}. Accept or reject ${many ? "them" : "it"} under Their changes, then request edits again.`
}

function listedRevisions(session: Session) {
  const authors = new Map<string, Set<string>>()
  for (const revision of session.listRevisions())
    for (const id of [revision.anchorId, ...revision.affectedAnchors.map(anchor => anchor.id)])
      if (id) authors.set(id, (authors.get(id) ?? new Set()).add(revision.author))
  return authors
}

function pendingAuthors(listed: Set<string> | undefined, paragraph: Element) {
  const properties = [] as Element[]
  for (let node = paragraph.parentNode; node && node.nodeType === 1; node = node.parentNode) {
    const element = node as Element
    if (element.namespaceURI === W && ["tbl", "tr", "tc"].includes(element.localName)) properties.push(...children(element, `${element.localName}Pr`))
  }
  const marked = [paragraph, ...properties].flatMap(node => REVISION_ELEMENTS.flatMap(name => elements(node, name)))
  if (!listed?.size && !marked.length) return
  return [...new Set([...listed ?? [], ...marked.map(node => node.getAttributeNS(W, "author") ?? "")])]
}

export async function openRevision(source: Uint8Array, options: DocxRevisionOptions = {}) {
  const parts = unzipSync(source)
  const documents = Object.fromEntries(Object.entries(parts).filter(([name]) => name.endsWith(".xml"))
    .map(([name, bytes]) => [name, parse(strFromU8(bytes))]))
  if (!documents["word/document.xml"]) throw new Error("invalid_docx_package")
  const layer = options.pendingRevisions === "layer"
  if (!layer && Object.values(documents).some(doc => ["ins", "del", "moveFrom", "moveTo"].some(name => elements(doc, name).length)))
    throw new DocxTrackedRevisionsError()
  const dx = await docxodus()
  // The projection is the accepted text in either mode; render_inline changes
  // only how this session's edits are written.
  const session = dx.openDocxSession(source, { smartQuotes: false, persistAnchorIds: false,
    ...options.revisionAuthor ? { trackedChanges: "render_inline" as const, revisionAuthor: options.revisionAuthor } : {} })
  try {
    const listed = layer ? listedRevisions(session) : undefined
    const projection = session.project()
    const tables = Object.entries(projection.anchorIndex).filter(([, anchor]) => anchor.kind === "tbl")
      .map(([token, anchor]) => ({ ...anchor, xml: session.raw.getXml(token) }))
    const counts = new Map<string, number>()
    const paragraphs = session.grep("^[\\s\\S]*$", { scope: 63, whitespace: 0 }).map(match => {
      const owner = projection.anchorIndex[match.enclosingAnchor.id]
        ?? tables.find(table => table.xml.includes(`"${match.enclosingAnchor.unid}"`))
      if (!owner) throw new Error("unowned_docx_paragraph")
      const part = owner.partUri.replace(/^\//, "")
      const sourceParagraphs = elements(documents[part], "p")
      let ordinal = counts.get(part) ?? 0
      // grep omits zero-length matches. Empty Word paragraphs remain in the
      // package preservation check, but have no literal amendment candidate.
      while (sourceParagraphs[ordinal] && emptyParagraph(sourceParagraphs[ordinal])) ordinal++
      counts.set(part, ordinal + 1)
      const node = sourceParagraphs[ordinal]
      const xml = parse(session.raw.getXml(match.enclosingAnchor.id)).documentElement
      if (!node || canonical(node) !== canonical(xml)) throw new Error("unsupported_docx_projection_mapping")
      const metadata = session.getBlockMetadata(match.enclosingAnchor.id)!
      // Kept off the revision views, so prompts and projection hashes do not change.
      const pending = listed && pendingAuthors(listed.get(match.enclosingAnchor.id), node)
      return { match, part, node, xml, metadata, pending }
    })
    if ([...counts].some(([part, count]) => elements(documents[part], "p").slice(count).some(node => !emptyParagraph(node))))
      throw new Error("incomplete_docx_projection")
    const source_sha256 = digest(source)
    const headingsByScope = new Map<string, { level: number; text: string }[]>()
    const views = paragraphs.map((paragraph, ordinal) => {
      const scope = paragraph.match.enclosingAnchor.scope
      const headingStack = headingsByScope.get(scope) ?? []
      headingsByScope.set(scope, headingStack)
      // The number Word's list numbering prints before the paragraph: shown in
      // the headings and the draft view, never part of `text`, which quotes match.
      const label = paragraph.metadata.list?.generatedLabel?.trim() || undefined
      if (paragraph.match.enclosingAnchor.kind === "h") {
        const level = (paragraph.metadata.outlineLevel ?? 0) + 1
        while (headingStack.length && headingStack.at(-1)!.level >= level) headingStack.pop()
        headingStack.push({ level, text: shownText({ text: paragraph.match.text, label }) })
      }
      const ancestors = [] as Element[]
      for (let node = paragraph.node.parentNode; node && node.nodeType === 1; node = node.parentNode) ancestors.unshift(node as Element)
      const table = ancestors.find(node => node.namespaceURI === W && node.localName === "tbl")
      const row = ancestors.findLast(node => node.namespaceURI === W && node.localName === "tr")
      const cell = ancestors.findLast(node => node.namespaceURI === W && node.localName === "tc")
      const tableContext = table && row && cell ? {
        table: elements(documents[paragraph.part], "tbl").indexOf(table) + 1,
        row: children(table, "tr").indexOf(row) + 1,
        column: children(row, "tc").indexOf(cell) + 1,
        row_context: children(row, "tc").map(node => elements(node, "t").map(t => t.textContent).join("")),
      } : null
      return { paragraph_handle: `paragraph-${digest(JSON.stringify([source_sha256, ordinal])).slice(0, 24)}`,
        scope,
        text: paragraph.match.text,
        ...(label ? { label } : {}),
        headings: headingStack.map(heading => ({ ...heading })),
        table: tableContext,
        style_id: paragraph.metadata.styleId ?? null,
        formatting: paragraph.match.fragments.map(fragment => ({ text: fragment.text, formatting: fragment.formatting })),
      }
    })
    // Header and footer text follows the body in this view. It is page
    // furniture, so each change of part opens with a label and no reader takes
    // it for body text. A body-only view has no label; escaped document text
    // cannot forge one.
    const markdown = views.map((view, ordinal) => {
      const part = view.scope !== (views[ordinal - 1]?.scope ?? "body") ? `[${partLabel(view.scope)}]\n\n` : ""
      const prefix = paragraphs[ordinal].match.enclosingAnchor.kind === "h" ? `${"#".repeat(view.headings.at(-1)!.level)} ` : ""
      const context = view.table ? `> Table ${view.table.table}, row ${view.table.row}, column ${view.table.column}\n\n` : ""
      const escape = (text: string) => text.replace(/[\\`*_{}[\]<>|]/g, "\\$&")
      return part + context + prefix + (view.label ? `${escape(view.label)} ` : "") + view.formatting.map(fragment => {
        const text = escape(fragment.text)
        return fragment.formatting.bold && fragment.formatting.italic ? `***${text}***`
          : fragment.formatting.bold ? `**${text}**` : fragment.formatting.italic ? `*${text}*` : text
      }).join("")
    }).join("\n\n")
    const revision = { schema, source_sha256, projection_sha256: digest(JSON.stringify(views)), markdown, paragraphs: views }
    return { session, parts, documents, paragraphs, revision, options }
  } catch (error) { session.close(); throw error }
}

export function verifyPreservation(before: Awaited<ReturnType<typeof openRevision>>, saved: Awaited<ReturnType<typeof openRevision>>, changed: Set<number>) {
  const after = before.options.revisionAuthor ? withoutTrackChangesSetting(before, saved) : saved
  if (before.paragraphs.length !== after.paragraphs.length) throw new Error("docx_paragraph_structure_changed")
  const ignoredBefore = new Set([...changed].map(ordinal => before.paragraphs[ordinal].node))
  const ignoredAfter = new Set([...changed].map(ordinal => after.paragraphs[ordinal].node))
  if (JSON.stringify(Object.keys(before.parts).sort()) !== JSON.stringify(Object.keys(after.parts).sort()))
    throw new Error("docx_package_parts_changed")
  for (const name of Object.keys(before.parts)) {
    if (before.documents[name]) {
      if (canonical(before.documents[name], ignoredBefore) !== canonical(after.documents[name], ignoredAfter))
        throw new Error(`docx_unrelated_structure_changed:${name}`)
      continue
    }
    if (digest(before.parts[name]) !== digest(after.parts[name])) throw new Error(`docx_unrelated_part_changed:${name}`)
  }
  for (const ordinal of changed) {
    const a = before.paragraphs[ordinal]
    const b = after.paragraphs[ordinal]
    if (canonical(a.node.cloneNode(false)) !== canonical(b.node.cloneNode(false))
      || canonical(children(a.node, "pPr")[0]) !== canonical(children(b.node, "pPr")[0]))
      throw new Error("docx_paragraph_properties_changed")
    // Native UNIDs are transient, but within this session identify original runs.
    const edited = parse(before.session.raw.getXml(a.match.enclosingAnchor.id))
    const remaining = new Map(elements(edited, "r").map(run => [run.getAttributeNS(PT, "Unid"), run]))
    for (const run of elements(a.xml, "r")) {
      const current = remaining.get(run.getAttributeNS(PT, "Unid"))
      if (current && canonical(children(run, "rPr")[0]) !== canonical(children(current, "rPr")[0]))
        throw new Error("docx_existing_run_properties_changed")
    }
  }
}

// A tracked-change save switches Track Changes on: Docxodus adds
// w:trackRevisions to the settings part and, when the package has none, creates
// the part with its relationship and content type. A part that differs only by
// that is compared as it was; any other change still refuses.
function withoutTrackChangesSetting(before: Opened, after: Opened): Opened {
  const settings = "word/settings.xml"
  const created = !before.parts[settings]
  const parts = { ...after.parts }
  const documents = { ...after.documents }
  const tolerated: [string, (node: Element) => boolean][] = [[settings, node => node.namespaceURI === W && node.localName === "trackRevisions"]]
  if (created) tolerated.push(
    ["[Content_Types].xml", node => node.localName === "Override" && node.getAttribute("PartName") === `/${settings}`],
    ["word/_rels/document.xml.rels", node => node.localName === "Relationship"
      && node.getAttribute("Type") === "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"])
  const stripped = (bytes: Uint8Array, added: (node: Element) => boolean) => {
    const root = parse(strFromU8(bytes)).documentElement
    for (const node of Array.from(root.childNodes)) if (node.nodeType === 1 && added(node as Element)) root.removeChild(node)
    return root
  }
  for (const [name, added] of tolerated) {
    if (!after.parts[name]) continue
    const root = stripped(after.parts[name], added)
    if (name === settings && created) {
      if (Array.from(root.childNodes).some(node => node.nodeType === 1)) continue
      delete parts[name]
      delete documents[name]
    } else if (before.parts[name] && canonical(root) === canonical(stripped(before.parts[name], added))) {
      parts[name] = before.parts[name]
      if (documents[name]) documents[name] = before.documents[name]!
    }
  }
  return { ...after, parts, documents }
}

function parse(xml: string) {
  const fail = (message: string) => { throw new Error(`invalid_docx_xml:${message}`) }
  return new DOMParser({ errorHandler: { warning: fail, error: fail, fatalError: fail } }).parseFromString(xml, "application/xml")
}
function partLabel(scope: string) {
  if (/^hdr\d*$/.test(scope)) return "Page header: page furniture, not body text"
  if (/^ftr\d*$/.test(scope)) return "Page footer: page furniture, not body text"
  return ({ body: "Body text", fn: "Footnotes", en: "Endnotes", cmt: "Comments" } as Record<string, string>)[scope] ?? scope
}
function elements(node: Document | Element, name: string) { return Array.from(node.getElementsByTagNameNS(W, name)) }
// grep projects accepted w:t text only: breaks of any type, tabs, field codes
// and text a pending deletion or move removes add nothing, so a paragraph holding
// only those has no match and is skipped here. Only a layer-mode reader meets
// such text; refuse mode has thrown before.
function emptyParagraph(node: Element) {
  return elements(node, "t").every(text => !text.textContent || removedByRevision(text, node))
}
function removedByRevision(text: Element, paragraph: Element) {
  for (let node = text.parentNode; node && node !== paragraph; node = node.parentNode)
    if ((node as Element).namespaceURI === W && ["del", "moveFrom"].includes((node as Element).localName)) return true
  return false
}
function children(node: Element, name: string) {
  return Array.from(node.childNodes).filter((child): child is Element => child.nodeType === 1)
    .filter(child => child.namespaceURI === W && child.localName === name)
}
function canonical(node: Node | undefined, ignored = new Set<Node>()): string {
  if (!node) return "null"
  if (ignored.has(node)) return '"amended-paragraph"'
  if (node.nodeType === 3 || node.nodeType === 4) return !node.nodeValue?.trim()
    && !(node.parentNode as Element)?.localName?.match(/^(t|delText|instrText)$/) ? "" : JSON.stringify(node.nodeValue)
  if (node.nodeType !== 1 && node.nodeType !== 9) return ""
  const element = node as Element
  const attributes = node.nodeType === 1 ? Array.from(element.attributes)
    .filter(attribute => attribute.namespaceURI !== "http://www.w3.org/2000/xmlns/" && attribute.namespaceURI !== PT)
    .map(attribute => [attribute.namespaceURI ?? "", attribute.localName ?? attribute.name, attribute.value]).sort() : []
  return JSON.stringify([element.namespaceURI ?? "", element.localName ?? "document", attributes,
    Array.from(node.childNodes).map(child => canonical(child, ignored)).filter(Boolean)])
}
function literal(text: string) { return text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") }
function digest(value: Uint8Array | string) { return createHash("sha256").update(value).digest("hex") }
