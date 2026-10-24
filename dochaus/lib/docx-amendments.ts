import { createHash } from "node:crypto"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { z } from "zod"
import { paragraphFields, refInstruction, runText, unlink } from "./docx-fields"
import { docxodus } from "./docxodus"
import { shownText } from "./list-label"

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

// A working-document projection is a view, not legal source authority. No native
// UNIDs or offsets leave this module. Handles are valid only for this exact base.
export async function createDocxRevision(bytes: Uint8Array) {
  const opened = await openRevision(new Uint8Array(bytes))
  try { return opened.revision } finally { opened.session.close() }
}
export type DocxRevisionPacket = Awaited<ReturnType<typeof createDocxRevision>>

export async function createDocxLocalization(bytes: Uint8Array, amendments: DocxAmendment[], limits: Limits = {}) {
  const opened = await openRevision(new Uint8Array(bytes))
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
export async function applyDocxLocalization(bytes: Uint8Array, packet: DocxLocalizationPacket, selections: unknown) {
  return applyDocxParagraphEdits(bytes, [], { packet, selections })
}

// Paragraph-addressed amendments for this exact base: each names one projection
// paragraph by its handle and the occurrence of its exact quote inside that
// paragraph's text. Only that paragraph is searched, so a localization packet's
// whole-document candidate limits do not apply. `localized` carries any
// handle-selected amendments for the same base under applyDocxLocalization's
// rules. Admission is all-or-nothing across both; this never writes files.
export async function applyDocxParagraphEdits(bytes: Uint8Array, edits: DocxParagraphEdit[],
  localized?: { packet: DocxLocalizationPacket; selections: unknown }) {
  const source = new Uint8Array(bytes)
  if (localized && digest(source) !== localized.packet.revision.source_sha256) throw new Error("stale_docx_revision")
  const selected = localized ? selectionSchema.parse(localized.selections) : []
  const addressed = z.array(paragraphEditSchema).parse(edits)
  const opened = await openRevision(source)
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
  if (/[\r\n\t]/.test(amendment.replacement)) throw new Error("unsupported_structural_replacement")
  if (!amendment.replacement && match.text === opened.paragraphs[ordinal].match.text)
    throw new Error("unsupported_empty_paragraph_replacement")
  const change = changedDocxSpan(match, amendment.replacement)
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

async function writeEdits(opened: Opened, edits: ReturnType<typeof admitted>[]) {
  if (edits.some((a, i) => edits.slice(i + 1).some(b => a.ordinal === b.ordinal
    && a.match.span.start < b.match.span.start + b.match.span.length
    && b.match.span.start < a.match.span.start + a.match.span.length))) throw new Error("overlapping_docx_amendments")
  const expected = opened.paragraphs.map(p => p.match.text)
  const ordered = edits.toSorted((a, b) => a.ordinal - b.ordinal || b.match.span.start - a.match.span.start)
  // Edits touching a live cross-reference are written to the source with that
  // field unwrapped: the same paragraphs and text, so the same spans.
  const unlinked = unlinkedSource(opened, edits)
  const target = unlinked ? await openRevision(unlinked) : opened
  let output: Uint8Array
  try {
    for (const edit of ordered) {
      const text = expected[edit.ordinal]
      expected[edit.ordinal] = text.slice(0, edit.match.span.start) + edit.amendment.replacement
        + text.slice(edit.match.span.start + edit.match.span.length)
      // The complete authenticated quote localizes the occurrence. Unchanged
      // context keeps its own runs; only the changed text is mutated.
      const paragraph = target.paragraphs[edit.ordinal]!.match
      const change = target === opened ? edit.change : changedDocxSpan(paragraphSpan(paragraph,
        edit.match.span.start - opened.paragraphs[edit.ordinal]!.match.span.start, edit.match.text.length), edit.amendment.replacement)
      if (!change.match.text && !change.replacement) continue
      const result = target.session.replaceMatch(change.match, change.replacement)
      if (!result.success) throw new Error(`docx_mutation_failed:${result.error?.code}`)
    }
    output = target.session.save()
  } finally { if (target !== opened) target.session.close() }
  const reopened = await openRevision(output)
  try {
    if (JSON.stringify(reopened.paragraphs.map(p => p.match.text)) !== JSON.stringify(expected))
      throw new Error("docx_saved_text_mismatch")
    const changed = new Set(edits.map(edit => edit.ordinal))
    verifyPreservation(opened, reopened, changed)
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
        limitation: "changed text inherits first touched run formatting; structural edits and existing tracked revisions unsupported",
      },
    }
  } finally { reopened.session.close() }
}

export function changedDocxSpan(match: Match, replacement: string) {
  const { offset, replaced, changed } = changedRange(match.text, replacement)
  const formatting = match.fragments.find((fragment, index) => match.fragments.slice(0, index + 1)
    .reduce((length, row) => length + row.text.length, 0) > offset)?.formatting ?? match.fragments.at(-1)?.formatting
  return { match: { ...match, text: replaced, span: { start: match.span.start + offset, length: replaced.length } },
    replacement: changed, formatting }
}

/** The part of `text` a replacement by `replacement` rewrites: from `offset`
 * (UTF-16 units), `replaced` becomes `changed`. */
export function changedRange(text: string, replacement: string) {
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

export async function openRevision(source: Uint8Array) {
  const parts = unzipSync(source)
  const documents = Object.fromEntries(Object.entries(parts).filter(([name]) => name.endsWith(".xml"))
    .map(([name, bytes]) => [name, parse(strFromU8(bytes))]))
  if (!documents["word/document.xml"]) throw new Error("invalid_docx_package")
  if (Object.values(documents).some(doc => ["ins", "del", "moveFrom", "moveTo"].some(name => elements(doc, name).length)))
    throw new DocxTrackedRevisionsError()
  const dx = await docxodus()
  const session = dx.openDocxSession(source, { smartQuotes: false, persistAnchorIds: false })
  try {
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
      return { match, part, node, xml, metadata }
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
    return { session, parts, documents, paragraphs, revision }
  } catch (error) { session.close(); throw error }
}

export function verifyPreservation(before: Awaited<ReturnType<typeof openRevision>>, after: Awaited<ReturnType<typeof openRevision>>, changed: Set<number>) {
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
// grep projects w:t text only: breaks of any type, tabs and field codes add
// nothing, so a paragraph holding only those has no match and is skipped here.
function emptyParagraph(node: Element) {
  return elements(node, "t").every(text => !text.textContent)
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
