import { createHash } from "node:crypto"
import { DOMParser } from "@xmldom/xmldom"
import { strFromU8, unzipSync } from "fflate"
import { z } from "zod"
import { docxodus } from "./docxodus"

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const PT = "http://powertools.codeplex.com/2011"
const schema = "docx-amendments/1" as const
const amendmentSchema = z.object({ id: z.string().min(1), before_quote: z.string().min(1), replacement: z.string() }).strict()
const selectionSchema = z.array(z.union([
  z.object({ amendment_id: z.string(), target_handle: z.string().min(1) }).strict(),
  z.object({ amendment_id: z.string(), unresolved_reason: z.string().min(1) }).strict(),
]))
export type DocxAmendment = z.infer<typeof amendmentSchema>
export type DocxLocalizationSelection = z.infer<typeof selectionSchema>[number]
type Session = ReturnType<Awaited<ReturnType<typeof docxodus>>["openDocxSession"]>
type Match = ReturnType<Session["grep"]>[number]
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
  const source = new Uint8Array(bytes)
  if (digest(source) !== packet.revision.source_sha256) throw new Error("stale_docx_revision")
  const selected = selectionSchema.parse(selections)
  const opened = await openRevision(source)
  try {
    const rebuilt = localize(opened, packet.amendments, packet.limits)
    if (JSON.stringify(rebuilt.packet) !== JSON.stringify(packet)) throw new Error("docx_packet_binding_changed")
    if (selected.length !== packet.amendments.length || new Set(selected.map(s => s.amendment_id)).size !== selected.length)
      throw new Error("missing_or_duplicate_amendment_selection")
    const edits = selected.map(selection => {
      const amendment = packet.amendments.find(a => a.id === selection.amendment_id)
      if (!amendment) throw new Error("unknown_amendment_selection")
      if ("unresolved_reason" in selection) throw new Error(`unresolved_docx_localization:${selection.amendment_id}`)
      const binding = rebuilt.bindings.get(`${selection.amendment_id}:${selection.target_handle}`)
      if (!binding || binding.amendment.id !== selection.amendment_id) throw new Error("unknown_or_conflicting_target_handle")
      if (binding.match.text !== amendment.before_quote) throw new Error("before_quote_changed")
      if (/[\r\n\t]/.test(amendment.replacement)) throw new Error("unsupported_structural_replacement")
      if (!amendment.replacement && binding.match.text === opened.paragraphs[binding.ordinal].match.text)
        throw new Error("unsupported_empty_paragraph_replacement")
      const xml = opened.session.raw.getXml(binding.match.enclosingAnchor.id)
      if (elements(parse(xml), "fldChar").length || elements(parse(xml), "instrText").length)
        throw new Error("unsupported_field_replacement")
      return { ...binding, change: changedDocxSpan(binding.match, amendment.replacement) }
    })
    if (edits.some((a, i) => edits.slice(i + 1).some(b => a.ordinal === b.ordinal
      && a.match.span.start < b.match.span.start + b.match.span.length
      && b.match.span.start < a.match.span.start + a.match.span.length))) throw new Error("overlapping_docx_amendments")
    const expected = opened.paragraphs.map(p => p.match.text)
    const ordered = edits.toSorted((a, b) => a.ordinal - b.ordinal || b.match.span.start - a.match.span.start)
    for (const edit of ordered) {
      const text = expected[edit.ordinal]
      expected[edit.ordinal] = text.slice(0, edit.match.span.start) + edit.amendment.replacement
        + text.slice(edit.match.span.start + edit.match.span.length)
      // The complete authenticated quote localizes the occurrence. Unchanged
      // context keeps its own runs; only the changed text is mutated.
      if (!edit.change.match.text && !edit.change.replacement) continue
      const result = opened.session.replaceMatch(edit.change.match, edit.change.replacement)
      if (!result.success) throw new Error(`docx_mutation_failed:${result.error?.code}`)
    }
    const output = opened.session.save()
    const reopened = await openRevision(output)
    try {
      if (JSON.stringify(reopened.paragraphs.map(p => p.match.text)) !== JSON.stringify(expected))
        throw new Error("docx_saved_text_mismatch")
      const changed = new Set(edits.map(edit => edit.ordinal))
      verifyPreservation(opened, reopened, changed)
      return {
        bytes: output,
        report: {
          source_sha256: packet.revision.source_sha256,
          output_sha256: reopened.revision.source_sha256,
          amendments: edits.map(edit => ({ amendment_id: edit.amendment.id, target_handle: edit.handle,
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
  } finally { opened.session.close() }
}

export function changedDocxSpan(match: Match, replacement: string) {
  // Native offsets count UTF-16 units, but trim whole code points so a shared
  // surrogate never becomes an unchanged half of a changed character.
  const before = Array.from(match.text)
  const after = Array.from(replacement)
  let first = 0
  let lastBefore = before.length
  let lastAfter = after.length
  while (first < lastBefore && first < lastAfter && before[first] === after[first]) first++
  while (lastBefore > first && lastAfter > first && before[lastBefore - 1] === after[lastAfter - 1]) {
    lastBefore--
    lastAfter--
  }
  const text = before.slice(first, lastBefore).join("")
  const inserted = after.slice(first, lastAfter).join("")
  // Native replacement needs a nonempty span. A pure insertion retains one
  // adjacent code point inside the authenticated quote as its formatting anchor.
  const anchor = !text && inserted ? Math.min(first, before.length - 1) : first
  const offset = before.slice(0, anchor).join("").length
  const replaced = !text && inserted ? before[anchor]! : text
  const changed = !text && inserted ? first < before.length ? inserted + replaced : replaced + inserted : inserted
  const formatting = match.fragments.find((fragment, index) => match.fragments.slice(0, index + 1)
    .reduce((length, row) => length + row.text.length, 0) > offset)?.formatting ?? match.fragments.at(-1)?.formatting
  return { match: { ...match, text: replaced, span: { start: match.span.start + offset, length: replaced.length } },
    replacement: changed, formatting }
}

function localize(opened: Awaited<ReturnType<typeof openRevision>>, proposed: DocxAmendment[], limits: Limits) {
  const amendments = z.array(amendmentSchema).min(1).parse(proposed)
  if (new Set(amendments.map(a => a.id)).size !== amendments.length) throw new Error("duplicate_amendment_id")
  const bounded = { maxCandidates: limits.maxCandidates ?? 512, maxContextCharacters: limits.maxContextCharacters ?? 200_000 }
  if (!Number.isSafeInteger(bounded.maxCandidates) || bounded.maxCandidates < 1
    || !Number.isSafeInteger(bounded.maxContextCharacters) || bounded.maxContextCharacters < 1) throw new Error("invalid_localization_limits")
  const bindings = new Map<string, { handle: string; amendment: DocxAmendment; match: Match; ordinal: number }>()
  const candidates = amendments.flatMap(amendment => {
    if (/[\r\n\t]/.test(amendment.before_quote)) throw new Error("unsupported_cross_block_quote")
    const matches = opened.session.grep(literal(amendment.before_quote), { scope: 63, whitespace: 0, contextChars: 160 })
    if (!matches.length) throw new Error(`before_quote_not_found:${amendment.id}`)
    return matches.map(match => {
      const ordinal = opened.paragraphs.findIndex(p => p.match.enclosingAnchor.id === match.enclosingAnchor.id)
      if (ordinal < 0 || match.text !== amendment.before_quote) throw new Error("unowned_docx_match")
      const handle = `target-${digest(JSON.stringify([opened.revision.source_sha256, ordinal, match.span.start, match.span.length, match.text])).slice(0, 24)}`
      bindings.set(`${amendment.id}:${handle}`, { handle, amendment, match, ordinal })
      return { amendment_id: amendment.id, target_handle: handle, exact_text: match.text,
        context_before: match.contextBefore, context_after: match.contextAfter,
        paragraph: opened.revision.paragraphs[ordinal],
        touched_formatting: match.fragments.map(fragment => ({ text: fragment.text, formatting: fragment.formatting })),
      }
    })
  })
  if (candidates.length > bounded.maxCandidates) throw new Error(`docx_candidate_limit_exceeded:${candidates.length}`)
  if (JSON.stringify(candidates).length > bounded.maxContextCharacters) throw new Error("docx_candidate_context_limit_exceeded")
  return { packet: { schema, revision: opened.revision, amendments, limits: bounded, candidates }, bindings }
}

export async function openRevision(source: Uint8Array) {
  const parts = unzipSync(source)
  const documents = Object.fromEntries(Object.entries(parts).filter(([name]) => name.endsWith(".xml"))
    .map(([name, bytes]) => [name, parse(strFromU8(bytes))]))
  if (!documents["word/document.xml"]) throw new Error("invalid_docx_package")
  if (Object.values(documents).some(doc => ["ins", "del", "moveFrom", "moveTo"].some(name => elements(doc, name).length)))
    throw new Error("unsupported_existing_tracked_revisions")
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
      if (paragraph.match.enclosingAnchor.kind === "h") {
        const level = (paragraph.metadata.outlineLevel ?? 0) + 1
        while (headingStack.length && headingStack.at(-1)!.level >= level) headingStack.pop()
        headingStack.push({ level, text: paragraph.match.text })
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
        headings: headingStack.map(heading => ({ ...heading })),
        table: tableContext,
        style_id: paragraph.metadata.styleId ?? null,
        formatting: paragraph.match.fragments.map(fragment => ({ text: fragment.text, formatting: fragment.formatting })),
      }
    })
    const markdown = views.map((view, ordinal) => {
      const prefix = paragraphs[ordinal].match.enclosingAnchor.kind === "h" ? `${"#".repeat(view.headings.at(-1)!.level)} ` : ""
      const context = view.table ? `> Table ${view.table.table}, row ${view.table.row}, column ${view.table.column}\n\n` : ""
      return context + prefix + view.formatting.map(fragment => {
        const text = fragment.text.replace(/[\\`*_{}[\]<>|]/g, "\\$&")
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
function elements(node: Document | Element, name: string) { return Array.from(node.getElementsByTagNameNS(W, name)) }
function emptyParagraph(node: Element) {
  return elements(node, "t").every(text => !text.textContent)
    && !["tab", "br", "instrText"].some(name => elements(node, name).length)
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
