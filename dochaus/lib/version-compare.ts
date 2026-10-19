import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { normalizeParagraphText } from "./drafting-amendment"
import { docxodus } from "./docxodus"

// Compares two versions of one Word document. The blackline is Word Compare's:
// left → right with every pending tracked change in either file accepted first.
// A second pass compares that same accepted left with right after accepting only
// the proposals of ours it still carries and rejecting its other revisions, which
// leaves exactly the edits right made without tracking. Each of left's own
// pending proposals is reported as still pending, countered, accepted or
// rejected without tracking, from those same revisions and the untracked edits
// in its own paragraph. Pure: no model call, no file writes.

export type CompareOptions = { ignoreFormatting: boolean; detectMoves: boolean; headersFooters: boolean }

export type ChangeLocation = {
  scope: "body" | "header" | "footer" | "footnote" | "endnote"
  /** 1-based paragraph ordinal within the scope, counted in `side`. */
  paragraph: number
  side: "left" | "right"
  /** Nearest body heading at or before the paragraph, with its Word numbering. */
  heading?: string
  /** Word auto-number of the paragraph itself. */
  number?: string
}

export type ChangeHunk = {
  id: string
  kind: "modified" | "inserted" | "deleted" | "moved"
  location: ChangeLocation
  before?: string
  after?: string
  untracked: boolean
}

export type SideRevisions = { count: number; authors: string[] }

/** What became of one of left's pending tracked insertions or deletions in right. */
export type ProposalOutcome = {
  /** P1, P2… in left's document order. */
  id: string
  kind: "inserted" | "deleted"
  /** The proposed inserted text, or the text proposed for deletion. */
  text: string
  author: string
  /**
   * pending: right still carries it as ours. countered: carried, but right also
   * strikes or restores it with a tracked change of its own. accepted: its
   * effect is in right as plain text. rejected: removed or changed in right with
   * no tracked change, which is what Word's Reject Change leaves.
   */
  outcome: "pending" | "accepted" | "rejected" | "countered"
  /**
   * U-id of an untracked edit in the proposal's own paragraph: for a rejected
   * proposal the one that removed it, otherwise the first there.
   */
  untracked?: string
  /** Its paragraph in accept-all(left), counted as the hunks count (a deleted paragraph: in left); paragraph 0 if not located. */
  location: ChangeLocation
}

export type VersionComparison = {
  blackline: Uint8Array
  changes: ChangeHunk[]
  untracked: ChangeHunk[]
  attribution: "authors" | "label"
  left: SideRevisions
  right: SideRevisions
  proposals: ProposalOutcome[]
}

type Dx = Awaited<ReturnType<typeof docxodus>>
type DiffRevision = NonNullable<Awaited<ReturnType<Dx["docxDiffCompareProducts"]>>["revisions"]>[number]
type InputRevision = Awaited<ReturnType<Dx["getRevisions"]>>[number]
type ListedRevision = ReturnType<ReturnType<Dx["openDocxSession"]>["listRevisions"]>[number]
type Hunk = ChangeHunk & { left: string[]; right: string[] }
type Block = { location: Omit<ChangeLocation, "side">; text: string }
type Carried = ListedRevision & { countered: boolean }

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const STORY = /^word\/(?:document|header\d*|footer\d*|footnotes|endnotes)\.xml$/
// ProjectionScopes: body, headers, footers, footnotes and endnotes; not comments.
const STORIES = 31
const HEADERS_FOOTERS = 6
// Anchor kinds that are paragraphs (li: numbered list item); others are containers.
const PARAGRAPHS = new Set(["p", "h", "li"])
// Property-change markup records formatting only, never wording.
const FORMAT_CHANGES = new Set(["rPrChange", "pPrChange", "sectPrChange", "tblPrChange", "tblPrExChange", "trPrChange", "tcPrChange", "tblGridChange"])
const REVISION = new Map<string, "ins" | "del">([["ins", "ins"], ["moveTo", "ins"], ["del", "del"], ["moveFrom", "del"]])

export const untrackedAuthor = (label: string) => `Not tracked (${label})`

export async function compareVersions(
  left: { bytes: Uint8Array; label: string },
  right: { bytes: Uint8Array; label: string },
  options: CompareOptions,
  date: string,
): Promise<VersionComparison> {
  const dx = await docxodus()
  const leftInputs = await dx.getRevisions(left.bytes)
  const rightInputs = await dx.getRevisions(right.bytes)
  // Both comparisons read left accepted with its paragraph anchor ids kept, so
  // their anchors and left's own revisions address one set of paragraphs; they
  // pre-accept right, so blackline anchors address accept-all(right) and the
  // check's received.
  const { view: leftView, sent: sentRevisions } = leftInputs.length ? acceptLeft(dx, left.bytes) : { view: left.bytes, sent: [] }
  const rightView = rightInputs.length ? await dx.docxDiffAcceptRevisions(right.bytes) : right.bytes
  const carried = rightInputs.length && leftInputs.length ? acceptOurs(dx, sentRevisions, right.bytes) : { bytes: right.bytes, ours: [] }
  const received = rightInputs.length ? await dx.docxDiffRejectRevisions(carried.bytes) : right.bytes
  const settings = { preAcceptInputRevisions: true, authorForRevisions: right.label, detectMoves: options.detectMoves,
    compareHeadersFooters: options.headersFooters, dateTimeForRevisions: new Date(date).toISOString().replace(/\.\d{3}Z$/, "Z") }
  const compared = await dx.docxDiffCompareProducts(leftView, right.bytes, settings, ["redline", "revisions"])
  // With no tracked changes in right, every difference is untracked: the check is the blackline's own comparison.
  const checked = received === right.bytes ? compared : await dx.docxDiffCompareProducts(leftView, received, settings, ["redline", "revisions"])

  const plain = options.ignoreFormatting ? withoutFormatChanges(compared.redline!) : compared.redline!
  const relabelled = relabel(plain, checked.redline!, rightInputs, right.label)
  // Never ship a partially re-labelled file: on any miss, the plain Word Compare blackline.
  const attribution = relabelled && await roundTrips(relabelled, leftView, rightView, options.headersFooters) ? "authors" : "label"

  const changed = compared.revisions!.filter(row => !options.ignoreFormatting || row.revisionType !== "FormatChanged")
  const unchecked = checked.revisions!.filter(row => row.revisionType !== "FormatChanged")
  const anchors = (rows: DiffRevision[]) => rows.flatMap(row => [row.leftAnchor ?? [], row.rightAnchor ?? []].flat())
  const leftSide = await side(leftView, anchors([...changed, ...unchecked]))
  const rightSide = await side(rightView, anchors(received === rightView ? [...changed, ...unchecked] : changed))
  const receivedSide = received === rightView ? rightSide : await side(received, anchors(unchecked))
  const untracked = hunks(unchecked, leftSide, receivedSide, "U")
  const changes = hunks(changed, leftSide, rightSide, "C")
  // A change holds an untracked edit when an untracked hunk shares its left
  // paragraph or, for an inserted paragraph, its right one. Both passes read the
  // same left; right anchors are content-derived, so an inserted paragraph reads
  // the same in both unless a tracked change touches it. Derived the same way
  // under both attributions, so the flag never depends on whether re-labelling
  // held.
  const leftKeys = new Set(untracked.flatMap(hunk => hunk.left)), rightKeys = new Set(untracked.flatMap(hunk => hunk.right))
  for (const hunk of changes)
    hunk.untracked = hunk.left.length ? hunk.left.some(anchor => leftKeys.has(anchor)) : hunk.right.some(anchor => rightKeys.has(anchor))
  for (const hunk of untracked) hunk.untracked = true

  const disclosed = (rows: InputRevision[]) => ({ count: rows.length, authors: unique(rows.map(row => row.author)) })
  const exported = ({ left: _left, right: _right, ...hunk }: Hunk): ChangeHunk => hunk
  return { blackline: attribution === "authors" ? relabelled! : plain, changes: changes.map(exported), untracked: untracked.map(exported),
    attribution, left: disclosed(leftInputs), right: disclosed(rightInputs), proposals: await outcomes(left.bytes, sentRevisions, carried.ours, untracked, leftSide) }
}

// A revision is ours when its type, author, date and text equal one pending in left.
const revisionKey = (row: ListedRevision) => JSON.stringify([row.type, row.author, row.date ?? "", row.text])
const tally = (rows: ListedRevision[]) => {
  const counts = new Map<string, number>()
  for (const row of rows) counts.set(revisionKey(row), (counts.get(revisionKey(row)) ?? 0) + 1)
  return counts
}
const anchor = (row: ListedRevision) => row.anchorId ?? row.affectedAnchors[0]?.id
const flat = (text = "") => text.replace(/[¶\s]+/g, " ").trim()
// Copies of a non-empty text in another; one that starts or ends with a letter,
// combining mark or digit must not run on into one there.
function occurrences(text: string, within: string) {
  const word = /[\p{L}\p{M}\p{N}]/u
  let count = 0
  for (let at = within.indexOf(text); at >= 0;) {
    const inside = word.test(text[0]!) && word.test(within[at - 1] ?? "") || word.test(text.at(-1)!) && word.test(within[at + text.length] ?? "")
    if (!inside) count++
    at = within.indexOf(text, inside ? at + 1 : at + text.length)
  }
  return count
}

// Left with every revision accepted in one session that writes each
// paragraph's anchor id into the file, so a revised paragraph keeps the id
// left's own revisions address (Docxodus's accept-all derives a new one from its
// content), and the revisions it listed. A refused accept leaves that revision
// to the comparisons, which accept their inputs first.
function acceptLeft(dx: Dx, bytes: Uint8Array) {
  const session = dx.openDocxSession(bytes, { persistAnchorIds: true })
  try {
    const sent = session.listRevisions()
    for (const row of sent) session.acceptRevision(row.id)
    return { view: session.save(), sent }
  } finally {
    session.close()
  }
}

// Right with the proposals of ours it still carries accepted, as left's view has
// them, and those carried revisions. A carried one is countered when a revision
// of theirs of the opposite type in the same paragraph holds its text: our
// insertion struck, or our deletion restored, with tracking. If any accept is
// refused, right's bytes come back unchanged and the check rejects every
// revision.
function acceptOurs(dx: Dx, sent: ListedRevision[], right: Uint8Array) {
  const pending = tally(sent)
  const session = dx.openDocxSession(right, { persistAnchorIds: false })
  try {
    const rows = session.listRevisions()
    const carried = new Set(rows.filter(row => {
      const count = pending.get(revisionKey(row)) ?? 0
      pending.set(revisionKey(row), count - 1)
      return count > 0
    }))
    const opposite: Record<string, string> = { insert: "delete", delete: "insert" }
    const ours: Carried[] = [...carried].map(row => ({ ...row, countered: !!flat(row.text) && !!anchor(row) && rows.some(other => !carried.has(other)
      && other.type === opposite[row.type] && anchor(other) === anchor(row) && occurrences(flat(row.text), flat(other.text)) > 0) }))
    if (!ours.length) return { bytes: right, ours }
    return { bytes: ours.every(row => session.acceptRevision(row.id).success) ? session.save() : right, ours }
  } finally {
    session.close()
  }
}

// What became of each of left's pending insertions and deletions, judged in its
// own paragraph (its anchor survives accepting left's revisions) by the
// untracked hunks there. Such a hunk removes it when it holds more copies of the
// inserted text before than after, or the other way round for deleted text; each
// surplus copy goes to one proposal, longer texts first, and copies inside a
// longer claimed text are not a shorter one's. A deleted whole paragraph has no
// anchor there: an untracked paragraph with exactly its text restores it. A
// proposal right carries as ours is pending (countered if right also strikes or
// restores it), carried copies going first to proposals nothing removed;
// otherwise rejected when removed, else accepted, its effect in right. Format,
// move, paragraph-mark and blank revisions are not proposals.
async function outcomes(left: Uint8Array, sent: ListedRevision[], ours: Carried[], untracked: Hunk[], leftSide: Map<string, Block>): Promise<ProposalOutcome[]> {
  const proposed = sent.filter(row => (row.type === "insert" || row.type === "delete") && row.family !== "paragraph_mark" && flat(row.text))
  if (!proposed.length) return []
  const homeless = proposed.flatMap(row => leftSide.has(anchor(row) ?? "") ? [] : anchor(row) ?? [])
  const located = homeless.length ? await side(left, homeless) : new Map<string, Block>()
  const copies = new Map<string, boolean[]>()
  for (const row of ours) copies.set(revisionKey(row), [...copies.get(revisionKey(row)) ?? [], row.countered])
  // Surplus copies of a text a hunk removed, keyed by hunk id, type and text.
  const claims = new Map<string, { hunk: Hunk; type: string; text: string; left: number }>()
  const rows = proposed.map(row => {
    const text = flat(row.text), home = leftSide.get(anchor(row) ?? ""), whole = !home && row.type === "delete"
    const tied = home ? untracked.filter(hunk => hunk.left.includes(anchor(row)!))
      : whole ? untracked.filter(hunk => !hunk.left.length && (hunk.after ?? "").split("\n").some(line => flat(line) === text)) : []
    const [gone, back] = row.type === "insert" ? (["before", "after"] as const) : (["after", "before"] as const)
    const removals = tied.flatMap(hunk => {
      const surplus = whole ? 1 : occurrences(text, flat(hunk[gone])) - occurrences(text, flat(hunk[back])), key = `${hunk.id}|${row.type}|${text}`
      if (surplus > 0 && !claims.has(key)) claims.set(key, { hunk, type: row.type, text, left: surplus })
      return surplus > 0 ? [claims.get(key)!] : []
    })
    const block = home ?? located.get(anchor(row) ?? "")
    return { row, text, tied, removals, location: block ? { ...block.location, side: "left" as const } : { scope: scopeOf(row.scope), paragraph: 0, side: "left" as const } }
  })
  const carried = new Map<number, boolean>()
  for (const removed of [false, true]) rows.forEach(({ row, removals }, index) => {
    const copy = !!removals.length === removed ? copies.get(revisionKey(row))?.shift() : undefined
    if (copy !== undefined) carried.set(index, copy)
  })
  const claimed = new Map<number, Hunk>()
  for (const index of rows.map((_, index) => index).toSorted((a, b) => rows[b]!.text.length - rows[a]!.text.length)) {
    const claim = carried.has(index) ? undefined : rows[index]!.removals.find(entry => entry.left > 0)
    if (!claim) continue
    claimed.set(index, claim.hunk)
    claim.left--
    for (const entry of claims.values())
      if (entry !== claim && entry.hunk === claim.hunk && entry.type === claim.type) entry.left -= occurrences(entry.text, claim.text)
  }
  return rows.map(({ row, tied, location }, index) => {
    const countered = carried.get(index), claim = claimed.get(index), link = claim?.id ?? tied[0]?.id
    return { id: `P${index + 1}`, kind: row.type === "insert" ? "inserted" : "deleted", text: row.text.replace(/¶$/, ""), author: row.author,
      outcome: countered !== undefined ? countered ? "countered" : "pending" : claim ? "rejected" : "accepted", ...link && { untracked: link }, location }
  })
}

// The re-labelled blackline must still reject to left and accept to right, both
// with their own tracked changes accepted. Compared on normalised paragraph text
// per story; headers and footers only when compared, since otherwise the
// blackline carries left's.
export async function roundTrips(blackline: Uint8Array, left: Uint8Array, right: Uint8Array, headersFooters: boolean) {
  const dx = await docxodus()
  const scope = headersFooters ? STORIES : STORIES & ~HEADERS_FOOTERS
  const text = (bytes: Uint8Array) => {
    const session = dx.openDocxSession(bytes, { persistAnchorIds: false })
    try {
      return JSON.stringify(session.grep("^[\\s\\S]*$", { scope }).flatMap(match => {
        const paragraph = normalizeParagraphText(match.text)
        return paragraph ? [[match.enclosingAnchor.scope, paragraph]] : []
      }))
    } finally {
      session.close()
    }
  }
  return text(await dx.docxDiffRejectRevisions(blackline)) === text(left) && text(await dx.docxDiffAcceptRevisions(blackline)) === text(right)
}

// Paragraph-level hunks. Revisions that share a paragraph on either side, or a
// move, form one hunk, so a split, a merge or a moved paragraph edited in place
// reads as one change. A hunk whose revisions carry no text (a lone paragraph
// mark) is left out; the blackline still shows it.
function hunks(revisions: DiffRevision[], leftSide: Map<string, Block>, rightSide: Map<string, Block>, prefix: "C" | "U"): Hunk[] {
  const parent = revisions.map((_, index) => index)
  const root = (index: number): number => parent[index] === index ? index : (parent[index] = root(parent[index]))
  const first = new Map<string, number>()
  revisions.forEach((row, index) => {
    for (const key of [row.leftAnchor && `L${row.leftAnchor}`, row.rightAnchor && `R${row.rightAnchor}`, row.moveGroupId !== undefined && `M${row.moveGroupId}`]) {
      if (!key) continue
      const other = first.get(key)
      if (other === undefined) first.set(key, index)
      else parent[root(index)] = root(other)
    }
  })
  const groups = new Map<number, DiffRevision[]>()
  revisions.forEach((row, index) => groups.set(root(index), [...groups.get(root(index)) ?? [], row]))
  return [...groups.values()].filter(rows => rows.some(row => row.text.replace(/\n/g, ""))).map((rows, index) => {
    const left = unique(rows.flatMap(row => row.leftAnchor ?? [])), right = unique(rows.flatMap(row => row.rightAnchor ?? []))
    const kind = rows.some(row => row.revisionType === "Moved") ? "moved" : left.length && right.length ? "modified" : left.length ? "deleted" : "inserted"
    const at = (anchor: string | undefined, side: "left" | "right") => {
      const block = anchor ? (side === "left" ? leftSide : rightSide).get(anchor) : undefined
      return block && { ...block.location, side }
    }
    // An anchor outside the projection (none seen so far) still reports the change, without an ordinal.
    const location: ChangeLocation = (kind === "deleted" ? at(left[0], "left") : at(right[0], "right") ?? at(left[0], "left"))
      ?? { scope: scopeOf((right[0] ?? left[0])!.split(":")[1]!), paragraph: 0, side: right.length ? "right" : "left" }
    const text = (anchors: string[], side: Map<string, Block>) => anchors.length ? { text: anchors.map(anchor => side.get(anchor)?.text ?? "").join("\n") } : undefined
    const before = kind === "inserted" ? undefined : text(left, leftSide), after = kind === "deleted" ? undefined : text(right, rightSide)
    return { id: `${prefix}${index + 1}`, kind, location, ...before && { before: before.text }, ...after && { after: after.text }, untracked: false, left, right }
  })
}

// Location and full text of every block one comparison side addresses, read
// from the same bytes the comparison consumed: anchors are stable per bytes.
// Container anchors (a whole table, row or section) take the ordinal of their
// first paragraph and the text of their paragraphs.
async function side(bytes: Uint8Array, wanted: string[]) {
  const dx = await docxodus()
  const session = dx.openDocxSession(bytes, { persistAnchorIds: false })
  try {
    // The session projection carries Word's resolved auto-number on each anchor
    // (its declared type omits the field); the markdown projection does not.
    const anchors = Object.entries(session.project().anchorIndex as Record<string, { kind: string; scope: string; textPreview: string; autoNumberPrefix?: string }>)
    const texts = new Map(session.grep("^[\\s\\S]*$", { scope: STORIES }).map(match => [match.enclosingAnchor.id, match.text]))
    const containers = new Set(wanted)
    const blocks = new Map<string, Block>()
    const ordinals = new Map<string, number>()
    let heading: string | undefined
    let waiting: { id: string; scope: string; heading?: string; text: string }[] = []
    for (const [id, anchor] of anchors) {
      const body = anchor.scope === "body"
      if (!PARAGRAPHS.has(anchor.kind)) {
        if (containers.has(id)) waiting.push({ id, scope: anchor.scope, heading: body ? heading : undefined, text: containerText(session.raw.getXml(id)) })
        continue
      }
      const paragraph = (ordinals.get(anchor.scope) ?? 0) + 1
      ordinals.set(anchor.scope, paragraph)
      const text = texts.get(id) ?? ""
      if (anchor.kind === "h" && body) heading = [anchor.autoNumberPrefix, text || anchor.textPreview].filter(Boolean).join(" ")
      blocks.set(id, { text, location: { scope: scopeOf(anchor.scope), paragraph,
        ...body && heading ? { heading } : {}, ...anchor.autoNumberPrefix ? { number: anchor.autoNumberPrefix } : {} } })
      for (const row of waiting.filter(row => row.scope === anchor.scope))
        blocks.set(row.id, { text: row.text, location: { scope: scopeOf(row.scope), paragraph, ...row.heading ? { heading: row.heading } : {} } })
      waiting = waiting.filter(row => row.scope !== anchor.scope)
    }
    for (const row of waiting)
      blocks.set(row.id, { text: row.text, location: { scope: scopeOf(row.scope), paragraph: Math.max(1, ordinals.get(row.scope) ?? 0), ...row.heading ? { heading: row.heading } : {} } })
    return blocks
  } finally {
    session.close()
  }
}

// A row reads as its cells separated by " | "; a table as its rows; anything
// else as its paragraphs, one per line.
function containerText(xml: string) {
  const node = new DOMParser().parseFromString(xml, "application/xml").documentElement
  const paragraphs = (element: Element) => elements(element, "p").map(paragraph => elements(paragraph, "*")
    .map(child => child.localName === "t" ? child.textContent ?? "" : child.localName === "tab" && (child.parentNode as Element).localName === "r" ? "\t" : "").join(""))
  const row = (element: Element) => childElements(element, "tc").map(cell => paragraphs(cell).join(" ")).join(" | ")
  return node.localName === "tr" ? row(node) : node.localName === "tbl" ? childElements(node, "tr").map(row).join("\n") : paragraphs(node).join("\n")
}

// Re-labels the plain blackline. A revision whose key matches one in the
// check's comparison is an edit right made without tracking and is labelled as
// such. Any other revision takes its input author (and date) only
// when unambiguous: the right file has a single author, or its text matches
// exactly one input revision. Otherwise it keeps the side label, as do
// formatting changes. Returns nothing unless every untracked revision found its
// counterpart.
function relabel(plain: Uint8Array, check: Uint8Array, inputs: InputRevision[], label: string) {
  const parts = unzipSync(plain)
  const docs = storyDocuments(parts)
  const pending = new Map<string, number>()
  for (const row of carriers(storyDocuments(unzipSync(check)))) if (row.text) pending.set(row.key, (pending.get(row.key) ?? 0) + 1)
  const authors = unique(inputs.map(row => row.author)), dates = unique(inputs.map(row => row.date ?? ""))
  const sign = (node: Element, author: string, date?: string) => {
    node.setAttributeNS(W, "w:author", author)
    if (date) node.setAttributeNS(W, "w:date", date)
  }
  for (const row of carriers(docs)) {
    if (!row.text) continue
    const count = pending.get(row.key) ?? 0
    if (count) {
      pending.set(row.key, count - 1)
      sign(row.node, untrackedAuthor(label))
      continue
    }
    const matches = inputs.filter(input => (input.type === "move" || input.type === (row.kind === "ins" ? "insert" : "delete"))
      && (input.family === "paragraph_mark" ? "¶" : input.text.replace(/¶$/, "")) === row.text)
    if (matches.length === 1) sign(row.node, matches[0].author, matches[0].date)
    else if (authors.length === 1) sign(row.node, authors[0], dates.length === 1 ? dates[0] : undefined)
  }
  if ([...pending.values()].some(count => count > 0)) return undefined
  for (const [name, doc] of docs) {
    // Word also reads a move's author from its range start.
    let starts: Element[] = []
    for (const node of elements(doc.documentElement, "*")) {
      if (node.localName === "moveFromRangeStart" || node.localName === "moveToRangeStart") starts.push(node)
      else if ((node.localName === "moveFrom" || node.localName === "moveTo") && starts.length) {
        const own = starts.filter(start => start.localName === `${node.localName}RangeStart`)
        for (const start of own) sign(start, node.getAttributeNS(W, "author") ?? label, node.getAttributeNS(W, "date") ?? undefined)
        starts = starts.filter(start => !own.includes(start))
      }
    }
    parts[name] = strToU8(new XMLSerializer().serializeToString(doc))
  }
  return zipSync(parts, { mtime: new Date(1980, 0, 1) })
}

// Revision carriers in document order, keyed by part, kind, offset in left's
// text and their own text. Both comparisons reject to accept-all(left), so an
// edit right made without tracking has the same key in the blackline and in the
// check.
function carriers(docs: Map<string, Document>) {
  const rows: { key: string; kind: "ins" | "del"; text: string; node: Element }[] = []
  for (const [part, doc] of docs) {
    let offset = 0
    const open: { kind: "ins" | "del"; start: number; text: string; node: Element }[] = []
    const add = (row: (typeof open)[number]) => rows.push({ key: `${part}|${row.kind}|${row.start}|${row.text}`, kind: row.kind, text: row.text, node: row.node })
    const visit = (node: Element) => {
      const name = node.namespaceURI === W ? node.localName : ""
      const kind = REVISION.get(name)
      // Paragraph-mark revisions sit in pPr/rPr and are read with their paragraph.
      if (kind && (node.parentNode as Element).localName !== "rPr") {
        const row = { kind, start: offset, text: "", node }
        open.push(row)
        childElements(node).forEach(visit)
        open.pop()
        return add(row)
      }
      if (name === "t" || name === "delText" || name === "tab" && (node.parentNode as Element).localName === "r") {
        const text = name === "tab" ? "\t" : node.textContent ?? ""
        const current = open.at(-1)
        if (current) current.text += text
        if (!current || current.kind === "del") offset += text.length
        return
      }
      childElements(node).forEach(visit)
      if (name !== "p") return
      const markProperties = childElements(node, "pPr").flatMap(properties => childElements(properties, "rPr"))[0]
      const mark = markProperties && childElements(markProperties).find(child => REVISION.has(child.localName))
      const markKind = mark && REVISION.get(mark.localName)
      if (mark && markKind) add({ kind: markKind, start: offset, text: "¶", node: mark })
      if (markKind !== "ins") offset += 1
    }
    visit(doc.documentElement)
  }
  return rows
}

function withoutFormatChanges(bytes: Uint8Array) {
  const parts = unzipSync(bytes)
  let changed = false
  for (const [name, doc] of storyDocuments(parts)) {
    const marks = elements(doc.documentElement, "*").filter(node => FORMAT_CHANGES.has(node.localName))
    for (const node of marks) node.parentNode!.removeChild(node)
    if (marks.length) {
      parts[name] = strToU8(new XMLSerializer().serializeToString(doc))
      changed = true
    }
  }
  return changed ? zipSync(parts, { mtime: new Date(1980, 0, 1) }) : bytes
}

function storyDocuments(parts: Record<string, Uint8Array>) {
  return new Map(Object.keys(parts).filter(name => STORY.test(name)).toSorted()
    .map(name => [name, new DOMParser().parseFromString(strFromU8(parts[name]!), "application/xml")] as const))
}

function scopeOf(scope: string): ChangeLocation["scope"] {
  return scope.startsWith("hdr") ? "header" : scope.startsWith("ftr") ? "footer" : scope.startsWith("fn") ? "footnote" : scope.startsWith("en") ? "endnote" : "body"
}

const elements = (node: Element, name: string) => Array.from(node.getElementsByTagNameNS(W, name))
const childElements = (node: Element, name?: string) => Array.from(node.childNodes)
  .filter((child): child is Element => child.nodeType === 1 && (!name || (child as Element).namespaceURI === W && (child as Element).localName === name))
const unique = <T,>(rows: T[]) => [...new Set(rows)]
