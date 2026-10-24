import { readFile } from "node:fs/promises"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { z } from "zod"
import { changedRange, openRevision } from "./docx-amendments"
import { paragraphFields, refInstruction, runText, unlink } from "./docx-fields"
import { amendmentSchema, amendmentParagraphs, amendmentText, ensureParagraphPointers, normalizeParagraphText, owningParagraph, paragraphPointers, paragraphSha256, type Amendment } from "./drafting-amendment"
import { writeDurableBytes } from "./durable-json"
import { FIELD_HIGHLIGHT_FILL, OPEN_ITEM_HIGHLIGHT_FILL, hasOpenItemMarker } from "./open-items"
import { OPEN_ITEM_PATTERN } from "./marker-patterns"

export type AppliedAmendment = { id: string; pid: string; mode: "code" | "already_applied"; result_pids?: string[] }
// What a reviewer needs to act on an unapplied amendment. `original` is the text
// the amendment was written against (unknown when stale or deleted), `current`
// the paragraph the applier saw, `target` the requested text, `reason` the
// amendment's reason and `issue` what went wrong.
export type UnappliedDetail = { original?: string; current?: string; target: string; reason: string; issue: string }
export type UnappliedAmendment = {
  id: string
  pid: string
  reason: "stale" | "target_deleted" | "apply_failed"
  detail?: UnappliedDetail
}
type Outcome = AppliedAmendment | UnappliedAmendment
type Opened = Awaited<ReturnType<typeof openRevision>>
type Edit = { ordinal: number; initial: string; final: string[]; rows: Amendment[] }

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const BODY = "word/document.xml"
// Above this many token insertions plus deletions a paragraph is rewritten as a
// whole; the diff's memory grows with the square of this bound.
const MAX_EDITS = 2000
const HIGHLIGHT_FILLS = new Set([OPEN_ITEM_HIGHLIGHT_FILL, FIELD_HIGHLIGHT_FILL])

// Applies whole-paragraph amendments addressed by paragraph pointer, writing
// each amendment's own text and nothing else. Overlapping changes are
// consolidated before change-review, so a paragraph takes at most one
// amendment: one whose existing_sha256 matches is applied by code, one whose
// text is already there is already_applied, and any other is stale. A later
// amendment on a paragraph an earlier one in the batch changed is also stale,
// never merged. A missing or duplicated pointer is target_deleted; there is no
// text fallback. Edits touch only the changed words, so run formatting outside
// them is kept; a paragraph the engine cannot write or that reads back
// differently is left as it was (apply_failed). A live cross-reference keeps
// its field unless the amendment rewrites its number (unlinkEditedReferences).
// outPath is written atomically.
export async function applyAmendments(input: { docxPath: string; outPath: string; amendments: Amendment[] }) {
  const amendments = requireAmendments(input.amendments)
  const source = new Uint8Array(await readFile(input.docxPath))
  const opened = await openRevision(source)
  try {
    const located = locatePointers(opened)
    const outcomes = new Map<string, Outcome>()
    const resolved = amendments.flatMap(row => {
      const ordinals = located.get(row.pid) ?? []
      if (ordinals.length === 1) return [{ row, ordinal: ordinals[0]! }]
      outcomes.set(row.id, unapplied(row, "target_deleted", {
        issue: ordinals.length ? "More than one paragraph carries this pointer." : "No paragraph carries this pointer.",
      }))
      return []
    })
    // Group by paragraph, not pid, so two pointers left on one paragraph are
    // still one paragraph.
    const edits = [...Map.groupBy(resolved, item => item.ordinal)].flatMap(([ordinal, items]): Edit[] => {
      const initial = opened.paragraphs[ordinal]!.match.text
      const rows = items.map(item => item.row)
      const final = settleParagraph(initial, rows, outcomes)
      return final ? [{ ordinal, initial, final, rows }] : []
    })
    const written = edits.length ? await writeVerified(source, opened, edits.toSorted((a, b) => a.ordinal - b.ordinal)) : { bytes: source, failed: [] }
    for (const failure of written.failed) for (const row of failure.edit.rows) {
      const outcome = outcomes.get(row.id)!
      if (!("mode" in outcome) || outcome.mode === "already_applied") continue
      outcomes.set(row.id, unapplied(row, "apply_failed", { original: failure.edit.initial, current: failure.edit.initial, issue: failure.issue }))
    }
    writeDurableBytes(input.outPath, written.bytes)
    const structural = edits.filter(edit => edit.rows.some(row => row.operation === "delete" || row.operation === "split") && !written.failed.some(failure => failure.edit === edit))
    if (structural.length) {
      await ensureParagraphPointers(input.outPath)
      const final = await openRevision(new Uint8Array(await readFile(input.outPath)))
      try {
        const changed = new Map(edits.filter(edit => !written.failed.some(failure => failure.edit === edit)).map(edit => [edit.ordinal, edit.final]))
        const starts = new Map<number, number>()
        let at = 0
        opened.paragraphs.forEach((_, ordinal) => { starts.set(ordinal, at); at += changed.get(ordinal)?.length ?? 1 })
        for (const edit of structural) for (const row of edit.rows) {
          const outcome = outcomes.get(row.id)!
          if (!("mode" in outcome)) continue
          const start = starts.get(edit.ordinal)!
          outcomes.set(row.id, { ...outcome, result_pids: final.paragraphs.slice(start, start + edit.final.length)
            .flatMap(paragraph => paragraphPointers(paragraph.node).map(pointer => pointer.pid)) })
        }
      } finally { final.session.close() }
    }
    const ordered = amendments.map(row => outcomes.get(row.id)!)
    return {
      applied: ordered.filter((row): row is AppliedAmendment => "mode" in row),
      unapplied: ordered.filter((row): row is UnappliedAmendment => "reason" in row),
    }
  } finally {
    opened.session.close()
  }
}

function requireAmendments(input: Amendment[]) {
  const parsed = z.array(amendmentSchema).safeParse(input)
  if (!parsed.success) throw new Error(`invalid_amendments: ${parsed.error.issues.slice(0, 10)
    .map(issue => `${issue.path.join(".") || "(root)"}: ${issue.message}`).join("; ")}`)
  if (parsed.data.some(row => row.operation === "delete" && !row.sources.length)) throw new Error("invalid_amendments: an explicit deletion requires supporting source passages or settled instructions")
  const duplicate = parsed.data.find((row, index) => parsed.data.findIndex(other => other.id === row.id) !== index)
  if (duplicate) throw new Error(`invalid_amendments: duplicate amendment id ${duplicate.id}`)
  return parsed.data
}

function unapplied(row: Amendment, reason: UnappliedAmendment["reason"], detail: Omit<UnappliedDetail, "target" | "reason">): UnappliedAmendment {
  return { id: row.id, pid: row.pid, reason, detail: { ...detail, target: amendmentText(row), reason: row.reason } }
}

function locatePointers(opened: Opened) {
  const located = new Map<string, number[]>()
  opened.paragraphs.forEach((paragraph, ordinal) => {
    if (paragraph.part !== BODY) return
    for (const row of paragraphPointers(paragraph.node)) located.set(row.pid, [...located.get(row.pid) ?? [], ordinal])
  })
  return located
}

// The paragraph's amendments in input order. The first written against this
// text lands by code and its result is returned. Any later amendment on the
// changed paragraph is an overlap review should have consolidated: it is
// stale, and the paragraph keeps the landed text.
function settleParagraph(initial: string, rows: Amendment[], outcomes: Map<string, Outcome>) {
  let final: string[] | undefined
  for (const target of rows) {
    const requested = amendmentParagraphs(target)
    if (final) outcomes.set(target.id, unapplied(target, "stale", { current: final.join("\n\n"), issue: "An earlier amendment in this batch already changed this paragraph." }))
    else if (requested.length === 1 && (normalizeParagraphText(initial) === normalizeParagraphText(requested[0]!) || paragraphSha256(requested[0]!) === target.existing_sha256))
      outcomes.set(target.id, { id: target.id, pid: target.pid, mode: "already_applied" })
    else if (paragraphSha256(initial) !== target.existing_sha256)
      outcomes.set(target.id, unapplied(target, "stale", { current: initial, issue: "The paragraph no longer matches the text this amendment was written against." }))
    else {
      final = requested
      outcomes.set(target.id, { id: target.id, pid: target.pid, mode: "code" })
    }
  }
  return final
}

// Writes the edits in one engine session and reads the saved file back. A
// paragraph whose edit fails or reads back differently is left as it was and
// the remaining edits are rewritten from the source; each round removes at
// least one paragraph, so this ends. Each round unlinks only the references
// its own edits need unlinked, from the untouched source, so a paragraph left
// as it was keeps its live cross-references and, when no edit is kept, the
// source is returned unchanged.
async function writeVerified(source: Uint8Array, opened: Opened, edits: Edit[]) {
  const failed: { edit: Edit; issue: string }[] = []
  for (let round = 0; ; round++) {
    const kept = edits.filter(edit => !failed.some(row => row.edit === edit))
    if (!kept.length) return { bytes: source, failed }
    const base = round ? await openRevision(source) : opened
    const unlinked = unlinkEditedReferences(base, kept)
    const revision = unlinked ? await openRevision(unlinked) : base
    if (unlinked && base !== opened) base.session.close()
    try {
      const written = writeParagraphs(revision, kept)
      if (written.kind === "failed") {
        failed.push({ edit: written.failed, issue: "The document engine could not write this paragraph." })
        continue
      }
      const bytes = clearResolvedHighlights(written.bytes, new Set(kept.flatMap(edit => edit.rows.map(row => row.pid))))
      const mismatched = await readBack(bytes, opened, kept)
      if (!mismatched.length) return { bytes, failed }
      failed.push(...mismatched.map(edit => ({ edit, issue: "The saved paragraph did not read back as the amended text." })))
    } finally {
      if (revision !== opened) revision.session.close()
    }
  }
}

// Edits in document order. Each paragraph state is written as word-level hunks
// so formatting between separate changes is kept.
function writeParagraphs(opened: Opened, edits: Edit[]) {
  const cache = { matches: undefined as ReturnType<Opened["session"]["grep"]> | undefined }
  const live = (anchor: string) => {
    cache.matches ??= opened.session.grep("^[\\s\\S]*$", { scope: 63, whitespace: 0 })
    return cache.matches.find(row => row.enclosingAnchor.id === anchor)
  }
  const edit = edits.find(row => !writeParagraph(opened, row, live, () => { cache.matches = undefined }))
  return edit ? { kind: "failed" as const, failed: edit } : { kind: "saved" as const, bytes: preserveDeletedCellParagraphs(opened.session.save(), opened, edits) }
}

function writeParagraph(opened: Opened, edit: Edit, live: (anchor: string) => ReturnType<Opened["session"]["grep"]>[number] | undefined, refresh: () => void) {
  const anchor = opened.paragraphs[edit.ordinal]!.match.enclosingAnchor.id
  const final = edit.final
  try {
    if (final.length !== 1 && !safeStructuralParagraph(opened.paragraphs[edit.ordinal]!.node)) return false
    if (!final.length) {
      const result = opened.session.deleteBlock(anchor)
      refresh()
      return result.success
    }
    for (const text of hunkTexts(live(anchor)?.text ?? "", final.join(""))) {
      const match = live(anchor)
      if (!match) return false
      const { offset, replaced, changed } = hunkRange(match.text, text)
      if (!replaced && !changed) continue
      const result = opened.session.replaceMatch({ ...match, text: replaced, span: { start: match.span.start + offset, length: replaced.length } }, changed)
      refresh()
      if (!result.success) return false
    }
    // Split the concatenated literal replacement, preserving its native runs.
    // Work from the end so every offset still addresses the original anchor.
    const boundaries = final.slice(0, -1).map((_, index) => final.slice(0, index + 1).join("").length).reverse()
    for (const offset of boundaries) {
      if (!opened.session.splitParagraph(anchor, offset).success) return false
      refresh()
    }
    return true
  } catch { return false }
}

// A REF field whose number an edit rewrites, and one in a paragraph split or
// deleted, become the text they show, as before the edit; the cross-reference
// check then reads them like any other reference. The text is unchanged, so
// the edits still apply. A deleted clause's bookmark that only body REF fields
// cite goes first, so the engine may delete the clause; those fields stay as
// they are, and the refresh after renumbering (linkCrossReferences) makes each
// an open-item marker, since its number may by then be another clause's. A
// bookmark anything else cites stays, and the engine refuses the deletion.
// Undefined when nothing is changed.
function unlinkEditedReferences(opened: Opened, edits: Edit[]) {
  let changed = false
  const cited = (body: boolean) => new Set(opened.paragraphs.filter(paragraph => (paragraph.part === BODY) === body)
    .flatMap(paragraph => paragraphFields(paragraph.node).map(field => refInstruction(field)?.name ?? "")))
  const elsewhere = new Set([...cited(false), ...Object.values(opened.documents)
    .flatMap(doc => Array.from(doc.getElementsByTagNameNS(W, "hyperlink")).map(node => node.getAttributeNS(W, "anchor") ?? ""))])
  const body = edits.some(edit => !edit.final.length) ? cited(true) : new Set<string>()
  for (const edit of edits) {
    const paragraph = opened.paragraphs[edit.ordinal]!
    if (paragraph.part !== BODY) continue
    const fields = paragraphFields(paragraph.node).filter(field => refInstruction(field))
    const touched = edit.final.length === 1 && runText(paragraph.node) === edit.initial ? touchedRanges(edit.initial, edit.final[0]!) : undefined
    for (const field of fields) if (!touched || touched.some(([start, end]) => start < field.to && field.from < end)) {
      unlink(field)
      changed = true
    }
    if (edit.final.length) continue
    const ends = Array.from(paragraph.node.getElementsByTagNameNS(W, "bookmarkEnd"))
    for (const start of Array.from(paragraph.node.getElementsByTagNameNS(W, "bookmarkStart"))) {
      const name = start.getAttributeNS(W, "name") ?? ""
      const end = ends.find(node => node.getAttributeNS(W, "id") === start.getAttributeNS(W, "id"))
      if (!/^_Ref\d+$/.test(name) || !body.has(name) || elsewhere.has(name) || !end) continue
      start.parentNode!.removeChild(start)
      end.parentNode!.removeChild(end)
      changed = true
    }
  }
  return changed ? zipSync({ ...opened.parts, [BODY]: strToU8(new XMLSerializer().serializeToString(opened.documents[BODY]!)) }, { mtime: new Date(1980, 0, 1) }) : undefined
}

// Where in `from` each word-level hunk to `to` writes, as writeParagraph
// writes them: [start, end) offsets of the original text.
function touchedRanges(from: string, to: string) {
  const ranges: [number, number][] = []
  let current = from
  let shift = 0
  for (const text of hunkTexts(from, to)) {
    const { offset, replaced, changed } = hunkRange(current, text)
    current = text
    if (!replaced && !changed) continue
    ranges.push([offset - shift, offset - shift + replaced.length])
    shift += changed.length - replaced.length
  }
  return ranges
}

// The part of `from` one hunk to `to` rewrites (changedRange). The engine
// writes a replacement in the formatting of the first run its span touches,
// and a pure insertion's span is the one character after it, so words inserted
// just before an open-item marker would be written in the marker's shaded run
// and read as open. Such an insertion takes the character before the marker
// instead, unless that character is in a marker too. Only a pure insertion's
// text ends with the character it keeps. A replacement that runs from before a
// marker into it would keep the marker's end (its closing bracket) in the
// marker's run while writing what comes before it into another run, so the
// marker that text ends would span two runs and go unshaded; it rewrites the
// rest of the marker too.
function hunkRange(from: string, to: string) {
  const range = changedRange(from, to)
  const { offset, replaced, changed } = range
  const markers = markerSpans(from)
  const end = offset + replaced.length
  const entered = markers.find(([start, stop]) => offset < start && start < end && end < stop)
  if (entered) return { offset, replaced: from.slice(offset, entered[1]), changed: changed + from.slice(end, entered[1]) }
  if (!offset || changed === replaced || !changed.endsWith(replaced) || !markers.some(([start]) => start === offset)) return range
  const previous = Array.from(from.slice(0, offset)).at(-1)!
  const at = offset - previous.length
  if (markers.some(([start, end]) => start <= at && at < end)) return range
  return { offset: at, replaced: previous, changed: previous + changed.slice(0, -replaced.length) }
}

// [start, end) of each open-item marker in a paragraph's text.
function markerSpans(text: string) {
  return [...text.matchAll(OPEN_ITEM_PATTERN)].map(marker => [marker.index ?? 0, (marker.index ?? 0) + marker[0].length] as const)
}

// Cardinality edits must not silently discard section breaks, fields, or ranges
// whose other boundary lies outside the target paragraph.
function safeStructuralParagraph(paragraph: Element) {
  if (["sectPr", "fldChar", "instrText"].some(name => paragraph.getElementsByTagNameNS(W, name).length)) return false
  for (const [start, end] of [["bookmarkStart", "bookmarkEnd"], ["commentRangeStart", "commentRangeEnd"], ["moveFromRangeStart", "moveFromRangeEnd"], ["moveToRangeStart", "moveToRangeEnd"]]) {
    const ids = (name: string) => Array.from(paragraph.getElementsByTagNameNS(W, name)).map(node => node.getAttributeNS(W, "id")).sort()
    if (JSON.stringify(ids(start!)) !== JSON.stringify(ids(end!))) return false
  }
  return true
}

// Word requires a paragraph in a table cell even when its only visible text was
// explicitly deleted. Repair only cells affected by successful deletions; the
// empty structural paragraph carries no deleted paragraph pointer.
function preserveDeletedCellParagraphs(bytes: Uint8Array, opened: Opened, edits: Edit[]) {
  const affected = edits.filter(edit => !edit.final.length).flatMap(edit => {
    const node = opened.paragraphs[edit.ordinal]!.node
    const cell = node.parentNode as Element | null
    if (cell?.namespaceURI !== W || cell.localName !== "tc") return []
    return [Array.from(node.ownerDocument!.getElementsByTagNameNS(W, "tc")).indexOf(cell)]
  })
  if (!affected.length) return bytes
  const parts = unzipSync(bytes)
  const doc = new DOMParser().parseFromString(strFromU8(parts[BODY]!), "application/xml")
  const cells = Array.from(doc.getElementsByTagNameNS(W, "tc"))
  let changed = false
  for (const index of affected) {
    const cell = cells[index]
    if (!cell || Array.from(cell.childNodes).some(node => node.nodeType === 1 && (node as Element).namespaceURI === W && (node as Element).localName === "p")) continue
    cell.appendChild(doc.createElementNS(W, "w:p"))
    changed = true
  }
  return changed ? zipSync({ ...parts, [BODY]: strToU8(new XMLSerializer().serializeToString(doc)) }, { mtime: new Date(1980, 0, 1) }) : bytes
}

// Inserted text takes the formatting of the first run it touches, so resolving a
// highlighted open item would leave the answer shaded as open, and words written
// beside a marker into its run would be shaded with it. Clears open-item shading
// from runs of edited paragraphs that hold anything but markers; the apply step
// shades each marker in them on its own, as drafting does.
function clearResolvedHighlights(bytes: Uint8Array, pids: Set<string>) {
  const parts = unzipSync(bytes)
  const doc = new DOMParser().parseFromString(strFromU8(parts[BODY]!), "application/xml")
  const shading = Array.from(doc.getElementsByTagNameNS(W, "shd")).filter(shd => {
    const run = shd.parentNode?.parentNode as Element | null
    if (run?.localName !== "r" || !HIGHLIGHT_FILLS.has((shd.getAttributeNS(W, "fill") ?? "").toUpperCase())) return false
    const paragraph = owningParagraph(run) as Element | null
    if (!paragraph || !paragraphPointers(paragraph).some(row => pids.has(row.pid))) return false
    const text = Array.from(run.getElementsByTagNameNS(W, "t")).map(text => text.textContent ?? "").join("")
    return !hasOpenItemMarker(text) || text.replace(OPEN_ITEM_PATTERN, "") !== ""
  })
  if (!shading.length) return bytes
  for (const shd of shading) shd.parentNode!.removeChild(shd)
  return zipSync({ ...parts, [BODY]: strToU8(new XMLSerializer().serializeToString(doc)) }, { mtime: new Date(1980, 0, 1) })
}

// Edited paragraphs that do not read back as their final text. A change to a
// paragraph that was not edited cannot be attributed and stops the batch.
async function readBack(bytes: Uint8Array, before: Opened, edits: Edit[]) {
  const after = await openRevision(bytes)
  try {
    const edited = new Map(edits.map(edit => [edit.ordinal, edit]))
    const expected = before.paragraphs.flatMap((paragraph, ordinal) =>
      (edited.get(ordinal)?.final ?? [paragraph.match.text]).map(text => ({ text, part: paragraph.part, ordinal })))
    if (after.paragraphs.length !== expected.length) return edits
    const changed = expected.filter((row, index) => after.paragraphs[index].part !== row.part ||
      normalizeParagraphText(after.paragraphs[index].match.text) !== normalizeParagraphText(row.text)).map(row => row.ordinal)
    if (changed.some(ordinal => !edited.has(ordinal))) throw new Error("docx_write_changed_unedited_paragraph")
    return edits.filter(edit => changed.includes(edit.ordinal))
  } finally {
    after.session.close()
  }
}

// Intermediate paragraph texts from `from` to `to`, one per word-level hunk.
// Whitespace runs compare equal, so the document's own spacing is kept. A word
// in an open-item marker never matches one outside a marker: kept, it would
// keep its run, and so the marker's shading where it became clause text, or
// leave a rewritten marker's opening in the clause text's run and the marker
// split across runs and unshaded.
function hunkTexts(from: string, to: string) {
  const a = from.match(/\s+|\S+/g) ?? []
  const b = to.match(/\s+|\S+/g) ?? []
  const keys = (text: string, tokens: string[]) => {
    const markers = markerSpans(text)
    let at = 0
    return tokens.map(token => {
      const start = at
      at += token.length
      return (markers.some(([from, to]) => from < at && start < to) ? "\u0000" : "") + (/^\s/.test(token) ? " " : token)
    })
  }
  const pairs = matchTokens(keys(from, a), keys(to, b))
  if (!pairs) return [to]
  // Hunks separated only by whitespace become one edit.
  const hunks = gaps(pairs, a.length, b.length).reduce<{ start: number; end: number; text: string }[]>((joined, gap) => {
    const previous = joined.at(-1)
    const text = b.slice(gap.from, gap.to).join("")
    if (!previous || !a.slice(previous.end, gap.start).every(token => /^\s/.test(token))) return [...joined, { start: gap.start, end: gap.end, text }]
    previous.text += a.slice(previous.end, gap.start).join("") + text
    previous.end = gap.end
    return joined
  }, [])
  return hunks.map((_, index) => {
    const applied = hunks.slice(0, index + 1)
    const pieces = applied.flatMap((hunk, k) => [a.slice(k ? applied[k - 1]!.end : 0, hunk.start).join(""), hunk.text])
    return pieces.join("") + a.slice(applied.at(-1)!.end).join("")
  })
}

// Unmatched runs between matched pairs: a[start, end) became b[from, to).
function gaps(pairs: [number, number][], n: number, m: number) {
  return [...pairs, [n, m] as [number, number]].flatMap(([i, j], index) => {
    const [pi, pj] = index ? pairs[index - 1]! : [-1, -1]
    return i > pi + 1 || j > pj + 1 ? [{ start: pi + 1, end: i, from: pj + 1, to: j }] : []
  })
}

// Matched index pairs of a shortest token edit script, after trimming the
// common prefix and suffix; undefined when more than MAX_EDITS are needed.
function matchTokens(a: string[], b: string[]) {
  let head = 0
  while (head < a.length && head < b.length && a[head] === b[head]) head++
  let tail = 0
  while (tail < a.length - head && tail < b.length - head && a[a.length - 1 - tail] === b[b.length - 1 - tail]) tail++
  const middle = shortestEdit(a.slice(head, a.length - tail), b.slice(head, b.length - tail))
  if (!middle) return undefined
  return [
    ...Array.from({ length: head }, (_, k) => [k, k] as [number, number]),
    ...middle.map(([i, j]) => [i + head, j + head] as [number, number]),
    ...Array.from({ length: tail }, (_, k) => [a.length - tail + k, b.length - tail + k] as [number, number]),
  ]
}

// Myers' O(ND) diff; trace[d] keeps diagonals -d-1..d+1 from before round d.
function shortestEdit(a: string[], b: string[]) {
  const offset = a.length + b.length + 1
  const v = new Int32Array(2 * offset + 1)
  const trace: Int32Array[] = []
  for (let d = 0; d <= Math.min(a.length + b.length, MAX_EDITS); d++) {
    trace.push(v.slice(offset - d - 1, offset + d + 2))
    for (let k = -d; k <= d; k += 2) {
      let x = k === -d || (k !== d && v[offset + k - 1]! < v[offset + k + 1]!) ? v[offset + k + 1]! : v[offset + k - 1]! + 1
      let y = x - k
      while (x < a.length && y < b.length && a[x] === b[y]) {
        x++
        y++
      }
      v[offset + k] = x
      if (x >= a.length && y >= b.length) return backtrack(trace, a.length, b.length)
    }
  }
  return undefined
}

function backtrack(trace: Int32Array[], n: number, m: number) {
  const pairs: [number, number][] = []
  let x = n
  let y = m
  for (let d = trace.length - 1; d >= 0; d--) {
    const at = (k: number) => trace[d]![k + d + 1]!
    const k = x - y
    const previous = k === -d || (k !== d && at(k - 1) < at(k + 1)) ? k + 1 : k - 1
    const px = at(previous)
    const py = px - previous
    while (x > px && y > py) {
      x--
      y--
      pairs.push([x, y])
    }
    x = px
    y = py
  }
  return pairs.reverse()
}
