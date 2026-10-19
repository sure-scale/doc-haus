import { readFile } from "node:fs/promises"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { z } from "zod"
import { changedDocxSpan, openRevision } from "./docx-amendments"
import { amendmentSchema, amendmentParagraphs, amendmentText, ensureParagraphPointers, normalizeParagraphText, owningParagraph, paragraphPointers, paragraphSha256, type Amendment } from "./drafting-amendment"
import { writeDurableBytes } from "./durable-json"
import { FIELD_HIGHLIGHT_FILL, OPEN_ITEM_HIGHLIGHT_FILL, hasOpenItemMarker } from "./open-items"

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
// differently is left as it was (apply_failed). outPath is written atomically.
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
// least one paragraph, so this ends.
async function writeVerified(source: Uint8Array, opened: Opened, edits: Edit[]) {
  const failed: { edit: Edit; issue: string }[] = []
  for (let round = 0; ; round++) {
    const kept = edits.filter(edit => !failed.some(row => row.edit === edit))
    if (!kept.length) return { bytes: source, failed }
    const revision = round ? await openRevision(source) : opened
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
      if (round) revision.session.close()
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
      const change = changedDocxSpan(match, text)
      if (!change.match.text && !change.replacement) continue
      const result = opened.session.replaceMatch(change.match, change.replacement)
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
// highlighted open item would leave the answer shaded as open. Clears open-item
// shading from runs of edited paragraphs that no longer hold a marker.
function clearResolvedHighlights(bytes: Uint8Array, pids: Set<string>) {
  const parts = unzipSync(bytes)
  const doc = new DOMParser().parseFromString(strFromU8(parts[BODY]!), "application/xml")
  const shading = Array.from(doc.getElementsByTagNameNS(W, "shd")).filter(shd => {
    const run = shd.parentNode?.parentNode as Element | null
    if (run?.localName !== "r" || !HIGHLIGHT_FILLS.has((shd.getAttributeNS(W, "fill") ?? "").toUpperCase())) return false
    const paragraph = owningParagraph(run) as Element | null
    if (!paragraph || !paragraphPointers(paragraph).some(row => pids.has(row.pid))) return false
    return !hasOpenItemMarker(Array.from(run.getElementsByTagNameNS(W, "t")).map(text => text.textContent ?? "").join(""))
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
// Whitespace runs compare equal, so the document's own spacing is kept.
function hunkTexts(from: string, to: string) {
  const a = from.match(/\s+|\S+/g) ?? []
  const b = to.match(/\s+|\S+/g) ?? []
  const key = (token: string) => (/^\s/.test(token) ? " " : token)
  const pairs = matchTokens(a.map(key), b.map(key))
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
