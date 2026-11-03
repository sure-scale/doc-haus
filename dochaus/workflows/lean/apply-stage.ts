import { copyFileSync, existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import { join } from "node:path"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { comparisonBaseline } from "./comparison"
import { compareDocxWithPointers, paragraphPointers, withStableRevisionDates } from "../../lib/drafting-amendment"
import { amendmentParagraphs } from "../../lib/drafting-amendment"
import { applyAmendments } from "../../lib/drafting-amendment-apply"
import { finalBodyItems, readChangeList, readParagraphs, settleApply, type ApplyResult, type ChangeList } from "./checks"
import type { DocID } from "./plan"
import { draftedReferences, paragraphRows, stepOutputs } from "./stages"
import { leanFile, matterRelative, readLeanJson, writeLeanBytes, writeLeanText, type LeanState, type OpenItem } from "./state"
import { draftStatusItems } from "./draft-status"
import { templateItems } from "./template-coverage"
import { consistencyItems } from "./consistency"
import { tableItems } from "./original-tables"
import { exportAppliedChangeComments } from "./review-evidence"
import { draftingAttribution, withGeneratedDocxAttribution } from "../../lib/docx-attribution"
import { continueListStarts, numberDocxClauses, numberedList } from "../../lib/docx-numbering"
import { linkCrossReferences, referenceFields } from "../../lib/docx-crossrefs"
import { paragraphFields, refInstruction, unlink } from "../../lib/docx-fields"
import { shadeOpenItemsInDocx, unshadeOpenItems } from "../../lib/open-items"
import { assertRedlineRoundTrip } from "../../lib/redline-round-trip"
import { discloseSourceRevisions } from "../../lib/docx-revision-views"
import { openItem } from "./quotes"
import { normalize } from "../../lib/quote-verify"
import { paragraphSha256, type Amendment } from "../../lib/drafting-amendment"
import { LEAN_MAX_SUBSTANTIVE_REPAIRS, readRepairRound } from "./repairs"
import { paragraphShowsOpenItem } from "./gate"

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const BODY = "word/document.xml"

// The apply step of one document. Applies the document's change list to its
// drafted body, always from body.docx and never onto an earlier apply output,
// checks the final body for unsupported content and compares the two bodies
// into the document's redline. It writes only the reviewed change list's own
// text and opens no session: overlapping changes were consolidated before
// change-review.
//
// Every amendment ends applied or as a visible open item: an amendment the
// applier could not apply (an overlap included), did not report, or lost
// because the applier itself stopped is unapplied and carries the applier's
// detail.

export type ApplyStageRun =
  | { status: "done"; outputs: string[]; openItems: OpenItem[] }
  | { status: "failed"; reason: string }

export async function runApply(directory: string, state: LeanState, document: DocID): Promise<ApplyStageRun> {
  const step = `${document}.apply` as const
  const changes = readChangeList(directory, document)
  if (!changes) return { status: "failed", reason: `${step}: ${document}/changes.json is missing or unreadable.` }
  const body = leanFile(directory, `${document}/body.docx`)
  const out = leanFile(directory, `${document}/body-applied.docx`)
  // An earlier round's output must not stand in for this round's.
  rmSync(out, { force: true })
  const result = await applyAmendments({ docxPath: body, outPath: out, amendments: changes.amendments, ownedMarkerDecoration: true })
    .then((value): ApplyResult => value, (error: unknown) => ({ error: message(directory, error) }))
  const settled = "error" in result ? settleApply(changes, undefined, false, result.error) : settleApply(changes, result, existsSync(out))
  if (settled.bodyUnchanged) copyFileSync(body, out)
  const marked = await markUnverified(directory, out)
  const identity = draftingAttribution()
  // A changed paragraph whose text now carries a clause number the body's list
  // supplies is numbered by the list, as at drafting. The later parts of a
  // split paragraph leave its list unless they open with its next number at
  // the paragraph's own level or a deeper one. A run of the list that counted
  // on from the one before it in the drafted body still does, so the numbers
  // after an added or removed item follow on. Cross-references then show
  // their targets' new numbers; one whose target the changes removed becomes
  // an open-item marker, checked below like any other. A paragraph the changes
  // rewrote is linked again from its new text, as drafting links a document's
  // references (none in a document kept to typed numbers, advice or a letter),
  // so a number a change wrote or extended points at the clause it names; the
  // other paragraphs are only refreshed.
  const renumbered = renumberApplied(readFileSync(body), readFileSync(out), settled.result)
  const rewritten = new Set([...settled.result.applied.flatMap(row => row.mode === "code" ? row.result_pids ?? [row.pid] : []), ...marked.map(row => row.pid)])
  const relink = (paragraph: Element) => paragraphPointers(paragraph).some(pointer => rewritten.has(pointer.pid))
  const links = draftedReferences(directory, document).links
  const linked = await linkCrossReferences(renumbered, links ? { ...links, relink } : { refreshOnly: true })
  // The body's open items are shaded once its text is final, each marker
  // whole however the changes and the refresh split its runs; in a paragraph a
  // change rewrote, nothing else keeps an open item's shading. The other
  // paragraphs keep their own formatting.
  writeLeanBytes(directory, `${document}/body-applied.docx`, withGeneratedDocxAttribution(shadeOpenItemsInDocx(linked, relink), identity))
  const rows = await paragraphRows(out)
  writeLeanText(directory, `${document}/paragraphs-applied.json`, `${JSON.stringify(rows, null, 2)}\n`)
  const annotated = await documentRedline(directory, document, readFileSync(out), changes, settled.result.applied, identity, marked)
  // Tracked changes are dated when this step started applying them; a change
  // already in the previous redline keeps its date, so an unchanged redline
  // keeps its bytes and the evidence keyed on them.
  const previous = leanFile(directory, `${document}/redline.docx`)
  writeLeanBytes(directory, `${document}/redline.docx`, withStableRevisionDates(annotated.bytes, state.stages[step]?.started_at ?? new Date().toISOString(),
    existsSync(previous) ? readFileSync(previous) : undefined))
  writeLeanText(directory, `${document}/apply.json`, `${JSON.stringify({ ...settled.result, ...(marked.length ? { marked } : {}), comments: annotated.comments }, null, 2)}\n`)
  // A marker this step wrote shows its citation or quotation item, which the
  // notes name; it is not a second open item.
  const written = new Set(marked.flatMap(row => row.markers.map(marker => normalize(marker).toLowerCase())))
  const bodyItems = finalBodyItems(directory, rows).filter(item => item.kind !== "placeholder" || !written.has(normalize(item.ref.placeholder ?? "").toLowerCase()))
  return {
    status: "done", outputs: stepOutputs(step),
    openItems: [
      ...settled.openItems, ...annotated.openItems, ...bodyItems, ...draftStatusItems(directory, document, rows, "apply"),
      ...templateItems(directory, document, readParagraphs(directory, document), rows, changes.amendments, new Set(settled.result.applied.map(row => row.id))),
      ...consistencyItems(directory, document, readParagraphs(directory, document), rows, referenceFields(readFileSync(out))),
      ...tableItems(directory, document, readFileSync(out), rows),
    ],
  }
}

type MarkedParagraph = { pid: string; markers: string[] }

/** Once the run's substantive repair rounds are used up, no later round can
 * correct an unverified quotation or an unsupported citation left in the final
 * body, so each is marked where it stands: its paragraph gets a visible open-item
 * marker, written before the body is numbered, shaded and compared, and the
 * item itself, named in the drafting notes, goes with a review draft. While
 * rounds remain the gate blocks on it and Continue starts a repair round. */
async function markUnverified(directory: string, out: string): Promise<MarkedParagraph[]> {
  if ((readRepairRound(directory)?.substantive_attempts ?? 0) < LEAN_MAX_SUBSTANTIVE_REPAIRS || !existsSync(out)) return []
  const rows = await paragraphRows(out)
  const texts = new Map(rows.map(row => [row.pid, row.text]))
  const wanted = new Map<string, string[]>()
  for (const item of finalBodyItems(directory, rows)) {
    // A paragraph that already shows an open item satisfies the gate as it stands.
    if (item.kind !== "citation_unsupported" && item.kind !== "quote_miss" || !item.ref.pid || !texts.has(item.ref.pid) || paragraphShowsOpenItem(texts.get(item.ref.pid)!)) continue
    const marker = unverifiedMarker(item)
    const markers = wanted.get(item.ref.pid) ?? []
    if (!markers.includes(marker)) wanted.set(item.ref.pid, [...markers, marker])
  }
  const amendments: Amendment[] = [...wanted].map(([pid, markers], index) => ({ id: `m${index + 1}`, pid, existing_sha256: paragraphSha256(texts.get(pid)!),
    updated: `${texts.get(pid)!.trimEnd()} ${markers.join(" ")}`, reason: "Marked as an open item: it could not be verified against a source after the repair rounds were used up.", sources: [] }))
  if (!amendments.length) return []
  // A paragraph the applier could not mark keeps its item unmarked, and the gate blocks on it.
  const result = await applyAmendments({ docxPath: out, outPath: out, amendments, ownedMarkerDecoration: true }).catch(() => undefined)
  const applied = new Set(result?.applied.map(row => row.pid) ?? [])
  return [...wanted].flatMap(([pid, markers]) => applied.has(pid) ? [{ pid, markers }] : [])
}

/** The marker for an unverified quotation or citation. A citation is named
 * unless its text would end the marker early; a quotation is never repeated. */
function unverifiedMarker(item: OpenItem) {
  const citation = item.ref.citation?.replace(/\s+/g, " ").trim()
  if (item.kind === "quote_miss") return "[open item: quotation not verified against its source]"
  return citation && !/[\[\]()]/.test(citation) ? `[open item: citation not verified against a source: ${citation}]` : "[open item: citation not verified against a source]"
}

/** Applied amendments as the final body shows them, with any marker the apply
 * step added to their paragraphs (markUnverified), so their comments still
 * match the wording they explain. */
function withMarkers(amendments: Amendment[], applied: ApplyResult["applied"], marked: MarkedParagraph[]) {
  if (!marked.length) return amendments
  const markers = new Map(marked.map(row => [row.pid, row.markers]))
  return amendments.map(amendment => {
    const row = applied.find(outcome => outcome.id === amendment.id && outcome.mode === "code")
    const pids = row?.result_pids ?? (row ? [row.pid] : [])
    if (!pids.some(pid => markers.has(pid)) || amendment.operation === "delete") return amendment
    const parts = amendmentParagraphs(amendment).map((text, index) => markers.has(pids[index] ?? "") ? `${text.trimEnd()} ${markers.get(pids[index]!)!.join(" ")}` : text)
    return { ...amendment, updated: typeof amendment.updated === "string" ? parts[0]! : parts }
  })
}

// The applied body numbered by the drafted body's list (see runApply).
function renumberApplied(body: Uint8Array, applied: Uint8Array, result: ApplyResult) {
  const later = new Set(result.applied.flatMap(row => row.result_pids?.slice(1) ?? []))
  const split = (paragraph: Element) => paragraphPointers(paragraph).some(pointer => later.has(pointer.pid))
  const list = numberedList(applied)
  return continueListStarts(body, numberDocxClauses(applied, { house: false, split, ...(list ? { list } : {}) }))
}

/** The numbers the projected body change-review reads will show once applied:
 * the change list applied to a scratch copy of the drafted body and numbered
 * as apply numbers it. Saved as <document>/paragraphs-projected.json, one row
 * per projected paragraph (its initial pid, its text as the change list
 * gives it and the label it then shows), which the projection takes only
 * while its own text matches. Nothing is saved when a change would not apply:
 * the projection then shows the drafted numbers. */
export async function writeProjectedNumbers(directory: string, document: DocID) {
  const file = `${document}/paragraphs-projected.json`
  const changes = readChangeList(directory, document)
  const body = leanFile(directory, `${document}/body.docx`)
  rmSync(leanFile(directory, file), { force: true })
  if (!changes || !existsSync(body)) return
  const scratch = mkdtempSync(join(tmpdir(), "dochaus-projected-"))
  try {
    const out = join(scratch, "body.docx")
    const result = await applyAmendments({ docxPath: body, outPath: out, amendments: changes.amendments }).catch(() => undefined)
    const settled = settleApply(changes, result, existsSync(out))
    if (settled.result.unapplied.length) return
    writeFileSync(out, renumberApplied(readFileSync(body), readFileSync(out), settled.result))
    const labels = new Map((await paragraphRows(out)).map(row => [row.pid, row.label]))
    const outcomes = new Map(settled.result.applied.map(row => [row.pid, row]))
    const amendments = new Map(changes.amendments.map(row => [row.pid, amendmentParagraphs(row)]))
    const rows = readParagraphs(directory, document).flatMap(row => (amendments.get(row.pid) ?? [row.text]).map((text, at) => {
      const label = labels.get(outcomes.get(row.pid)?.result_pids?.[at] ?? row.pid)
      return { pid: row.pid, text, ...(label ? { label } : {}) }
    }))
    writeLeanText(directory, file, `${JSON.stringify(rows, null, 2)}\n`)
  } finally {
    rmSync(scratch, { recursive: true, force: true })
  }
}

/** A document's redline: its comparison baseline (comparisonBaseline) compared
 * with `final`, each applied change's explanation attached as a Word comment.
 * The apply step passes its own outcome. A change the lawyer accepts into the
 * delivered body later rebuilds the delivered redline from the outcome the step
 * recorded (changes.json and apply.json). */
export async function documentRedline(directory: string, document: DocID, final: Uint8Array,
  changes: ChangeList | undefined = readChangeList(directory, document),
  applied: ApplyResult["applied"] | undefined = recordedApplied(directory, document), identity = draftingAttribution(),
  marked: MarkedParagraph[] = recordedMarked(directory, document)) {
  if (!changes || !applied) throw new Error(`${document}: the change list or the apply outcome is missing or unreadable.`)
  const source = comparisonBaseline(directory, document).bytes
  const [before, after] = comparableBodies(source, final)
  // The comparison splits an open item a change rewrote word by word into
  // several runs; the redline shades it whole, as the body shows it.
  const compared = await compareDocxWithPointers(before, after, { authorName: identity.author })
  const redline = withGeneratedDocxAttribution(shadeOpenItemsInDocx(compared), identity)
  const exported = exportAppliedChangeComments(redline, final, document, withMarkers(changes.amendments, applied, marked), applied, changes.explanations, readParagraphs(directory, document), identity.author)
  const disclosed = await discloseSourceRevisions(exported.bytes, source, identity.author)
  await assertRedlineRoundTrip(disclosed.bytes, before, after)
  return { ...exported, bytes: disclosed.bytes, openItems: [...exported.openItems,
    ...(disclosed.disclosure ? [openItem("apply", "omission", disclosed.disclosure.message, { document })] : [])] }
}

/** The two bodies as the redline compares them. The comparer marks a whole
 * paragraph deleted and inserted again when its cross-reference fields differ
 * (one linked, unlinked, pointed at another clause or showing a new number),
 * so a paragraph whose fields differ between the two, by paragraph pointer, is
 * the text it shows on both sides, and so is every paragraph when the baseline
 * is a supplied original: the redline marks only the words that changed, and
 * its references there are plain text. Open items are compared unshaded, so a
 * word kept inside a new open item ("2.1" in "[open item: refers to 2.1, which
 * was removed]") is no formatting change; the redline shades its open items
 * afterwards. The body itself keeps its fields and shading. Each body is
 * returned itself when nothing in it is unlinked or unshaded. */
function comparableBodies(baseline: Uint8Array, final: Uint8Array): [Uint8Array, Uint8Array] {
  const read = (bytes: Uint8Array) => {
    const parts = unzipSync(bytes)
    const doc = new DOMParser().parseFromString(strFromU8(parts[BODY]!), "application/xml")
    const paragraphs = new Map<string, Element>()
    for (const paragraph of Array.from(doc.getElementsByTagNameNS(W, "p"))) {
      const pid = paragraphPointers(paragraph)[0]?.pid
      if (pid && !paragraphs.has(pid)) paragraphs.set(pid, paragraph)
    }
    return { bytes, parts, doc, paragraphs, changed: unshadeOpenItems(doc) }
  }
  const before = read(baseline)
  const after = read(final)
  const references = (paragraph: Element) => paragraphFields(paragraph).filter(field => refInstruction(field))
  const shown = (paragraph: Element) => JSON.stringify(references(paragraph).map(field => [field.instruction.trim(), field.result]))
  const unlinked = (body: typeof before, paragraphs: Iterable<Element>) => {
    for (const paragraph of paragraphs) for (const field of references(paragraph)) {
      unlink(field)
      body.changed = true
    }
  }
  // A supplied original has no paragraph pointers, so none of its paragraphs
  // pairs; its references, typed or its own fields, are compared as text with
  // the text the body's show.
  if (!before.paragraphs.size) for (const body of [before, after]) unlinked(body, Array.from(body.doc.getElementsByTagNameNS(W, "p")))
  else for (const [pid, paragraph] of after.paragraphs) {
    const earlier = before.paragraphs.get(pid)
    if (!earlier || shown(earlier) === shown(paragraph)) continue
    unlinked(before, [earlier])
    unlinked(after, [paragraph])
  }
  const serialized = (body: typeof before) => body.changed
    ? zipSync({ ...body.parts, [BODY]: strToU8(new XMLSerializer().serializeToString(body.doc)) }, { mtime: new Date(1980, 0, 1) }) : body.bytes
  return [serialized(before), serialized(after)]
}

function recordedApplied(directory: string, document: DocID) {
  const applied = (readLeanJson(directory, `${document}/apply.json`) as Partial<ApplyResult> | undefined)?.applied
  return Array.isArray(applied) ? applied : undefined
}

function recordedMarked(directory: string, document: DocID): MarkedParagraph[] {
  const marked = (readLeanJson(directory, `${document}/apply.json`) as { marked?: unknown } | undefined)?.marked
  return Array.isArray(marked) ? marked as MarkedParagraph[] : []
}

function message(directory: string, error: unknown) {
  return matterRelative(directory, error instanceof Error ? error.message : String(error))
}
