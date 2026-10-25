import { copyFileSync, existsSync, readFileSync, rmSync } from "node:fs"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { comparisonBaseline } from "./comparison"
import { compareDocxWithPointers, paragraphPointers, withStableRevisionDates } from "../../lib/drafting-amendment"
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
import { numberDocxClauses } from "../../lib/docx-numbering"
import { linkCrossReferences, referenceFields } from "../../lib/docx-crossrefs"
import { paragraphFields, refInstruction, unlink } from "../../lib/docx-fields"
import { shadeOpenItemsInDocx, unshadeOpenItems } from "../../lib/open-items"
import { assertRedlineRoundTrip } from "../../lib/redline-round-trip"
import { discloseSourceRevisions } from "../../lib/docx-revision-views"
import { openItem } from "./quotes"

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
  const result = await applyAmendments({ docxPath: body, outPath: out, amendments: changes.amendments })
    .then((value): ApplyResult => value, (error: unknown) => ({ error: message(directory, error) }))
  const settled = "error" in result ? settleApply(changes, undefined, false, result.error) : settleApply(changes, result, existsSync(out))
  if (settled.bodyUnchanged) copyFileSync(body, out)
  const identity = draftingAttribution()
  // A changed paragraph whose text now carries a clause number the body's list
  // supplies is numbered by the list, as at drafting. The later parts of a
  // split paragraph leave its list unless they open with its next number at
  // the paragraph's own level or a deeper one. Cross-references then show
  // their targets' new numbers; one whose target the changes removed becomes
  // an open-item marker, checked below like any other. A paragraph the changes
  // rewrote is linked again from its new text, as drafting links a document's
  // references (none in a document kept to typed numbers, advice or a letter),
  // so a number a change wrote or extended points at the clause it names; the
  // other paragraphs are only refreshed.
  const later = new Set(settled.result.applied.flatMap(row => row.result_pids?.slice(1) ?? []))
  const split = (paragraph: Element) => paragraphPointers(paragraph).some(pointer => later.has(pointer.pid))
  const renumbered = numberDocxClauses(readFileSync(out), { house: false, split })
  const rewritten = new Set(settled.result.applied.flatMap(row => row.mode === "code" ? row.result_pids ?? [row.pid] : []))
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
  const annotated = await documentRedline(directory, document, readFileSync(out), changes, settled.result.applied, identity)
  // Tracked changes are dated when this step started applying them; a change
  // already in the previous redline keeps its date, so an unchanged redline
  // keeps its bytes and the evidence keyed on them.
  const previous = leanFile(directory, `${document}/redline.docx`)
  writeLeanBytes(directory, `${document}/redline.docx`, withStableRevisionDates(annotated.bytes, state.stages[step]?.started_at ?? new Date().toISOString(),
    existsSync(previous) ? readFileSync(previous) : undefined))
  writeLeanText(directory, `${document}/apply.json`, `${JSON.stringify({ ...settled.result, comments: annotated.comments }, null, 2)}\n`)
  return {
    status: "done", outputs: stepOutputs(step),
    openItems: [
      ...settled.openItems, ...annotated.openItems, ...finalBodyItems(directory, rows), ...draftStatusItems(directory, document, rows, "apply"),
      ...templateItems(directory, document, readParagraphs(directory, document), rows, changes.amendments, new Set(settled.result.applied.map(row => row.id))),
      ...consistencyItems(directory, document, readParagraphs(directory, document), rows, referenceFields(readFileSync(out))),
      ...tableItems(directory, document, readFileSync(out), rows),
    ],
  }
}

/** A document's redline: its comparison baseline (comparisonBaseline) compared
 * with `final`, each applied change's explanation attached as a Word comment.
 * The apply step passes its own outcome. A change the lawyer accepts into the
 * delivered body later rebuilds the delivered redline from the outcome the step
 * recorded (changes.json and apply.json). */
export async function documentRedline(directory: string, document: DocID, final: Uint8Array,
  changes: ChangeList | undefined = readChangeList(directory, document),
  applied: ApplyResult["applied"] | undefined = recordedApplied(directory, document), identity = draftingAttribution()) {
  if (!changes || !applied) throw new Error(`${document}: the change list or the apply outcome is missing or unreadable.`)
  const source = comparisonBaseline(directory, document).bytes
  const [before, after] = comparableBodies(source, final)
  // The comparison splits an open item a change rewrote word by word into
  // several runs; the redline shades it whole, as the body shows it.
  const compared = await compareDocxWithPointers(before, after, { authorName: identity.author })
  const redline = withGeneratedDocxAttribution(shadeOpenItemsInDocx(compared), identity)
  const exported = exportAppliedChangeComments(redline, final, document, changes.amendments, applied, changes.explanations, readParagraphs(directory, document), identity.author)
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

function message(directory: string, error: unknown) {
  return matterRelative(directory, error instanceof Error ? error.message : String(error))
}
