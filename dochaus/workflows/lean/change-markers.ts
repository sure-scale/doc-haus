import { amendmentParagraphs, amendmentText, normalizeParagraphText, type Amendment } from "../../lib/drafting-amendment"
import { openItemMarkers } from "../../lib/open-items"
import { draftingPlaceholders, stripDraftingPlaceholders } from "../../lib/placeholder-resolution"
import { paragraphShowsOpenItem } from "./gate"


/** The body must show the lawyer what an amendment did. A change resting on
 * drafting judgment or a view of the law has no passage to quote, so it stands
 * unsourced only with an open-item marker of its own: a marker, ___ field or
 * placeholder already in the paragraph points at an earlier gap, not at this
 * change. A change of markers alone is not a drafting change: on a paragraph that
 * already shows an open item it adds nothing, and elsewhere it belongs with a
 * question about that paragraph. Nor may markers replace the wording: that is a
 * deletion the body would not show as one. */
export function markerIssues(amendment: Amendment, index: number, current: string, questionPids: Set<string>): string[] {
  const at = `amendments[${index}] ([${amendment.pid}])`
  if (amendment.operation === "delete") return amendment.sources.length ? [] : [
    `${at} is an explicit deletion without source support. Quote the source or settled instruction that supports removing this paragraph; otherwise retain its wording with a new visible [open item: confirm ...] marker.`,
  ]
  if (amendment.operation === "split") return amendmentParagraphs(amendment).flatMap((text, part) =>
    markerIssues({ ...amendment, operation: "replace", updated: text }, index, current, questionPids))
  const updated = amendmentText(amendment)
  const before = markerSet(current)
  const after = markerSet(updated)
  const added = [...after].filter(marker => !before.has(marker) && /^[[(]open item:/.test(marker))
  // Markers the body would not show: the applier writes `updated` as given, and
  // the final body's open items and highlights read the raw text.
  const shown = new Set(openItemMarkers(updated.normalize("NFC")).map(marker => marker.toLowerCase()))
  const unreadable = openItemMarkers(normalizeParagraphText(updated)).filter(marker => !shown.has(marker.toLowerCase()))
  const empty = added.filter(marker => !marker.replace(/^[[(]open item:|[\])]$/g, "").replace(/\bconfirm\b/i, "").replace(/[\s.,;:\u2026-]+/g, ""))
  const markerOnly = withoutMarkers(updated) === withoutMarkers(current) && [...before].every(marker => after.has(marker))
  // Wording replaced by markers alone would delete the paragraph without the
  // redline or the drafting notes showing a deletion.
  const emptied = !withoutMarkers(updated) && Boolean(withoutMarkers(current))
  return [
    ...(emptied ? [`${at} replaces the paragraph's wording with markers only, which would delete it without showing a deletion. ` +
      "Keep the wording and add an [open item: ...] marker saying why the lawyer should consider removing it."] : []),
    ...unreadable.map(marker => `${at} writes the marker "${marker}" with spacing that will not show as an open item. ` +
      "Write it as [open item: ...] with ordinary single spaces."),
    ...empty.map(marker => `${at} adds the marker "${marker}", which does not say what the lawyer is to confirm or supply. Name that point in the marker.`),
    ...(!amendment.sources.length && !added.length
      ? [`${at} has no source passage. Quote the passage of a matter document that supports the change itself; where the change rests on drafting judgment or a view of the law, add a new visible [open item: confirm ...] marker saying what the lawyer is to confirm. A marker, ___ field or placeholder already in the paragraph does not count.`]
      : []),
    ...(markerOnly && paragraphShowsOpenItem(current)
      ? [`${at} only adds an open-item marker to a paragraph that already shows one. Drop it, or change the drafting.`] : []),
    ...(markerOnly && !paragraphShowsOpenItem(current) && !questionPids.has(amendment.pid)
      ? [`${at} changes only the paragraph's open-item markers. Change the drafting and keep the marker, or put the point in "questions" with the pid "${amendment.pid}".`]
      : []),
  ]
}

/** The markers the final body will show, read with the gate's extractor from
 * the text as the applier writes it. */
function markerSet(text: string) {
  return new Set([...openItemMarkers(text), ...draftingPlaceholders(text)].map(marker => marker.toLowerCase()))
}

/** The paragraph's wording apart from its markers, for comparing two versions. */
function withoutMarkers(text: string) {
  const normalized = normalizeParagraphText(text)
  return stripDraftingPlaceholders(openItemMarkers(normalized).reduce((rest, marker) => rest.replaceAll(marker, " "), normalized)).replace(/\s+/g, "")
}
