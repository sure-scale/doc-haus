import { expect, setSystemTime, test } from "bun:test"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { createDocxLocalization } from "./docx-amendments"
import { exportDocxComments, preserveDocxComments, readDocxCommentParagraphs, readDocxComments } from "./docx-comments"
import { compareDocxWithPointers, withStableRevisionDates } from "./drafting-amendment"
import { renderMarkdownToDocx } from "./render"
import type { RedlineRow } from "../../services/ingest/src/db"
import { bake, buildRedlined, RedlinePreviewConflictError } from "../../services/ingest/src/redline"

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const author = "Reviewing Lawyer & 李"
const row = (id: number, find: string, replacement: string): RedlineRow => ({ id,
  doc_path: "source.docx", doc_name: "source.docx", scope: "phrase", find_text: find, old_text: find,
  new_text: replacement, author, review_comment: `Reason for proposal ${id}.\nSource: supplied instruction & review.`,
  anchor_id: null, status: "pending", created_at: id })
const document = async (...text: string[]) => (await renderMarkdownToDocx(text.join("\n\n"))).bytes
const xml = (bytes: Uint8Array, part: string) => new DOMParser().parseFromString(strFromU8(unzipSync(bytes)[part]), "application/xml")
const serialized = (bytes: Uint8Array, part: string, name: string) => Array.from(xml(bytes, part).getElementsByTagNameNS(W, name))
  .map(node => new XMLSerializer().serializeToString(node))

test("legacy redline exports native notes on final paragraphs without changing revisions, embedded comments or inputs", async () => {
  const original = exportDocxComments(await document("The notice period is five days.", "Keep the retained record."), [
    { id: "prior", text: "Original embedded note.", author: "Original Commenter", anchor: { quote: "retained record" } },
  ]).bytes
  const source = new Uint8Array(original)
  const rows = [row(1, "five", "seven"), row(2, "notice", "review")]
  const stored = structuredClone(rows)
  const baseline = await buildRedlined(original, rows.map(row => ({ ...row, review_comment: "" })))
  const output = await buildRedlined(original, rows)
  for (const kind of ["ins", "del"]) {
    expect(serialized(baseline, "word/document.xml", kind).length).toBeGreaterThan(0)
    // Comparison timestamps can differ between invocations; annotation preserves
    // the complete revision XML apart from the comparator's generated date.
    expect(serialized(output, "word/document.xml", kind).map(text => text.replace(/ w:date="[^"]*"/g, "")))
      .toEqual(serialized(baseline, "word/document.xml", kind).map(text => text.replace(/ w:date="[^"]*"/g, "")))
    expect(Array.from(xml(output, "word/document.xml").getElementsByTagNameNS(W, kind)).every(node => node.getAttributeNS(W, "author") === author)).toBe(true)
  }
  expect(serialized(output, "word/comments.xml", "comment")[0]).toBe(serialized(original, "word/comments.xml", "comment")[0])
  const comments = readDocxComments(output)
  expect(comments.map(comment => comment.text)).toEqual(["Original embedded note.", ...rows.map(row => row.review_comment)])
  expect(comments[0].markers).toEqual(readDocxComments(original)[0].markers)
  expect(comments.slice(1).every(comment => comment.markers.length === 3
    && comment.markers.every(marker => marker.accepted_paragraph_text === "The review period is seven days."))).toBe(true)
  expect(Array.from(xml(output, "word/comments.xml").getElementsByTagNameNS(W, "comment")).map(node => node.getAttributeNS(W, "author")))
    .toEqual(["Original Commenter", author, author])
  expect(readDocxCommentParagraphs(output).map(row => row.paragraph_text)).toEqual(["The review period is seven days.", "Keep the retained record."])
  expect(original).toEqual(source)
  expect(rows).toEqual(stored)
})

test("original comments on edited ranges refuse export instead of moving or dropping the original annotation", async () => {
  const original = exportDocxComments(await document("The notice period is five days.", "Keep the retained record."), [
    { id: "prior", text: "Check this exact range.", author: "Original Commenter", anchor: { quote: "five days" } },
  ]).bytes
  const before = new Uint8Array(original)
  await expect(buildRedlined(original, [row(1, "five", "seven")])).rejects.toThrow("original unchanged anchor")
  expect(original).toEqual(before)
})

test("restoring a commented hyperlink rejects coherent comparison relationship reallocation instead of retargeting the link", async () => {
  const R = "http://schemas.openxmlformats.org/package/2006/relationships"
  const O = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  const parts = unzipSync(await document("Notice is five days.", "Read the retained record."))
  const body = xml(zipSync(parts), "word/document.xml")
  const paragraph = body.getElementsByTagNameNS(W, "p")[1]
  const link = body.createElementNS(W, "w:hyperlink")
  link.setAttributeNS(O, "r:id", "rIdOriginalLink")
  const run = paragraph.getElementsByTagNameNS(W, "r")[0]
  paragraph.insertBefore(link, run)
  link.appendChild(run)
  const relationships = xml(zipSync(parts), "word/_rels/document.xml.rels")
  const relationship = relationships.createElementNS(R, "Relationship")
  relationship.setAttribute("Id", "rIdOriginalLink")
  relationship.setAttribute("Type", `${O}/hyperlink`)
  relationship.setAttribute("Target", "https://example.test/retained-record")
  relationship.setAttribute("TargetMode", "External")
  relationships.documentElement.appendChild(relationship)
  parts["word/document.xml"] = strToU8(new XMLSerializer().serializeToString(body))
  parts["word/_rels/document.xml.rels"] = strToU8(new XMLSerializer().serializeToString(relationships))
  const original = exportDocxComments(zipSync(parts), [{ id: "link", text: "Review this linked record.", author: "Prior reviewer", anchor: { quote: "retained record" } }]).bytes
  const modified = await bake(original, [row(1, "five", "seven")])
  const compared = unzipSync(await compareDocxWithPointers(original, modified, { authorName: author }))
  const comparedBody = xml(zipSync(compared), "word/document.xml")
  const comparedRelationships = xml(zipSync(compared), "word/_rels/document.xml.rels")
  // This comparator removes the wrapper but keeps its relationship. Restoration
  // is safe while that mapping is unchanged; also prove a missing map refuses.
  expect(readDocxComments(preserveDocxComments(original, zipSync(compared)))[0].text).toBe("Review this linked record.")
  const comparedRelationship = Array.from(comparedRelationships.getElementsByTagNameNS(R, "Relationship"))
    .find(node => node.getAttribute("Id") === "rIdOriginalLink")!
  comparedRelationship.parentNode!.removeChild(comparedRelationship)
  compared["word/_rels/document.xml.rels"] = strToU8(new XMLSerializer().serializeToString(comparedRelationships))
  expect(() => preserveDocxComments(original, zipSync(compared))).toThrow("original_comment_paragraph_dependency_changed")
  comparedRelationships.documentElement.appendChild(comparedRelationship)
  // Reintroduce a coherent hyperlink in the comparison package, then reallocate
  // its ID below while keeping the comparison's own link correct.
  const comparedParagraph = comparedBody.getElementsByTagNameNS(W, "p")[1]
  const comparedLink = comparedBody.createElementNS(W, "w:hyperlink")
  comparedLink.setAttributeNS(O, "r:id", "rIdOriginalLink")
  const comparedRun = comparedParagraph.getElementsByTagNameNS(W, "r")[0]
  comparedParagraph.insertBefore(comparedLink, comparedRun)
  comparedLink.appendChild(comparedRun)
  compared["word/document.xml"] = strToU8(new XMLSerializer().serializeToString(comparedBody))
  compared["word/_rels/document.xml.rels"] = strToU8(new XMLSerializer().serializeToString(comparedRelationships))
  const matching = zipSync(compared)
  const restored = preserveDocxComments(original, matching)
  expect(readDocxComments(restored)[0].text).toBe("Review this linked record.")
  const restoredID = xml(restored, "word/document.xml").getElementsByTagNameNS(W, "hyperlink")[0].getAttributeNS(O, "id")
  expect(Array.from(xml(restored, "word/_rels/document.xml.rels").getElementsByTagNameNS(R, "Relationship"))
    .find(node => node.getAttribute("Id") === restoredID)?.getAttribute("Target")).toBe("https://example.test/retained-record")
  // The comparison remains internally valid: its link uses the newly allocated
  // ID. Reusing the original paragraph would instead point at the other target.
  comparedRelationship.setAttribute("Id", "rIdComparisonLink")
  comparedLink.setAttributeNS(O, "r:id", "rIdComparisonLink")
  const other = relationship.cloneNode(true) as Element
  other.setAttribute("Target", "https://example.test/unrelated-record")
  comparedRelationships.documentElement.appendChild(other)
  compared["word/document.xml"] = strToU8(new XMLSerializer().serializeToString(comparedBody))
  compared["word/_rels/document.xml.rels"] = strToU8(new XMLSerializer().serializeToString(comparedRelationships))
  const reallocated = zipSync(compared)
  const before = new Uint8Array(reallocated)
  expect(() => preserveDocxComments(original, reallocated)).toThrow("original_comment_paragraph_dependency_changed")
  expect(reallocated).toEqual(before)
})

test("revision-bound notes follow a selected duplicate paragraph and preserve stored attribution", async () => {
  const original = await document("Notice is five days.", "Notice is five days.")
  const proposal = row(1, "five days", "seven days")
  const packet = await createDocxLocalization(original, [{ id: "1", before_quote: proposal.find_text, replacement: proposal.new_text }])
  proposal.target_json = JSON.stringify({ version: 1, source_sha256: packet.revision.source_sha256,
    target_handle: packet.candidates[1].target_handle, amendment_id: "1" })
  const output = await buildRedlined(original, [proposal])
  expect(readDocxComments(output)[0]).toMatchObject({ text: proposal.review_comment,
    markers: Array.from({ length: 3 }, () => expect.objectContaining({ paragraph_index: 1, accepted_paragraph_text: "Notice is seven days." })) })
  expect(readDocxCommentParagraphs(output).map(row => row.paragraph_text)).toEqual(["Notice is five days.", "Notice is seven days."])
  const stale = new Uint8Array(await document("Notice is four days.", "Notice is five days."))
  await expect(buildRedlined(stale, [proposal])).rejects.toBeInstanceOf(RedlinePreviewConflictError)
  await expect(buildRedlined(original, [{ ...proposal, target_json: "{invalid" }])).rejects.toThrow("fresh change list")
})

test("annotated exports reject missing, ambiguous, structural and missing-author targets instead of losing notes", async () => {
  const original = await document("Notice is five days.")
  expect(readDocxComments(await buildRedlined(original, [row(1, "five", "seven")]))).toHaveLength(1)
  await expect(buildRedlined(original, [row(1, "missing", "seven")])).rejects.toThrow("missing or ambiguous")
  await expect(buildRedlined(await document("Notice is five days.", "Notice is five days."), [row(1, "five", "seven")])).rejects.toThrow("ambiguous")
  await expect(buildRedlined(original, [row(1, "five", "seven\nmore")])).rejects.toThrow("paragraph-preserving")
  await expect(buildRedlined(original, [{ ...row(1, "five", "seven"), author: "" }])).rejects.toThrow("stored author")
  await expect(buildRedlined(original, [row(1, "Notice is five days.", "")])).rejects.toThrow("exact changed paragraph")
})

test("a regenerated comparison keeps first-made revision dates and exact bytes; only a new change takes the new time", async () => {
  const original = await document("The notice period is five days.", "Keep the retained record.")
  const revised = await document("The notice period is seven days.", "Keep the retained record.")
  const dates = (bytes: Uint8Array) => Array.from(xml(bytes, "word/document.xml").getElementsByTagNameNS(W, "p"))
    .map(paragraph => ["del", "ins"].flatMap(kind => Array.from(paragraph.getElementsByTagNameNS(W, kind)).map(node => node.getAttributeNS(W, "date"))))
  const withoutDates = (bytes: Uint8Array, kind: string) => serialized(bytes, "word/document.xml", kind).map(text => text.replace(/ w:date="[^"]*"/g, ""))
  try {
    setSystemTime(new Date("2026-03-04T05:06:07.890Z"))
    const raw = await compareDocxWithPointers(original, revised, { authorName: author })
    const first = withStableRevisionDates(raw, new Date().toISOString())
    expect(dates(first)).toEqual([["2026-03-04T05:06:07Z", "2026-03-04T05:06:07Z"], []])
    for (const kind of ["ins", "del"]) expect(withoutDates(first, kind)).toEqual(withoutDates(raw, kind))
    // The comparator dates and ids differ on every run; the export does not.
    setSystemTime(new Date("2026-05-06T07:08:09Z"))
    const again = withStableRevisionDates(await compareDocxWithPointers(original, revised, { authorName: author }), new Date().toISOString(), first)
    expect(again).toEqual(first)
    const later = await document("The notice period is seven days.", "Keep the signed record.")
    const changed = withStableRevisionDates(await compareDocxWithPointers(original, later, { authorName: author }), new Date().toISOString(), first)
    expect(dates(changed)).toEqual([["2026-03-04T05:06:07Z", "2026-03-04T05:06:07Z"], ["2026-05-06T07:08:09Z", "2026-05-06T07:08:09Z"]])
    expect(withStableRevisionDates(await compareDocxWithPointers(original, later, { authorName: author }), new Date().toISOString(), new Uint8Array([1, 2, 3])))
      .not.toEqual(changed)
  } finally {
    setSystemTime()
  }
})
