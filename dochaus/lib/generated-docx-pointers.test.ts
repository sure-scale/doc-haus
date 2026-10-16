import { expect, test } from "bun:test"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises"
import { tmpdir } from "node:os"
import path from "node:path"
import { compareDocxWithPointers, ensureParagraphPointers, withParagraphPointers } from "./drafting-amendment"
import { docxodus } from "./docxodus"
import { exportDocxComments, readDocxCommentParagraphs, readDocxRevisionText } from "./docx-comments"
import { renderMarkdownToDocx } from "./render"

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

test("shared rendering gives drafting notes and table cells unique hidden paragraph pointers", async () => {
  const rendered = await renderMarkdownToDocx("# Drafting notes\n\nA remaining question is disclosed here.\n\n| Topic | Treatment |\n| --- | --- |\n| Timing | Confirm the date. |")
  const body = new DOMParser().parseFromString(strFromU8(unzipSync(rendered.bytes)["word/document.xml"]!), "application/xml")
  const paragraphs = Array.from(body.getElementsByTagNameNS(W, "p")).filter(node =>
    Array.from(node.getElementsByTagNameNS(W, "t")).some(text => text.textContent?.trim()))
  const ids = paragraphs.map(node => Array.from(node.getElementsByTagNameNS(W, "bookmarkStart"))
    .filter(mark => mark.getAttributeNS(W, "name")?.startsWith("_dh_")))
  expect(paragraphs.length).toBeGreaterThan(2)
  expect(ids.every(rows => rows.length === 1)).toBe(true)
  expect(new Set(ids.flat().map(mark => mark.getAttributeNS(W, "name"))).size).toBe(paragraphs.length)
  expect(body.getElementsByTagNameNS(W, "tc").length).toBe(4)
})

test("a saved template round trip and a second pointer pass retain ids and exact bytes", async () => {
  const source = (await renderMarkdownToDocx("# Template\n\n[insert recipient] shall receive the record.\n\n| Topic | Detail |\n| --- | --- |\n| Notice | [insert date] | ")).bytes
  const session = (await docxodus()).openDocxSession(source, {})
  let saved: Uint8Array
  try { saved = session.save() } finally { session.close() }
  expect(readDocxCommentParagraphs(saved).map(row => row.pids)).toEqual(readDocxCommentParagraphs(source).map(row => row.pids))
  expect(withParagraphPointers(saved)).toBe(saved)
  const directory = await mkdtemp(path.join(tmpdir(), "dh-generated-pointer-"))
  try {
    const file = path.join(directory, "template.docx")
    await writeFile(file, saved)
    const first = await ensureParagraphPointers(file)
    expect(await ensureParagraphPointers(file)).toEqual(first)
    expect(new Uint8Array(await readFile(file))).toEqual(saved)
  } finally { await rm(directory, { recursive: true, force: true }) }
})

test("comparison export restores clean-body ids without changing revisions or native comments", async () => {
  const before = (await renderMarkdownToDocx("# Terms\n\nThe record is due within five days.\n\n| Topic | Detail |\n| --- | --- |\n| Notice | Keep a copy. | ")).bytes
  const current = (await renderMarkdownToDocx("# Terms\n\nThe record is due within seven days.\n\n| Topic | Detail |\n| --- | --- |\n| Notice | Keep a copy. | ")).bytes
  const raw = await (await docxodus()).compareDocuments(before, current, { authorName: "Review author" })
  const annotated = exportDocxComments(raw, [{ id: "timing", text: "Explains the changed period.\nSource: supplied instructions.", author: "Reviewer", anchor: { quote: "seven days" } }]).bytes
  const revisions = readDocxRevisionText(annotated)
  expect(revisions.inserted_text.length).toBeGreaterThan(0)
  expect(revisions.deleted_text.length).toBeGreaterThan(0)
  const output = withParagraphPointers(annotated, current)
  expect(readDocxCommentParagraphs(output)).toEqual(readDocxCommentParagraphs(current))
  expect(readDocxRevisionText(output)).toEqual(revisions)
  const originalParts = unzipSync(annotated)
  for (const [name, bytes] of Object.entries(originalParts).filter(([name]) => name !== "word/document.xml"))
    expect(unzipSync(output)[name]).toEqual(bytes)
  const ranges = (bytes: Uint8Array) => {
    const doc = new DOMParser().parseFromString(strFromU8(unzipSync(bytes)["word/document.xml"]!), "application/xml")
    return ["commentRangeStart", "commentRangeEnd", "commentReference"].flatMap(kind =>
      Array.from(doc.getElementsByTagNameNS(W, kind)).map(node => new XMLSerializer().serializeToString(node)))
  }
  expect(ranges(output)).toEqual(ranges(annotated))
  expect(withParagraphPointers(output, current)).toBe(output)
  const exported = await compareDocxWithPointers(before, current, { authorName: "Review author" })
  const repeated = await compareDocxWithPointers(before, current, { authorName: "Review author" })
  expect(readDocxCommentParagraphs(exported)).toEqual(readDocxCommentParagraphs(current))
  expect(readDocxCommentParagraphs(repeated)).toEqual(readDocxCommentParagraphs(exported))
  const directory = await mkdtemp(path.join(tmpdir(), "dh-pointer-revisions-"))
  try {
    const file = path.join(directory, "redline.docx")
    await writeFile(file, output)
    await expect(ensureParagraphPointers(file)).rejects.toThrow("unsupported_existing_tracked_revisions")
    expect(new Uint8Array(await readFile(file))).toEqual(output)
  } finally { await rm(directory, { recursive: true, force: true }) }
})

test("copied paragraphs receive distinct ids and a changed sequence does not borrow unrelated reference ids", async () => {
  const original = (await renderMarkdownToDocx("First paragraph.\n\nSecond paragraph.")).bytes
  const parts = unzipSync(original)
  const doc = new DOMParser().parseFromString(strFromU8(parts["word/document.xml"]!), "application/xml")
  const first = doc.getElementsByTagNameNS(W, "p")[0]!
  first.parentNode!.insertBefore(first.cloneNode(true), first.nextSibling)
  const copied = zipSync({ ...parts, "word/document.xml": strToU8(new XMLSerializer().serializeToString(doc)) })
  const output = withParagraphPointers(copied)
  const rows = readDocxCommentParagraphs(output)
  expect(rows.map(row => row.paragraph_text)).toEqual(["First paragraph.", "First paragraph.", "Second paragraph."])
  expect(rows.every(row => row.pids.length === 1)).toBe(true)
  expect(new Set(rows.flatMap(row => row.pids)).size).toBe(3)
  expect(rows[0]!.pids).toEqual(readDocxCommentParagraphs(original)[0]!.pids)
  expect(withParagraphPointers(output)).toBe(output)
  const other = (await renderMarkdownToDocx("A different first paragraph.\n\nA different second paragraph.")).bytes
  expect(withParagraphPointers(original, other)).toBe(original)
  expect(readDocxCommentParagraphs(other).map(row => row.pids)).not.toEqual(readDocxCommentParagraphs(original).map(row => row.pids))
})

test("a historical moved or deleted bookmark cannot stand in for an accepted paragraph pointer", async () => {
  const rendered = (await renderMarkdownToDocx("Kept.")).bytes
  for (const kind of ["moveFrom", "del"] as const) {
    const parts = unzipSync(rendered)
    const doc = new DOMParser().parseFromString(strFromU8(parts["word/document.xml"]!), "application/xml")
    for (const name of ["bookmarkStart", "bookmarkEnd"])
      Array.from(doc.getElementsByTagNameNS(W, name)).forEach(node => node.parentNode!.removeChild(node))
    const paragraph = doc.getElementsByTagNameNS(W, "p")[0]!
    const historical = new DOMParser().parseFromString(`<w:${kind} xmlns:w="${W}" w:id="3"><w:bookmarkStart w:id="4" w:name="_dh_p000000000000"/><w:r><w:${kind === "del" ? "delText" : "t"}>Moved.</w:${kind === "del" ? "delText" : "t"}></w:r><w:bookmarkEnd w:id="4"/></w:${kind}>`, "application/xml").documentElement!
    paragraph.appendChild(doc.importNode(historical, true))
    const tracked = exportDocxComments(zipSync({ ...parts, "word/document.xml": strToU8(new XMLSerializer().serializeToString(doc)) }), [
      { id: "original", text: "Retain the current wording.", author: "Reviewer", anchor: { paragraph_text: "Kept." } },
    ]).bytes
    expect(readDocxCommentParagraphs(tracked)).toEqual([{ paragraph_index: 0, paragraph_text: "Kept.", pids: [] }])
    const revisions = readDocxRevisionText(tracked)
    const fixed = withParagraphPointers(tracked)
    const rows = readDocxCommentParagraphs(fixed)
    expect(rows).toHaveLength(1)
    expect(rows[0]!.pids).toHaveLength(1)
    expect(rows[0]!.pids[0]).not.toBe("p000000000000")
    expect(readDocxRevisionText(fixed)).toEqual(revisions)
    const body = new DOMParser().parseFromString(strFromU8(unzipSync(fixed)["word/document.xml"]!), "application/xml")
    expect(new XMLSerializer().serializeToString(body.getElementsByTagNameNS(W, kind)[0]!)).toBe(new XMLSerializer().serializeToString(historical))
    expect(unzipSync(fixed)["word/comments.xml"]).toEqual(unzipSync(tracked)["word/comments.xml"])
    expect(exportDocxComments(fixed, [{ id: "current", text: "Current pointer works.", author: "Reviewer", anchor: { paragraph_text: "Kept.", pid: rows[0]!.pids[0] } }]).unresolved).toEqual([])
    expect(withParagraphPointers(fixed)).toBe(fixed)
    if (kind === "moveFrom") {
      // Restore the historical container into accepted content, as a resolved
      // move can do. Only now is its bookmark an active duplicate to repair.
      const moved = body.getElementsByTagNameNS(W, kind)[0]!
      while (moved.firstChild) moved.parentNode!.insertBefore(moved.firstChild, moved)
      moved.parentNode!.removeChild(moved)
      const restored = zipSync({ ...unzipSync(fixed), "word/document.xml": strToU8(new XMLSerializer().serializeToString(body)) })
      expect(readDocxCommentParagraphs(restored)[0]!.pids).toHaveLength(2)
      const repaired = withParagraphPointers(restored)
      expect(readDocxCommentParagraphs(repaired)).toEqual([{ paragraph_index: 0, paragraph_text: "Kept.Moved.", pids: rows[0]!.pids }])
      expect(unzipSync(repaired)["word/comments.xml"]).toEqual(unzipSync(fixed)["word/comments.xml"])
      expect(withParagraphPointers(repaired)).toBe(repaired)
    }
    const historicalParts = unzipSync(rendered)
    const historicalDoc = new DOMParser().parseFromString(strFromU8(historicalParts["word/document.xml"]!), "application/xml")
    const historicalParagraph = historicalDoc.getElementsByTagNameNS(W, "p")[0]!
    const container = historicalDoc.createElementNS(W, `w:${kind}`)
    container.setAttributeNS(W, "w:id", "3")
    historicalParagraph.parentNode!.replaceChild(container, historicalParagraph)
    container.appendChild(historicalParagraph)
    const historicalOnly = zipSync({ ...historicalParts, "word/document.xml": strToU8(new XMLSerializer().serializeToString(historicalDoc)) })
    expect(readDocxCommentParagraphs(historicalOnly)).toEqual([])
    expect(withParagraphPointers(historicalOnly)).toBe(historicalOnly)
  }
})
