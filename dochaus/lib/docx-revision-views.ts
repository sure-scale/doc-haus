import { DOMParser } from "@xmldom/xmldom"
import { strFromU8, unzipSync } from "fflate"
import { docxodus } from "./docxodus"
import { exportDocxComments, readDocxCommentParagraphs } from "./docx-comments"

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const KINDS = new Map([["ins", "insert"], ["del", "delete"], ["moveFrom", "move from"], ["moveTo", "move to"]])

export type RevisionBodyView = "current_body" | "baseline_body" | "marked_up_body"

/** A read-only projection of the main body (word/document.xml); formatting
 * changes are not marked. readDocxMarkup lists every part's changes. PIDs/
 * paragraph ordinals in a baseline belong to that projected artifact; they are
 * never amendment targets in current Word. */
export async function readDocxBodyView(bytes: Uint8Array, view: RevisionBodyView) {
  if (view !== "marked_up_body") {
    const dx = await docxodus()
    const projected = view === "baseline_body" ? await dx.docxDiffRejectRevisions(bytes) : await dx.docxDiffAcceptRevisions(bytes)
    return readDocxCommentParagraphs(projected)
  }
  const part = unzipSync(bytes)["word/document.xml"]
  if (!part) throw new Error("invalid_docx_package")
  const fail = () => { throw new Error("invalid_docx_xml") }
  const doc = new DOMParser({ errorHandler: { warning: fail, error: fail, fatalError: fail } }).parseFromString(strFromU8(part), "application/xml")
  const text = (node: Element): string => {
    const name = node.namespaceURI === W ? node.localName : ""
    // A property change contains historical formatting, not another copy of
    // the paragraph's wording. Paragraph-break changes are labelled below.
    if (["pPr", "rPr", "tblPr", "tcPr", "trPr", "sectPr"].includes(name)) return ""
    if (name === "t" || name === "delText") return node.textContent ?? ""
    if (name === "tab") return "\t"
    if (name === "br" || name === "cr") return "\n"
    const content = Array.from(node.childNodes).filter((child): child is Element => child.nodeType === 1 && (child as Element).localName !== "p").map(text).join("")
    const kind = KINDS.get(name)
    return kind ? `⟦${kind} by ${JSON.stringify(node.getAttributeNS(W, "author") || "unknown author")}: ${content}⟧` : content
  }
  return Array.from(doc.getElementsByTagNameNS(W, "p")).flatMap((paragraph, xmlIndex) => {
    const children = (node: Element, name: string) => Array.from(node.childNodes).filter((child): child is Element => child.nodeType === 1 && (child as Element).namespaceURI === W && (child as Element).localName === name)
    const marks = children(paragraph, "pPr").flatMap(properties =>
      children(properties, "rPr").flatMap(run => Array.from(run.childNodes)
        .filter((child): child is Element => child.nodeType === 1 && KINDS.has((child as Element).localName))))
    let paragraphText = text(paragraph) + marks.map(mark => `⟦${KINDS.get(mark.localName)} paragraph break by ${JSON.stringify(mark.getAttributeNS(W, "author") || "unknown author")}⟧`).join("")
    // Row/cell revisions are siblings of the paragraph in trPr/tcPr. Without
    // their enclosing marks a deleted row would look like ordinary body text.
    for (let parent = paragraph.parentNode; parent; parent = parent.parentNode) {
      if (parent.nodeType !== 1 || (parent as Element).namespaceURI !== W) continue
      const container = parent as Element
      const property = container.localName === "tr" ? "trPr" : container.localName === "tc" ? "tcPr" : undefined
      if (!property) continue
      for (const properties of children(container, property)) for (const mark of Array.from(properties.childNodes)) {
        if (mark.nodeType !== 1 || (mark as Element).namespaceURI !== W) continue
        const revision = mark as Element
        const operation = KINDS.get(revision.localName) ?? ({ cellIns: "insert", cellDel: "delete", cellMerge: "merge" } as Record<string, string>)[revision.localName]
        if (operation) paragraphText = `⟦${operation} ${container.localName === "tr" ? "row" : "cell"} by ${JSON.stringify(revision.getAttributeNS(W, "author") || "unknown author")}: ${paragraphText}⟧`
      }
    }
    return paragraphText ? [{ paragraph_index: xmlIndex, paragraph_text: paragraphText, pids: [] as string[] }] : []
  })
}

export async function sourceRevisionDisclosure(source: Uint8Array) {
  const revisions = await (await docxodus()).getRevisions(source)
  if (!revisions.length) return
  const authors = [...new Set(revisions.map(row => row.author || "unknown author"))]
  return { count: revisions.length, authors,
    message: `The source contains ${revisions.length} pending tracked change${revisions.length === 1 ? "" : "s"} by ${authors.map(author => JSON.stringify(author)).join(", ")}. This comparison uses the source with those changes accepted; they are not retained as pending changes in this redline. The original source is unchanged.` }
}

/** Put the limitation into the exported Word file itself, as well as the
 * viewer/notes notice. This records no accept/reject decision on the source. */
export async function discloseSourceRevisions(redline: Uint8Array, source: Uint8Array, author: string) {
  const disclosure = await sourceRevisionDisclosure(source)
  if (!disclosure) return { bytes: redline, disclosure }
  // This document-level notice has no original paragraph anchor to preserve.
  // A changed first paragraph may have an unresolved deleted boundary; use the
  // first surviving paragraph whose exact accepted range is safe instead.
  for (const paragraph of readDocxCommentParagraphs(redline)) {
    const annotated = exportDocxComments(redline, [{ id: "source-pending-revisions", author, text: disclosure.message,
      anchor: { paragraph_index: paragraph.paragraph_index, paragraph_text: paragraph.paragraph_text } }])
    if (!annotated.unresolved.length) return { bytes: annotated.bytes, disclosure }
  }
  throw new Error("The source's pending-revision notice could not be attached to a safe surviving paragraph. Review its tracked changes in Word before exporting this comparison.")
}
