import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { readDocxCommentParagraphs } from "../../../dochaus/lib/docx-comments"

// Word proposals accepted after delivery change a body that the drafting notes
// no longer describe. Re-rendering the notes would discard the lawyer's own edits,
// so the accepted changes are appended as plain dated paragraphs at the end of the
// notes body instead. Only word/document.xml changes, and only by the inserted
// paragraphs; every other part and all existing text keep their exact bytes.

export type AcceptedNotesChange = { document: string; before: string; after: string; reason?: string; acceptedAt: string }

const BODY = "word/document.xml"
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const HEADING = "Changes accepted after these notes were prepared"
const INTRO = "These changes were accepted in Word after the final check. The final check was not re-run on them. Changes accepted while the notes could not be updated are not listed."

/** Returns new notes bytes with `changes` appended in order. The heading and
 * introduction are added once; a later call appends only its change paragraphs. */
export function appendAcceptedChangesToNotes(notes: Uint8Array, changes: AcceptedNotesChange[]): Uint8Array {
  if (!changes.length) return notes
  const entries = changes.map(changeText)
  const parts = unzipSync(notes)
  if (!parts[BODY]) throw new Error("The drafting notes are not a Word document.")
  const xml = strFromU8(parts[BODY])
  const w = new RegExp(`\\bxmlns:([A-Za-z_][\\w.-]*)=(["'])${W.replace(/[.]/g, "\\.")}\\2`).exec(xml)?.[1]
  if (!w) throw new Error("The drafting notes do not use a supported Word document namespace.")
  const at = insertionPoint(xml, w)
  const started = readDocxCommentParagraphs(notes).some(row => row.paragraph_text.trim() === HEADING)
  const added = [
    ...started ? [] : [paragraph(w, HEADING, headingStyle(parts["word/styles.xml"])), paragraph(w, INTRO)],
    ...entries.map(text => paragraph(w, text)),
  ].join("")
  parts[BODY] = strToU8(xml.slice(0, at) + added + xml.slice(at))
  return zipSync(parts, { mtime: new Date(1980, 0, 1) })
}

// `<YYYY-MM-DD> — <document>: “<before>” → “<after>”.` with the quoted text
// exact; the date is the calendar date written in the acceptance timestamp.
function changeText(change: AcceptedNotesChange) {
  const date = /^(\d{4}-\d{2}-\d{2})(?:T|$)/.exec(change.acceptedAt)?.[1]
  if (!date || Number.isNaN(Date.parse(change.acceptedAt)))
    throw new Error(`Accepted change has an invalid acceptance time: ${JSON.stringify(change.acceptedAt)}.`)
  if (!change.document.trim()) throw new Error("Accepted change has no document name.")
  if (!change.before && !change.after) throw new Error(`Accepted change to ${change.document} has no text before or after it.`)
  const edit = !change.before ? `added “${change.after}”`
    : !change.after ? `deleted “${change.before}”`
    : `“${change.before}” → “${change.after}”`
  const reason = change.reason?.trim()
  return `${date} — ${change.document}: ${edit}.${reason ? ` Reason: ${reason}${/[.!?]$/.test(reason) ? "" : "."}` : ""}`
}

// After the last block and before the body-level section properties, which Word
// requires to stay the body's final child. A sectPrChange can nest another
// sectPr inside it, so the outer element is found by depth.
function insertionPoint(xml: string, w: string) {
  const close = xml.lastIndexOf(`</${w}:body>`)
  if (close < 0) throw new Error("The drafting notes have no document body.")
  const head = xml.slice(0, close).trimEnd()
  const tags = [...head.matchAll(new RegExp(`<${w}:sectPr(?=[\\s/>])[^>]*?(/?)>|</${w}:sectPr\\s*>`, "g"))]
  const last = tags.at(-1)
  if (!last || last.index + last[0].length !== head.length) return close
  let depth = 0
  for (let index = tags.length - 1; index >= 0; index--) {
    const tag = tags[index]!
    if (tag[0].startsWith("</")) depth++
    else if (!tag[1]) depth--
    if (!depth) return tag.index
  }
  throw new Error("The drafting notes have unbalanced section properties.")
}

// The renderer writes markdown "##" sections (such as its advisory section) as
// Word's built-in "heading 2". Match by name because Word can localize the id;
// without that style, or without a styles part, the heading stays unstyled.
function headingStyle(styles: Uint8Array | undefined) {
  if (!styles) return undefined
  for (const [style] of strFromU8(styles).matchAll(/<(?:[\w.-]+:)?style\b[^>]*?(?:\/>|>[\s\S]*?<\/(?:[\w.-]+:)?style>)/g)) {
    const open = /^<[^>]*>/.exec(style)![0]
    const id = /\s(?:[\w.-]+:)?styleId=(["'])(.*?)\1/.exec(open)?.[2]
    const name = /<(?:[\w.-]+:)?name\b[^>]*?\s(?:[\w.-]+:)?val=(["'])(.*?)\1/.exec(style)?.[2]
    if (id && /\s(?:[\w.-]+:)?type=(["'])paragraph\1/.test(open) && name?.toLowerCase() === "heading 2") return id
  }
  return undefined
}

function paragraph(w: string, text: string, style?: string) {
  const properties = style ? `<${w}:pPr><${w}:pStyle ${w}:val="${style}"/></${w}:pPr>` : ""
  // Keep line breaks and tabs in quoted text as Word breaks and tabs, not lost whitespace.
  const content = text.split(/(\r\n|\r|\n|\t)/).map(piece =>
    piece === "\t" ? `<${w}:tab/>`
      : /^[\r\n]+$/.test(piece) ? `<${w}:br/>`
      : piece ? `<${w}:t xml:space="preserve">${escapeXml(piece)}</${w}:t>` : "").join("")
  return `<${w}:p>${properties}<${w}:r>${content}</${w}:r></${w}:p>`
}

// Characters XML 1.0 cannot carry are dropped; they cannot appear in a Word document.
function escapeXml(text: string) {
  return text.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g, "")
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
}
