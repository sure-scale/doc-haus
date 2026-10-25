import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { polishHeadingSpacing } from "../../lib/docx-polish"
import { paragraphPointers } from "../../lib/drafting-amendment"
import { visibleInlineText } from "../../lib/markdown-visible-text"

/** `title: false` leaves the planned title out of publication: a letter reads
 * as one letter, not as a titled document. */
export type DraftPresentation = { source: string; parts: { heading: string; display_heading?: string; part?: number }[]; title?: false }

/** Keep drafting boundaries out of publication while retaining their existing
 * paragraph pointers for template coverage. A hidden label ("") is removed and
 * the headings of its part move up, keeping their relative depth, until the
 * shallowest is a second-level heading, so the document does not skip from its
 * title to a lower heading; a display heading replaces its label's text; with
 * `title` false the title, the first Heading 1, is removed. Body paragraphs,
 * tables and the headings of other parts retain their native structure,
 * including when their text repeats a part label.
 *
 * A part's text admits no "## " line (checkSection), so the second-level
 * headings of the rendered body are the part labels, in order. A label is found
 * by its rendered text, or by its letters and digits where the renderer reads
 * its Markdown differently, else by position when every label rendered as one.
 * A label the body does not show as a second-level heading is passed over
 * rather than failing the draft: it stays as rendered, and its paragraphs are
 * attributed to the part before it. */
export function presentDraft(bytes: Uint8Array, parts: DraftPresentation["parts"], title = true, headingSpacing = true) {
  const entries = unzipSync(bytes)
  const document = new DOMParser().parseFromString(strFromU8(entries["word/document.xml"]!), "application/xml")
  const namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  const paragraphs = Array.from(document.getElementsByTagNameNS(namespace, "p")).map(paragraph => {
    const styled = paragraph.getElementsByTagNameNS(namespace, "pStyle")[0]
    return {
      paragraph, styled, style: styled?.getAttributeNS(namespace, "val") ?? undefined,
      text: Array.from(paragraph.getElementsByTagNameNS(namespace, "t")).map(node => node.textContent ?? "").join(""),
    }
  })
  const titleAt = title ? -1 : paragraphs.findIndex(row => row.style === "Heading1")
  const labels = paragraphs.flatMap((row, at) => row.style === "Heading2" ? [at] : [])
  const plain = (text: string) => text.normalize("NFKC").toLowerCase().replace(/[^\p{L}\p{N}]+/gu, "")
  const reads = (text: string, heading: string) => text === visibleInlineText(heading)
    || (plain(text) !== "" && [visibleInlineText(heading), heading].some(form => plain(form) === plain(text)))
  // starts[boundary] is the paragraph index of that part's label, if found.
  const starts: (number | undefined)[] = []
  let from = 0
  parts.forEach((boundary, index) => {
    const found = labels.findIndex((at, label) => label >= from && reads(paragraphs[at]!.text, boundary.heading))
    if (found < 0) return
    starts[index] = labels[found]
    from = found + 1
  })
  if (labels.length === parts.length && parts.some((_, index) => starts[index] === undefined)
    && parts.every((_, index) => starts[index] === undefined || starts[index] === labels[index]))
    parts.forEach((_, index) => { starts[index] = labels[index] })
  const boundaryAt = new Map(starts.flatMap((at, index) => at === undefined ? [] : [[at, index] as const]))
  let current: number | undefined
  const owners = paragraphs.map((_, at) => (current = boundaryAt.get(at) ?? current))
  const shifts = parts.map((boundary, index) => {
    if (boundary.display_heading !== "" || starts[index] === undefined) return 0
    const levels = paragraphs.flatMap((row, at) => owners[at] === index && /^Heading[3-6]$/.test(row.style ?? "") ? [Number(row.style!.slice(-1))] : [])
    return levels.length ? Math.min(...levels) - 2 : 0
  })
  const attribution = new Map<string, { draft_part: number | null; draft_heading: boolean }>()
  let promoted = false
  paragraphs.forEach(({ paragraph, styled, style }, at) => {
    const boundary = boundaryAt.has(at) ? parts[boundaryAt.get(at)!] : undefined
    if (at === titleAt || boundary?.display_heading === "") {
      paragraph.parentNode?.removeChild(paragraph)
      return
    }
    if (boundary?.display_heading !== undefined) {
      const runs = Array.from(paragraph.getElementsByTagNameNS(namespace, "r"))
      runs.forEach(run => run.parentNode?.removeChild(run))
      const run = document.createElementNS(namespace, "w:r")
      const value = document.createElementNS(namespace, "w:t")
      value.appendChild(document.createTextNode(visibleInlineText(boundary.display_heading)))
      run.appendChild(value)
      paragraph.appendChild(run)
    }
    const owner = owners[at]
    const shift = owner === undefined ? 0 : shifts[owner]!
    if (!boundary && shift && /^Heading[3-6]$/.test(style ?? "")) {
      styled!.setAttributeNS(namespace, "w:val", `Heading${Number(style!.slice(-1)) - shift}`)
      promoted = true
    }
    const part = owner === undefined ? null : parts[owner]!.part ?? null
    paragraphPointers(paragraph).forEach(pointer => attribution.set(pointer.pid, { draft_part: part, draft_heading: Boolean(boundary && boundary.part !== undefined) || style === "Heading1" }))
  })
  const xml = new XMLSerializer().serializeToString(document)
  // A promoted house heading takes its new style's spacing, as rendering
  // gives it; a template's headings keep their styles' own (`headingSpacing`).
  entries["word/document.xml"] = strToU8(promoted && headingSpacing ? polishHeadingSpacing(xml) : xml)
  return { bytes: zipSync(entries, { mtime: new Date(1980, 0, 1) }), attribution }
}
