import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { paragraphPointers } from "../../lib/drafting-amendment"
import { visibleInlineText } from "../../lib/markdown-visible-text"

export type DraftPresentation = { source: string; parts: { heading: string; display_heading?: string; part?: number }[] }

/** Keep drafting boundaries out of publication while retaining their existing
 * paragraph pointers for template coverage. Only explicit top-level part
 * headings are changed; body paragraphs, tables and lower headings retain their
 * native structure, including when their text repeats a part label. */
export function presentDraft(bytes: Uint8Array, parts: DraftPresentation["parts"]) {
  const entries = unzipSync(bytes)
  const document = new DOMParser().parseFromString(strFromU8(entries["word/document.xml"]!), "application/xml")
  const namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  const attribution = new Map<string, { draft_part: number | null; draft_heading: boolean }>()
  let next = 0
  let part: number | null = null
  for (const paragraph of Array.from(document.getElementsByTagNameNS(namespace, "p"))) {
    const text = Array.from(paragraph.getElementsByTagNameNS(namespace, "t")).map(node => node.textContent ?? "").join("")
    const style = paragraph.getElementsByTagNameNS(namespace, "pStyle")[0]?.getAttributeNS(namespace, "val")
    const boundary = parts[next]
    const heading = style === "Heading2" && boundary && text === visibleInlineText(boundary.heading)
    if (heading) {
      part = boundary.part ?? null
      next += 1
      if (boundary.display_heading === "") {
        paragraph.parentNode?.removeChild(paragraph)
        continue
      }
      if (boundary.display_heading !== undefined) {
        const runs = Array.from(paragraph.getElementsByTagNameNS(namespace, "r"))
        runs.forEach(run => run.parentNode?.removeChild(run))
        const run = document.createElementNS(namespace, "w:r")
        const value = document.createElementNS(namespace, "w:t")
        value.appendChild(document.createTextNode(visibleInlineText(boundary.display_heading)))
        run.appendChild(value)
        paragraph.appendChild(run)
      }
    }
    paragraphPointers(paragraph).forEach(pointer => attribution.set(pointer.pid, { draft_part: part, draft_heading: Boolean(heading && boundary.part !== undefined) || style === "Heading1" }))
  }
  if (next !== parts.length) throw new Error("Rendered draft does not contain all planned part boundaries in order.")
  entries["word/document.xml"] = strToU8(new XMLSerializer().serializeToString(document))
  return { bytes: zipSync(entries), attribution }
}
