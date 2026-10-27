import { DOMParser } from "@xmldom/xmldom"
import { strFromU8, unzipSync } from "fflate"
import mammoth from "mammoth"
import { writeListLabels } from "../../../dochaus/lib/docx-list-labels"
import { markTrackedChanges } from "../../../dochaus/lib/docx-tracked-text"
import type { DeskSourcePreview } from "../../../dochaus/lib/word-source-desk-contract"
import { normalizeExtractedText } from "./sanitize"

type Block = NonNullable<DeskSourcePreview["blocks"]>[number]
type Model = { type: string; value?: string; children?: Model[] }
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const normalized = (text: string) => normalizeExtractedText(text).text
const child = (node: Element | undefined, name: string) => node && Array.from(node.childNodes)
  .find((row): row is Element => row.nodeType === 1 && (row as Element).namespaceURI === W && (row as Element).localName === name)
const value = (node: Element | undefined) => node?.getAttributeNS(W, "val") ?? undefined

function parse(bytes: Uint8Array) {
  const fail = () => { throw new Error("Invalid DOCX preview XML") }
  return new DOMParser({ errorHandler: { warning: () => {}, error: fail, fatalError: fail } })
    .parseFromString(strFromU8(bytes), "application/xml")
}

// An explicit body outline level must suppress an inherited/built-in heading.
function outline(node: Element | undefined): number | null | undefined {
  const setting = child(node, "outlineLvl")
  if (!setting) return
  const raw = value(setting)
  return raw !== undefined && /^[0-8]$/.test(raw) ? Number(raw) + 1 : null
}

function paragraphLevels(styles: Document | undefined) {
  const catalog = new Map<string, Element>()
  let defaultStyle: string | undefined
  for (const style of Array.from(styles?.getElementsByTagNameNS(W, "style") ?? [])) {
    if (style.getAttributeNS(W, "type") !== "paragraph") continue
    const id = style.getAttributeNS(W, "styleId")
    if (!id) continue
    catalog.set(id, style)
    if (["1", "true", "on"].includes(style.getAttributeNS(W, "default") ?? "")) defaultStyle = id
  }
  return (paragraph: Element): number | undefined => {
    const properties = child(paragraph, "pPr")
    const direct = outline(properties)
    if (direct !== undefined) return direct ?? undefined
    const chain: Element[] = []
    const seen = new Set<string>()
    for (let id = value(child(properties, "pStyle")) ?? defaultStyle; id && !seen.has(id);) {
      seen.add(id)
      const style = catalog.get(id)
      if (!style) break
      chain.push(style)
      const inherited = outline(child(style, "pPr"))
      if (inherited !== undefined) return inherited ?? undefined
      id = value(child(style, "basedOn"))
    }
    for (const style of chain) {
      const named = value(child(style, "name"))?.match(/^heading ([1-9])$/i)?.[1]
        ?? style.getAttributeNS(W, "styleId")?.match(/^Heading([1-9])$/)?.[1]
      if (named) return Number(named)
    }
  }
}

// This is Mammoth's raw-text traversal: only text, tabs and paragraph boundaries
// contribute. The full canonical equality check below guards model/API drift.
function rawText(model: Model): string {
  if (model.type === "text") return model.value ?? ""
  if (model.type === "tab") return "\t"
  return (model.children ?? []).map(rawText).join("") + (model.type === "paragraph" ? "\n\n" : "")
}

function xmlText(paragraph: Element) {
  return Array.from(paragraph.getElementsByTagNameNS(W, "*"))
    .map(node => node.localName === "t" ? node.textContent ?? "" : node.localName === "tab" && (node.parentNode as Element).localName === "r" ? "\t" : "").join("")
}

/** Add structure only to complete, exact DOCX previews. Other formats and any
 * unsupported structure retain canonical plain text. No HTML or images escape. */
export async function documentPreviewBlocks(fileName: string, bytes: Buffer, canonicalText: string, truncated = false): Promise<Block[] | undefined> {
  if (truncated || !fileName.toLowerCase().endsWith(".docx")) return
  try {
    // Keep this in lockstep with ingest.ts's DOCX extraction. In particular,
    // pending deletions/moves remain labelled evidence, never accepted text.
    const prepared = markTrackedChanges(await writeListLabels(bytes))
    const parts = unzipSync(prepared)
    if (!parts["word/document.xml"]) return
    const document = parse(parts["word/document.xml"])
    const body = document.getElementsByTagNameNS(W, "body")[0]
    if (!body) return
    const paragraphs = Array.from(body.getElementsByTagNameNS(W, "p"))
    const level = paragraphLevels(parts["word/styles.xml"] ? parse(parts["word/styles.xml"]) : undefined)
    const models: Model[] = []
    let extracted = ""
    await mammoth.convertToHtml({ buffer: prepared }, {
      externalFileAccess: false,
      transformDocument(model: Model) {
        extracted = normalized(rawText(model))
        const visit = (node: Model) => {
          if (node.type === "paragraph") models.push(node)
          for (const row of node.children ?? []) visit(row)
        }
        visit(model)
        // Use only the public model hook. An empty document avoids HTML/image
        // conversion; the generated HTML is neither used nor returned.
        return { ...model, children: [] }
      },
    })
    if (extracted !== canonicalText || models.length !== paragraphs.length) return
    const blocks: Block[] = []
    for (let index = 0; index < paragraphs.length; index++) {
      const paragraph = paragraphs[index]
      const text = normalized((models[index].children ?? []).map(rawText).join(""))
      // Do not assign metadata by ordinal unless every corresponding paragraph
      // agrees too (e.g. text boxes/AlternateContent may be read differently).
      if (text !== normalized(xmlText(paragraph))) return
      const heading = level(paragraph)
      blocks.push(heading ? { kind: "heading", level: heading, text } : { kind: "paragraph", text })
    }
    return normalized(blocks.map(block => block.text + "\n\n").join("")) === canonicalText ? blocks : undefined
  } catch {
    return undefined
  }
}
