import { unzipSync } from "fflate"
import { DOMParser } from "@xmldom/xmldom"
import { normalizeUntrustedText, type InjectionFinding } from "../../../dochaus/lib/untrusted-text"

export { detectInjection, type InjectionFinding } from "../../../dochaus/lib/untrusted-text"

// Keep the ingest API while sharing its exact normalization with research and
// live document extraction. Chunk offsets refer to this normalized text.
export function normalizeExtractedText(raw: string) {
  return normalizeUntrustedText(raw)
}

// Advisory counts of explicit current run properties in the main document.
// Inherited/style formatting, effective background/theme colour and other story
// parts are not inferred. This never rewrites extracted text or its offsets.
export function scanDocxHiddenContent(buffer: Buffer): InjectionFinding[] {
  try { return hiddenContentFindings(buffer) } catch {
    return [{ rule: "hidden-content-unavailable", detail: "The main-document hidden-formatting advisory could not be completed; hidden formatting has not been ruled out." }]
  }
}

function hiddenContentFindings(buffer: Buffer): InjectionFinding[] {
  const main = "word/document.xml", max = 16 * 1024 * 1024, aggregate = 128 * 1024 * 1024
  const reject = (): never => { throw Error("Hidden-formatting advisory unavailable") }
  if (!buffer.length || buffer.length > aggregate) return reject()
  let size = 0
  const names = new Set<string>()
  const parts = unzipSync(buffer, { filter(entry) {
    if (names.has(entry.name) || names.size >= 4096 || ![0, 8].includes(entry.compression) ||
      entry.originalSize > aggregate || (size += entry.originalSize) > aggregate || entry.name === main && entry.originalSize > max) return reject()
    names.add(entry.name)
    return entry.name === main
  } })
  const bytes = parts[main]
  if (!bytes?.length || bytes.length > max) return reject()
  const xml = new TextDecoder("utf-8", { fatal: true }).decode(bytes)
  if (/<!DOCTYPE|<!ENTITY/i.test(xml)) return reject()
  const document = new DOMParser({ errorHandler: { warning: reject, error: reject, fatalError: reject } }).parseFromString(xml, "application/xml")
  const root = document.documentElement, namespace = root?.namespaceURI
  if (document.doctype || root?.localName !== "document" || !namespace || ![
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "http://purl.oclc.org/ooxml/wordprocessingml/main",
  ].includes(namespace)) return reject()
  const elements = (node: Node) => Array.from(node.childNodes).filter((child): child is Element => child.nodeType === 1)
  const named = (node: Node, name: string) => elements(node).filter(child => child.namespaceURI === namespace && child.localName === name)
  const bodies = named(root, "body")
  if (bodies.length !== 1) return reject()
  const value = (property: Element) => property.hasAttributeNS(namespace, "val") ? property.getAttributeNS(namespace, "val")!.trim() : undefined
  let hidden = 0, white = 0, tiny = 0, unresolvedTheme = false
  const historical = new Set(["del", "moveFrom", "rPrChange", "pPrChange", "tblPrChange", "tblGridChange", "trPrChange", "tcPrChange", "sectPrChange"])
  const visit = (node: Element) => {
    if (node.namespaceURI === "http://schemas.openxmlformats.org/markup-compatibility/2006" && node.localName === "AlternateContent" ||
      node.namespaceURI === namespace && ["altChunk", "subDoc"].includes(node.localName)) return reject()
    if (node.namespaceURI === namespace && historical.has(node.localName)) return
    if (node.namespaceURI === namespace && node.localName === "r") {
      const properties = named(node, "rPr")
      if (properties.length > 1) return reject()
      if (properties[0]) {
        const single = (name: string) => {
          const matches = named(properties[0]!, name)
          if (matches.length > 1) return reject()
          return matches[0]
        }
        const vanish = single("vanish"), colour = single("color"), fontSize = single("sz")
        if (vanish) {
          const setting = value(vanish)
          if (setting !== undefined && !["0", "false", "off", "1", "true", "on"].includes(setting)) return reject()
          if (setting === undefined || ["1", "true", "on"].includes(setting)) hidden++
        }
        if (colour) {
          const setting = value(colour)
          if (setting === undefined || !/^(?:auto|[0-9a-fA-F]{6})$/.test(setting)) return reject()
          if (/^ffffff$/i.test(setting)) {
            if (colour.hasAttributeNS(namespace, "themeColor")) unresolvedTheme = true
            else white++
          }
        }
        if (fontSize) {
          const setting = value(fontSize)
          if (setting === undefined || !/^\d+$/.test(setting) || Number(setting) <= 0 || !Number.isSafeInteger(Number(setting))) return reject()
          if (Number(setting) <= 4) tiny++
        }
      }
    }
    for (const child of elements(node)) visit(child)
  }
  visit(bodies[0]!)
  return [
    ...(hidden
      ? [
          {
            rule: "hidden-text",
            detail: `${hidden} current main-document run(s) with explicit hidden ("vanish") formatting; extracted text is retained`,
          },
        ]
      : []),
    ...(white ? [{ rule: "white-text", detail: `${white} current main-document run(s) with explicit white text; effective background is not assessed` }] : []),
    ...(tiny ? [{ rule: "tiny-text", detail: `${tiny} current main-document run(s) with explicit near-invisible text size (2pt or smaller)` }] : []),
    ...(unresolvedTheme ? [{ rule: "hidden-content-unavailable", detail: "A main-document white colour fallback has a theme override; its effective colour was not assessed." }] : []),
  ]
}
