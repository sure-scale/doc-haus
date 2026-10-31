import { DOMParser } from "@xmldom/xmldom"
import { unzipSync } from "fflate"

// Browser-safe presence check, not a style resolver. A mixed-format new row
// states only its "on" spans; an absent property means "off" only when nothing
// in the cascade that reaches a new, unstyled cell paragraph names bold, italic
// or underline. Any presence (explicit off included), or anything this check
// cannot classify, makes the table not neutral.
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const REL = "http://schemas.openxmlformats.org/package/2006/relationships"
const STYLES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
const UNCLASSIFIED_NS = new Set([
  "http://purl.oclc.org/ooxml/wordprocessingml/main",
  "http://schemas.openxmlformats.org/markup-compatibility/2006",
])
const TOGGLES = new Set(["b", "bCs", "i", "iCs", "u"])
// Style references this check does not follow.
const REFERENCES = new Set(["rStyle", "pStyle", "link"])
const PARTS = { body: "word/document.xml", rels: "word/_rels/document.xml.rels", styles: "word/styles.xml" }
const PART_LIMIT = 16 * 1024 * 1024
const CHAIN_LIMIT = 64

export type DocxRowFormatCascade = { neutral: true } | { neutral: false; reason: string }
class Refusal extends Error {}
function refuse(reason: string): never { throw new Refusal(reason) }
const elements = (node: Node) => Array.from(node.childNodes).filter((n): n is Element => n.nodeType === 1)
const direct = (node: Node, name: string) => elements(node).filter(el => el.namespaceURI === W && el.localName === name)

function parse(bytes: Uint8Array, name: string) {
  let text: string
  try { text = new TextDecoder("utf-8", { fatal: true }).decode(bytes) } catch { refuse(`unreadable_part:${name}`) }
  const fail = () => refuse(`invalid_xml:${name}`)
  const doc = new DOMParser({ errorHandler: { warning: fail, error: fail, fatalError: fail } }).parseFromString(text, "application/xml")
  if (doc.doctype || Array.from(doc.childNodes).some(n => n.nodeType === 10) || !doc.documentElement) refuse(`invalid_xml:${name}`)
  return doc
}
function stylesTarget(rels: Document | undefined) {
  if (!rels) return undefined
  const root = rels.documentElement!
  if (root.namespaceURI !== REL || root.localName !== "Relationships") refuse("invalid_relationships")
  // A styles relationship of any other schema is unclassified, not absent.
  if (elements(root).some(el => el.getAttribute("Type") !== STYLES && (el.getAttribute("Type") ?? "").endsWith("/styles"))) refuse("unclassified_styles_part")
  const styles = elements(root).filter(el => el.getAttribute("Type") === STYLES)
  if (styles.length > 1) refuse("duplicate_styles_part")
  if (!styles.length) return undefined
  const rel = styles[0]!
  if (rel.namespaceURI !== REL || rel.localName !== "Relationship" || (rel.hasAttribute("TargetMode") && rel.getAttribute("TargetMode") !== "Internal") ||
    !["styles.xml", "/word/styles.xml"].includes(rel.getAttribute("Target") ?? "")) refuse("unclassified_styles_part")
  return PARTS.styles
}
// Every element in a scope that reaches the new paragraphs' runs.
function scan(scope: Element, where: string) {
  for (const el of [scope, ...Array.from(scope.getElementsByTagName("*"))]) {
    if (el.namespaceURI && UNCLASSIFIED_NS.has(el.namespaceURI)) refuse(`unclassified_markup:${where}`)
    if (el.namespaceURI !== W) continue
    if (TOGGLES.has(el.localName)) refuse(`${where}:${el.localName}`)
    if (REFERENCES.has(el.localName)) refuse(`style_reference:${where}:${el.localName}`)
  }
}

function check(source: Uint8Array, tableIndex: number) {
  if (!Number.isSafeInteger(tableIndex) || tableIndex < 0) refuse("invalid_table_index")
  let parts: Record<string, Uint8Array>
  try {
    parts = unzipSync(source, { filter: file => {
      const wanted = Object.values(PARTS).includes(file.name)
      if (wanted && file.originalSize > PART_LIMIT) throw new Error("oversized")
      return wanted
    } })
  } catch { refuse("unreadable_package") }
  if (!parts[PARTS.body]) refuse("main_document_missing")
  const body = parse(parts[PARTS.body]!, "document")
  const root = body.documentElement!
  if (root.namespaceURI !== W || root.localName !== "document") refuse("unclassified_document")
  const bodies = direct(root, "body")
  if (bodies.length !== 1) refuse("invalid_body")
  const table = direct(bodies[0]!, "tbl")[tableIndex]
  if (!table) refuse("table_missing")

  const target = stylesTarget(parts[PARTS.rels] ? parse(parts[PARTS.rels]!, "relationships") : undefined)
  // A styles part Word would not load is ambiguous rather than absent.
  if (!target && parts[PARTS.styles]) refuse("unrelated_styles_part")
  if (target && !parts[target]) refuse("styles_part_missing")
  const styles = target ? parse(parts[target]!, "styles") : undefined
  const stylesRoot = styles?.documentElement
  if (stylesRoot && (stylesRoot.namespaceURI !== W || stylesRoot.localName !== "styles")) refuse("unclassified_styles")
  const definitions = stylesRoot ? direct(stylesRoot, "style") : []
  if (styles && styles.getElementsByTagNameNS(W, "style").length !== definitions.length) refuse("nested_style_definition")
  const docDefaults = stylesRoot ? direct(stylesRoot, "docDefaults") : []
  if (styles && (docDefaults.length > 1 || styles.getElementsByTagNameNS(W, "docDefaults").length !== docDefaults.length)) refuse("ambiguous_doc_defaults")
  for (const defaults of docDefaults) scan(defaults, "doc_defaults")

  const type = (style: Element) => style.getAttributeNS(W, "type") || "paragraph"
  const isDefault = (style: Element) => {
    const value = style.getAttributeNS(W, "default")
    if (value === null || value === "") return false
    if (["1", "true", "on"].includes(value)) return true
    if (["0", "false", "off"].includes(value)) return false
    refuse("unclassified_default_style")
  }
  const defaultStyle = (kind: string) => {
    const found = definitions.filter(s => isDefault(s) && type(s) === kind)
    if (found.length > 1) refuse(`duplicate_default_style:${kind}`)
    if (found.length && !found[0]!.getAttributeNS(W, "styleId")) refuse(`default_style_without_id:${kind}`)
    return found[0]?.getAttributeNS(W, "styleId") ?? undefined
  }
  const chain = (kind: string, start: string | undefined, where: string) => {
    const seen = new Set<string>()
    let id = start
    while (id !== undefined) {
      if (seen.has(id)) refuse(`cyclic_style:${where}`)
      if (seen.size >= CHAIN_LIMIT) refuse(`style_chain_too_deep:${where}`)
      seen.add(id)
      const matches = definitions.filter(s => s.getAttributeNS(W, "styleId") === id)
      if (!matches.length) refuse(`missing_style:${where}`)
      if (matches.length > 1) refuse(`duplicate_style:${where}`)
      const style = matches[0]!
      if (type(style) !== kind) refuse(`style_type_mismatch:${where}`)
      // Conditional table formatting (tblStylePr) is part of this scan.
      scan(style, where)
      const bases = direct(style, "basedOn")
      if (bases.length > 1 || bases.length && !bases[0]!.getAttributeNS(W, "val")) refuse(`ambiguous_based_on:${where}`)
      id = bases[0]?.getAttributeNS(W, "val") ?? undefined
    }
  }

  const tblPr = direct(table, "tblPr")
  if (tblPr.length > 1) refuse("ambiguous_table_properties")
  const refs = tblPr.length ? direct(tblPr[0]!, "tblStyle") : []
  if (refs.length > 1 || refs.length && !refs[0]!.getAttributeNS(W, "val")) refuse("ambiguous_table_style")
  chain("table", refs.length ? refs[0]!.getAttributeNS(W, "val")! : defaultStyle("table"), "table_style")
  // New cell paragraphs carry no pStyle, so the default paragraph style applies.
  chain("paragraph", defaultStyle("paragraph"), "paragraph_style")
  chain("character", defaultStyle("character"), "character_style")
}

/** Is the cascade over new paragraphs in body table `tableIndex` (0-based,
 * direct children of w:body) free of bold, italic and underline sources? */
export function docxRowFormatCascade(source: Uint8Array, tableIndex: number): DocxRowFormatCascade {
  try {
    check(source, tableIndex)
    return { neutral: true }
  } catch (error) {
    if (error instanceof Refusal) return { neutral: false, reason: error.message }
    return { neutral: false, reason: "unclassified_package" }
  }
}
