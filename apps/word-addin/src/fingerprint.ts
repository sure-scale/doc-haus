import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync } from "fflate"

const wordNamespaces = new Set([
  "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  "http://purl.oclc.org/ooxml/wordprocessingml/main",
])
// Revision-session identifiers describe editing sessions, not the changes
// themselves. Keep unknown attributes and every actual change/format payload.
const sessionAttributes = new Set(["rsidR", "rsidRPr", "rsidDel", "rsidP", "rsidRDefault", "rsidTr", "rsidSect"])
const extendedPropertiesNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
const corePropertiesNamespace = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
const dublinCoreTermsNamespace = "http://purl.org/dc/terms/"
const schemaInstanceNamespace = "http://www.w3.org/2001/XMLSchema-instance"
const xmlNamespace = "http://www.w3.org/XML/1998/namespace"
const cachedCountProperties = ["Words", "Characters", "Paragraphs", "CharactersWithSpaces"] as const
const proofingAnchors = new Set(["spellStart", "spellEnd", "gramStart", "gramEnd"])
const flowContainers = new Set(["body", "tbl", "tr", "tc", "sdt", "sdtContent", "footnote", "endnote", "comment"])
// Only properties with ordinary character-format semantics. Other properties,
// revision records and extensions keep the original run boundaries significant.
const simpleFormatAttributes: Record<string, readonly string[]> = {
  b: ["val"],
  bCs: ["val"],
  i: ["val"],
  iCs: ["val"],
  rFonts: ["ascii", "hAnsi", "eastAsia", "cs"],
  lang: ["val", "eastAsia", "bidi"],
  sz: ["val"],
  szCs: ["val"],
  color: ["val"],
}

function normalizeNumericProperties(root: Element, namespace: string, names: readonly string[], max: number): boolean {
  let changed = false
  for (const name of names) {
    const leaves = Array.from(root.childNodes).filter(
      (node) =>
        node.nodeType === 1 && (node as Element).namespaceURI === namespace && (node as Element).localName === name,
    ) as Element[]
    if (leaves.length !== 1) continue
    const element = leaves[0]!
    const value = element.textContent ?? ""
    if (
      element.attributes.length ||
      !Array.from(element.childNodes).every((child) => child.nodeType === 3 || child.nodeType === 4) ||
      !/^[\t\r\n ]*\+?\d+[\t\r\n ]*$/.test(value) ||
      !Number.isSafeInteger(Number(value)) ||
      Number(value) > max
    )
      continue
    element.textContent = "0"
    changed = true
  }
  return changed
}

function standardTimestamp(value: string): boolean {
  const match = /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})(?:\.\d+)?(?:Z|[+-](\d{2}):(\d{2}))$/.exec(value)
  if (!match) return false
  const [year, month, day, hour, minute, second] = match.slice(1, 7).map(Number) as [
    number,
    number,
    number,
    number,
    number,
    number,
  ]
  const leap = year % 4 === 0 && (year % 100 !== 0 || year % 400 === 0)
  const days = [31, leap ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
  const offsetHour = Number(match[7] ?? 0)
  const offsetMinute = Number(match[8] ?? 0)
  return (
    year > 0 &&
    month >= 1 &&
    month <= 12 &&
    day >= 1 &&
    day <= days[month - 1]! &&
    hour < 24 &&
    minute < 60 &&
    second < 60 &&
    offsetHour <= 14 &&
    offsetMinute < 60 &&
    (offsetHour !== 14 || offsetMinute === 0)
  )
}

function isWordParagraph(paragraph: Element, root: Element): boolean {
  for (let ancestor = paragraph.parentNode; ancestor !== root; ancestor = ancestor.parentNode) {
    if (!ancestor || ancestor.nodeType !== 1) return false
    const element = ancestor as Element
    if (element.namespaceURI !== root.namespaceURI || !flowContainers.has(element.localName)) return false
  }
  return true
}

function simpleProperties(properties: Element, namespace: string, serializer: XMLSerializer): string | undefined {
  if (properties.namespaceURI !== namespace || properties.localName !== "rPr" || properties.attributes.length)
    return undefined
  const seen = new Set<string>()
  for (const node of Array.from(properties.childNodes)) {
    if (node.nodeType !== 1) return undefined
    const property = node as Element
    const allowed = Object.hasOwn(simpleFormatAttributes, property.localName)
      ? simpleFormatAttributes[property.localName]
      : undefined
    if (!allowed || property.namespaceURI !== namespace || property.childNodes.length || seen.has(property.localName))
      return undefined
    seen.add(property.localName)
    for (const attribute of Array.from(property.attributes)) {
      if (attribute.namespaceURI !== namespace || !allowed.includes(attribute.localName) || !attribute.value)
        return undefined
      if (
        ["b", "bCs", "i", "iCs"].includes(property.localName) &&
        !["true", "false", "on", "off", "0", "1"].includes(attribute.value)
      )
        return undefined
      if (["sz", "szCs"].includes(property.localName) && !/^\d+$/.test(attribute.value)) return undefined
      if (property.localName === "color" && !/^(?:[\da-fA-F]{6}|auto)$/.test(attribute.value)) return undefined
    }
    if (!["b", "bCs", "i", "iCs"].includes(property.localName) && !property.attributes.length) return undefined
  }
  return serializer.serializeToString(properties)
}

type SimpleRun = { node: Element; text: Element; value: string; properties: string }

function simpleRun(node: Node, namespace: string, serializer: XMLSerializer): SimpleRun | undefined {
  if (node.nodeType !== 1) return undefined
  const run = node as Element
  if (run.namespaceURI !== namespace || run.localName !== "r" || run.attributes.length) return undefined
  const children = Array.from(run.childNodes)
  let properties = ""
  if (children.length === 2 && children[0]?.nodeType === 1) {
    const key = simpleProperties(children.shift() as Element, namespace, serializer)
    if (key === undefined) return undefined
    properties = key
  }
  if (children.length !== 1 || children[0]?.nodeType !== 1) return undefined
  const text = children[0] as Element
  if (text.namespaceURI !== namespace || text.localName !== "t") return undefined
  if (!Array.from(text.childNodes).every((child) => child.nodeType === 3 || child.nodeType === 4)) return undefined
  for (const attribute of Array.from(text.attributes))
    if (
      attribute.namespaceURI !== xmlNamespace ||
      attribute.localName !== "space" ||
      !["default", "preserve"].includes(attribute.value)
    )
      return undefined
  const value = text.textContent ?? ""
  // Do not interpret default/inherited boundary whitespace, tabs or line breaks.
  // Explicit preservation is required for any spaces at a fragment boundary.
  if (
    !value ||
    /[\t\r\n]/.test(value) ||
    (/^ | $/.test(value) && text.getAttributeNS(xmlNamespace, "space") !== "preserve")
  )
    return undefined
  return { node: run, text, value, properties }
}

function normalizeTextFlow(name: string, root: Element): boolean {
  const expectedRoot =
    name === "word/document.xml"
      ? "document"
      : name === "word/comments.xml"
        ? "comments"
        : name === "word/footnotes.xml"
          ? "footnotes"
          : name === "word/endnotes.xml"
            ? "endnotes"
            : /^word\/header\d+\.xml$/.test(name)
              ? "hdr"
              : /^word\/footer\d+\.xml$/.test(name)
                ? "ftr"
                : undefined
  if (root.localName !== expectedRoot) return false
  const namespace = root.namespaceURI!
  const serializer = new XMLSerializer()
  let changed = false
  for (const paragraph of Array.from(root.getElementsByTagNameNS(namespace, "p"))) {
    if (!isWordParagraph(paragraph, root)) continue
    for (const node of Array.from(paragraph.childNodes)) {
      if (node.nodeType !== 1) continue
      const marker = node as Element
      const attribute = marker.attributes[0]
      if (
        marker.namespaceURI === namespace &&
        marker.localName === "proofErr" &&
        !marker.childNodes.length &&
        marker.attributes.length === 1 &&
        attribute?.namespaceURI === namespace &&
        attribute.localName === "type" &&
        proofingAnchors.has(attribute.value)
      ) {
        paragraph.removeChild(marker)
        changed = true
      }
    }
    let previous: SimpleRun | undefined
    for (const node of Array.from(paragraph.childNodes)) {
      const current = simpleRun(node, namespace, serializer)
      if (!current) {
        previous = undefined
        continue
      }
      if (previous && previous.properties === current.properties) {
        previous.value += current.value
        previous.text.textContent = previous.value
        paragraph.removeChild(current.node)
        changed = true
      } else previous = current
      // A preserve flag is redundant when there is no boundary whitespace.
      // Removing it makes a split/rejoined text run match the unsplit text.
      if (!/^ | $/.test(previous.value) && previous.text.hasAttributeNS(xmlNamespace, "space")) {
        previous.text.removeAttributeNS(xmlNamespace, "space")
        changed = true
      } else if (/^ | $/.test(previous.value) && previous.text.getAttributeNS(xmlNamespace, "space") !== "preserve") {
        previous.text.setAttributeNS(xmlNamespace, "xml:space", "preserve")
        changed = true
      }
    }
  }
  return changed
}

function comparisonContent(name: string, bytes: Uint8Array): Uint8Array {
  const wordPart = name.startsWith("word/") && name.endsWith(".xml")
  const extendedProperties = name === "docProps/app.xml"
  const coreProperties = name === "docProps/core.xml"
  if (!wordPart && !extendedProperties && !coreProperties) return bytes
  const text = strFromU8(bytes)
  const fail = () => {
    throw new Error("Word could not read the document for comparison. Keep the file and try saving again.")
  }
  if (/<!DOCTYPE|<!ENTITY/i.test(text)) fail()
  const doc = new DOMParser({ errorHandler: { warning: fail, error: fail, fatalError: fail } }).parseFromString(
    text,
    "application/xml",
  )
  if (!doc.documentElement || doc.getElementsByTagName("parsererror").length) fail()
  const root = doc.documentElement
  if (coreProperties) {
    if (root.namespaceURI !== corePropertiesNamespace || root.localName !== "coreProperties") return bytes
    let changed = normalizeNumericProperties(root, corePropertiesNamespace, ["revision"], Number.MAX_SAFE_INTEGER)
    const modified = Array.from(root.childNodes).filter(
      (node) =>
        node.nodeType === 1 &&
        (node as Element).namespaceURI === dublinCoreTermsNamespace &&
        (node as Element).localName === "modified",
    ) as Element[]
    const element = modified.length === 1 ? modified[0] : undefined
    if (element) {
      const attribute = element.attributes[0]
      const type = attribute?.value.split(":")
      if (
        element.attributes.length !== 1 ||
        attribute?.namespaceURI !== schemaInstanceNamespace ||
        attribute.localName !== "type" ||
        type?.length !== 2 ||
        !type[0] ||
        type[1] !== "W3CDTF" ||
        element.lookupNamespaceURI(type[0]!) !== dublinCoreTermsNamespace ||
        !Array.from(element.childNodes).every((child) => child.nodeType === 3 || child.nodeType === 4) ||
        !standardTimestamp(element.textContent ?? "")
      )
        return changed ? strToU8(new XMLSerializer().serializeToString(doc)) : bytes
      // Native save/reopen advances the package save timestamp without an edit.
      // Keep all other properties and preserve nonstandard/extended leaves.
      element.textContent = "1970-01-01T00:00:00Z"
      changed = true
    }
    return changed ? strToU8(new XMLSerializer().serializeToString(doc)) : bytes
  }
  if (extendedProperties) {
    if (root.namespaceURI !== extendedPropertiesNamespace || root.localName !== "Properties") return bytes
    let changed = normalizeNumericProperties(root, extendedPropertiesNamespace, cachedCountProperties, 2_147_483_647)
    for (const node of Array.from(root.childNodes)) {
      if (node.nodeType !== 1) continue
      const element = node as Element
      if (element.namespaceURI !== extendedPropertiesNamespace || element.localName !== "TotalTime") continue
      // TotalTime is a numeric editing-duration leaf, not document content.
      // Preserve its attributes and any nonstandard markup instead of dropping
      // the whole property or application-properties part from comparison.
      if (
        !Array.from(element.childNodes).every((child) => child.nodeType === 3 || child.nodeType === 4) ||
        !/^[\s]*[+-]?\d+[\s]*$/.test(element.textContent ?? "")
      )
        continue
      element.textContent = "0"
      changed = true
    }
    return changed ? strToU8(new XMLSerializer().serializeToString(doc)) : bytes
  }
  if (!wordNamespaces.has(root.namespaceURI ?? "")) return bytes
  let changed = false
  for (const node of Array.from(doc.getElementsByTagName("*"))) {
    if (!wordNamespaces.has(node.namespaceURI ?? "")) continue
    for (const attribute of Array.from(node.attributes)) {
      if (!wordNamespaces.has(attribute.namespaceURI ?? "") || !sessionAttributes.has(attribute.localName ?? ""))
        continue
      node.removeAttributeNode(attribute)
      changed = true
    }
  }
  if (name === "word/settings.xml" && wordNamespaces.has(root.namespaceURI ?? "") && root.localName === "settings") {
    for (const node of Array.from(root.childNodes)) {
      if (node.nodeType !== 1) continue
      const element = node as Element
      if (wordNamespaces.has(element.namespaceURI ?? "") && element.localName === "rsids") {
        root.removeChild(element)
        changed = true
      }
      // These two cached completion flags can change after Word finishes
      // background proofing. Keep extensions, annotations and all other data.
      if (
        element.namespaceURI === root.namespaceURI &&
        element.localName === "proofState" &&
        Array.from(element.attributes).every(
          (attribute) =>
            attribute.namespaceURI === element.namespaceURI &&
            (attribute.localName === "spelling" || attribute.localName === "grammar") &&
            (attribute.value === "clean" || attribute.value === "dirty"),
        ) &&
        Array.from(element.childNodes).every(
          (child) => (child.nodeType === 3 || child.nodeType === 4) && !child.nodeValue?.trim(),
        )
      ) {
        root.removeChild(element)
        changed = true
      }
    }
  }
  changed = normalizeTextFlow(name, root) || changed
  return changed ? strToU8(new XMLSerializer().serializeToString(doc)) : bytes
}

/** Compare package contents, not ZIP timestamps, compression, or entry order. */
export async function documentFingerprint(bytes: Uint8Array): Promise<string> {
  let count = 0
  let unpacked = 0
  const entries = unzipSync(bytes, {
    filter(file) {
      const path = file.name.split("/")
      if (file.name.endsWith("/") || path.includes("__MACOSX") || path.at(-1) === ".DS_Store") return false
      count++
      unpacked += file.originalSize
      // Match the receiver's package budget before fflate allocates each part.
      if (count > 1000 || unpacked > 256 * 1024 * 1024) {
        throw new Error("The Word document exceeds the package size limit. Keep the file and return to doc.haus.")
      }
      return true
    },
  })
  const encoder = new TextEncoder()
  const parts: Uint8Array[] = []
  for (const name of Object.keys(entries).sort()) {
    const content = comparisonContent(name, entries[name]!)
    const label = encoder.encode(name)
    const lengths = new Uint8Array(8)
    const view = new DataView(lengths.buffer)
    view.setUint32(0, label.length)
    view.setUint32(4, content.length)
    parts.push(lengths, label, content)
  }
  const canonical = new Uint8Array(parts.reduce((sum, part) => sum + part.length, 0))
  let offset = 0
  for (const part of parts) {
    canonical.set(part, offset)
    offset += part.length
  }
  const hash = await crypto.subtle.digest("SHA-256", canonical)
  return Array.from(new Uint8Array(hash), (byte) => byte.toString(16).padStart(2, "0")).join("")
}
