import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync } from "fflate"

const wordNamespaces = new Set([
  "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  "http://purl.oclc.org/ooxml/wordprocessingml/main",
])
// Revision-session identifiers describe editing sessions, not the changes
// themselves. Keep unknown attributes and every actual change/format payload.
const sessionAttributes = new Set(["rsidR", "rsidRPr", "rsidDel", "rsidP", "rsidRDefault", "rsidTr", "rsidSect"])

function comparisonContent(name: string, bytes: Uint8Array): Uint8Array {
  if (!name.startsWith("word/") || !name.endsWith(".xml")) return bytes
  const text = strFromU8(bytes)
  if (!text.includes("rsid")) return bytes
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
    }
  }
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
