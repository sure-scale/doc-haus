import { Database } from "bun:sqlite"
import { existsSync } from "node:fs"
import path from "node:path"
import { openLegalDbForWrite } from "./legal-db"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"

// Redaction support for the redact tool: the per-matter redaction log (what was
// removed, by whom, why — the work product a lawyer produces alongside a redacted
// document) and the package walk that removes the redacted text from what
// Docxodus cannot address (docProps, attribute values such as alt text, link
// targets and document variables, non-projected XML parts such as .vml), then
// re-scans every readable value and names the parts it cannot read.
//
// The log lives in the matter's index DB (.dochaus/legal.db) next to the redlines
// table; only the dochaus tools read and write it, so the DDL lives here alone.

export function recordRedaction(
  matterDir: string,
  row: {
    docPath: string
    docName: string
    redactedText: string
    label: string
    reason: string
    author: string
    occurrences: number
    metadataHits: number
  },
) {
  const db = openLegalDbForWrite(matterDir)
  try {
    db.run(`
      CREATE TABLE IF NOT EXISTS redactions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        doc_path TEXT NOT NULL,
        doc_name TEXT NOT NULL,
        redacted_text TEXT NOT NULL,
        label TEXT NOT NULL,
        reason TEXT NOT NULL,
        author TEXT NOT NULL,
        occurrences INTEGER NOT NULL,
        metadata_hits INTEGER NOT NULL,
        created_at INTEGER NOT NULL
      )
    `)
    const result = db.run(
      "INSERT INTO redactions (doc_path, doc_name, redacted_text, label, reason, author, occurrences, metadata_hits, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
      [row.docPath, row.docName, row.redactedText, row.label, row.reason, row.author, row.occurrences, row.metadataHits, Date.now()],
    )
    return Number(result.lastInsertRowid)
  } finally {
    db.close()
  }
}

export type RedactionRow = {
  id: number
  doc_path: string
  doc_name: string
  redacted_text: string
  label: string
  reason: string
  author: string
  occurrences: number
  metadata_hits: number
  created_at: number
}

export function listRedactions(matterDir: string, docPath?: string): RedactionRow[] {
  const dbFile = path.join(matterDir, ".dochaus", "legal.db")
  if (!existsSync(dbFile)) return []
  const db = new Database(dbFile, { readonly: true })
  const hasTable = db.query("SELECT name FROM sqlite_master WHERE type = 'table' AND name = 'redactions'").get()
  if (!hasTable) {
    db.close()
    return []
  }
  const rows = docPath
    ? db.query("SELECT * FROM redactions WHERE doc_path = ? ORDER BY created_at DESC, id DESC").all(docPath)
    : db.query("SELECT * FROM redactions ORDER BY created_at DESC, id DESC").all()
  db.close()
  return rows as RedactionRow[]
}

// The escape forms a needle can take inside serialized XML. Only the raw byte
// search of parts the walk cannot parse needs them; parsed parts are compared
// on decoded values.
function xmlVariants(text: string) {
  const textNode = text.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;")
  const attribute = textNode.replaceAll('"', "&quot;").replaceAll("'", "&apos;")
  return [...new Set([text, textNode, attribute])]
}

// One place where the needle survives: the part and what holds it there.
export type ResidueHit = { part: string; where: string }

// A readable value of a parsed part. `set` is absent where rewriting would break
// the package (part names, relationship ids, typed numbers), so a hit there
// stays residue and the redaction aborts instead of corrupting the file.
// `exact` marks a typed number, which holds the needle only by equality.
type Value = { text: string; where: string; set?: (value: string) => void; exact?: boolean }

const XMLNS = "http://www.w3.org/2000/xmlns/"
// A typed identifier or number: revision ids, colours, sizes, GUIDs. A needle
// inside a longer one is a different number, not a copy of the redacted value,
// so only an exact match counts, and it is never rewritten.
const NUMERIC_TOKEN = /^(?=.*\d)[-+.{}0-9A-Fa-f]+$/

// Parse a part as XML, or null when it is not readable XML (images, fonts,
// embedded binaries, malformed markup).
function readXml(data: Uint8Array) {
  let start = data[0] === 0xef && data[1] === 0xbb && data[2] === 0xbf ? 3 : 0
  while (start < data.length && [0x20, 0x09, 0x0a, 0x0d].includes(data[start])) start++
  if (data[start] !== 0x3c) return null
  const fail = (message: string) => {
    throw new Error(message)
  }
  try {
    return new DOMParser({ errorHandler: { warning: () => {}, error: fail, fatalError: fail } }).parseFromString(
      strFromU8(data),
      "application/xml",
    )
  } catch {
    return null
  }
}

// Every readable value of a parsed part in document order: character data,
// comments, processing instructions and decoded attribute values. Namespace
// declarations are schema identifiers, not content, and are skipped.
// [Content_Types].xml and relationship parts name other parts, so only an
// external relationship target (a link, a mailto address) is rewritable there.
function* values(part: string, node: Node): Generator<Value> {
  for (const child of Array.from(node.childNodes)) {
    if (child.nodeType === child.TEXT_NODE || child.nodeType === child.CDATA_SECTION_NODE || child.nodeType === child.COMMENT_NODE) {
      const data = child as CharacterData
      yield {
        text: data.data,
        where: child.nodeType === child.COMMENT_NODE ? "XML comment" : `text in <${node.nodeName}>`,
        set: (value) => data.replaceData(0, data.length, value),
      }
    } else if (child.nodeType === child.PROCESSING_INSTRUCTION_NODE) {
      yield { text: (child as ProcessingInstruction).data, where: `processing instruction <?${child.nodeName}?>` }
    } else if (child.nodeType === child.ELEMENT_NODE) {
      const element = child as Element
      for (const attr of Array.from(element.attributes)) {
        if (attr.namespaceURI === XMLNS || attr.name === "xmlns" || attr.prefix === "xmlns") continue
        const exact = NUMERIC_TOKEN.test(attr.value)
        const structural =
          exact ||
          part === "[Content_Types].xml" ||
          (part.endsWith(".rels") && !(attr.name === "Target" && element.getAttribute("TargetMode") === "External"))
        yield {
          text: attr.value,
          where: `attribute ${attr.name} on <${element.nodeName}>`,
          set: structural ? undefined : (value) => element.setAttribute(attr.name, value),
          exact,
        }
      }
      yield* values(part, element)
    }
  }
}

function contains(value: Value, needle: string) {
  return value.exact ? value.text === needle : value.text.includes(needle)
}

// Paragraph text joined across runs. A value split over differently formatted
// runs (w:p, or a:p in charts and shapes) is invisible per text node, so join
// the character data of each paragraph's leaf elements, not descending into
// nested paragraphs (text boxes), and report occurrences no single node holds.
function splitRunHits(node: Node, needle: string, out: string[] = []) {
  const count = (text: string) => text.split(needle).length - 1
  for (const child of Array.from(node.childNodes)) {
    if (child.nodeType !== child.ELEMENT_NODE) continue
    const element = child as Element
    if (element.localName === "p") {
      const pieces = leafText(element)
      if (count(pieces.join("")) > pieces.reduce((n, piece) => n + count(piece), 0))
        out.push(`text split across runs in <${element.nodeName}>`)
    }
    splitRunHits(element, needle, out)
  }
  return out
}

function leafText(element: Element, pieces: string[] = []) {
  const children = Array.from(element.childNodes)
  const leaf = !children.some((child) => child.nodeType === child.ELEMENT_NODE)
  for (const child of children) {
    if (child.nodeType === child.ELEMENT_NODE && (child as Element).localName !== "p") leafText(child as Element, pieces)
    else if (leaf && (child.nodeType === child.TEXT_NODE || child.nodeType === child.CDATA_SECTION_NODE))
      pieces.push((child as CharacterData).data)
  }
  return pieces
}

// Literal copies of the needle in a part the walk cannot parse: UTF-8 (raw and
// XML-escaped) and UTF-16LE, the encodings text takes inside binary parts. This
// proves only that no plain copy exists, not that the part is clean.
function bytesContain(data: Uint8Array, needle: string) {
  const haystack = Buffer.from(data.buffer, data.byteOffset, data.byteLength)
  const forms = [...xmlVariants(needle).map((variant) => Buffer.from(variant, "utf8")), Buffer.from(needle, "utf16le")]
  return forms.some((form) => haystack.includes(form))
}

// Scrub the needle from every readable value of the DOCX package and report
// what was removed where. Values the walk must not rewrite are left for the
// residue check to report.
export function scrubDocxPackage(bytes: Uint8Array, needle: string, label: string) {
  const parts = unzipSync(bytes)
  const hits: Record<string, number> = {}
  for (const [name, data] of Object.entries(parts)) {
    const document = readXml(data)
    if (!document) continue
    let count = 0
    for (const value of values(name, document)) {
      if (!value.set || !value.text.includes(needle)) continue
      count += value.text.split(needle).length - 1
      value.set(value.text.replaceAll(needle, label))
    }
    if (count > 0) {
      parts[name] = strToU8(new XMLSerializer().serializeToString(document))
      hits[name] = count
    }
  }
  return { bytes: zipSync(parts, { mtime: new Date(1980, 0, 1) }), hits }
}

// Residue check across the whole package. Every XML part (by content, whatever
// its extension: .xml, .rels, .vml) is parsed and every readable value checked,
// plus paragraph text joined across runs. Every other part is searched for a
// literal copy and listed as unread. No hits means the needle is absent from
// everything the scan reads, and only from that.
export function docxPackageResidue(bytes: Uint8Array, needle: string) {
  const hits: ResidueHit[] = []
  const read: string[] = []
  const unread: string[] = []
  for (const [part, data] of Object.entries(unzipSync(bytes))) {
    if (part.endsWith("/")) continue
    const document = readXml(data)
    if (!document) {
      if (bytesContain(data, needle)) hits.push({ part, where: "literal bytes in a part the scan cannot parse" })
      else unread.push(part)
      continue
    }
    read.push(part)
    const where = new Set([
      ...[...values(part, document)].filter((value) => contains(value, needle)).map((value) => value.where),
      ...splitRunHits(document, needle),
    ])
    for (const location of where) hits.push({ part, where: location })
  }
  return { hits, read, unread }
}
