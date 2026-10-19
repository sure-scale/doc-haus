import { Database } from "bun:sqlite"
import { existsSync } from "node:fs"
import path from "node:path"
import { openLegalDbForWrite } from "./legal-db"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"

// Redaction support for the redact tool: the per-matter redaction log (what was
// removed, by whom, why — the work product a lawyer produces alongside a redacted
// document) and the raw OOXML metadata scrub that removes the redacted text from
// the package parts Docxodus cannot address (docProps title/creator/keywords,
// comment author attributes, textbox/field text the projection does not cover).
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

// The escape forms a needle can take inside serialized OOXML. Word escapes & < >
// in text nodes and additionally " ' in attribute values; we match every variant
// so "Smith & Jones" is caught both as raw text and as "Smith &amp; Jones".
function xmlVariants(text: string) {
  const textNode = text.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;")
  const attribute = textNode.replaceAll('"', "&quot;").replaceAll("'", "&apos;")
  return [...new Set([textNode, attribute])]
}

function scrubXmlPart(xml: string, needle: string, label: string) {
  const variants = xmlVariants(needle)
  const replacement = xmlVariants(label)[0]
  let hits = 0
  // Text nodes only (`>…<` segments) — replacing inside markup could corrupt the
  // part, and a contiguous needle can only carry content when it sits in a text
  // node or an attribute value.
  let out = xml.replace(/>([^<]*)</g, (whole, content: string) => {
    const scrubbed = variants.reduce((acc, v) => {
      hits += acc.split(v).length - 1
      return acc.replaceAll(v, replacement)
    }, content)
    return scrubbed === content ? whole : `>${scrubbed}<`
  })
  // Author identity attributes (comments, revision marks) carry names that a
  // name-redaction must reach; other attributes are structural and never hold
  // document content.
  out = out.replace(/(w:(?:author|initials)=")([^"]*)(")/g, (whole, pre: string, value: string, post: string) => {
    const scrubbed = variants.reduce((acc, v) => acc.replaceAll(v, replacement), value)
    if (scrubbed !== value) hits += 1
    return scrubbed === value ? whole : `${pre}${scrubbed}${post}`
  })
  return { xml: out, hits }
}

// Count the needle occurrences scrubXmlPart would remove, without mutating —
// the residue check the redact tool runs over every part after the Docxodus
// pass, and the pre-flight count shown to the user before they approve.
export function xmlPartOccurrences(xml: string, needle: string) {
  const variants = xmlVariants(needle)
  let count = 0
  for (const match of xml.matchAll(/>([^<]*)</g))
    for (const v of variants) count += match[1].split(v).length - 1
  for (const match of xml.matchAll(/(w:(?:author|initials)=")([^"]*)(")/g))
    for (const v of variants) count += match[2].split(v).length - 1
  return count
}

// Scrub the needle from every XML part of the DOCX package and report what was
// removed where. Returns the rewritten package, the per-part hit counts, and the
// binary parts (images, embedded objects) the scrub cannot reach — the caller
// surfaces those to the user instead of claiming they are clean.
export function scrubDocxPackage(bytes: Uint8Array, needle: string, label: string) {
  const parts = unzipSync(bytes)
  const hits: Record<string, number> = {}
  for (const [name, data] of Object.entries(parts)) {
    if (!name.endsWith(".xml") && !name.endsWith(".rels")) continue
    const scrubbed = scrubXmlPart(strFromU8(data), needle, label)
    if (scrubbed.hits > 0) {
      parts[name] = strToU8(scrubbed.xml)
      hits[name] = scrubbed.hits
    }
  }
  const unreachable = Object.keys(parts).filter((name) => /^word\/(media|embeddings)\//.test(name))
  return { bytes: zipSync(parts, { mtime: new Date(1980, 0, 1) }), hits, unreachable }
}

// Residue check across the whole package: any XML part where the needle still
// appears in a text node or author attribute. Empty means the package is clean.
export function docxPackageResidue(bytes: Uint8Array, needle: string) {
  const parts = unzipSync(bytes)
  return Object.entries(parts)
    .filter(([name]) => name.endsWith(".xml") || name.endsWith(".rels"))
    .filter(([, data]) => xmlPartOccurrences(strFromU8(data), needle) > 0)
    .map(([name]) => name)
}
