import { createHash } from "node:crypto"
import { DOMParser } from "@xmldom/xmldom"
import { strFromU8, unzipSync } from "fflate"
import { CHRONOLOGY_EXPORT_LABEL, CHRONOLOGY_EXPORT_ORIGIN, CHRONOLOGY_EXPORT_PROPERTY } from "../../../dochaus/lib/chronology-export"
import { openDb } from "./db"

export class DerivedChronologyExportError extends Error {
  constructor(readonly originalName?: string) {
    super("This file is a derived chronology snapshot. Add it as supplied chronology to review changes; it does not independently corroborate earlier events.")
  }
}

/** Renaming unchanged bytes must not turn our own output into corroboration.
 * Origin metadata also survives ordinary Word re-saves. Edited copies with
 * their metadata stripped are outside this bounded recognition contract. */
export function recognizeChronologyExport(directory: string, name: string, bytes: Uint8Array): { originalName?: string } | undefined {
  if (!name.toLowerCase().endsWith(".docx")) return
  const db = openDb(directory)
  let known: { name: string } | null
  try {
    known = db.query(`
      SELECT d.name FROM document_versions v
      JOIN versioned_documents d ON d.id=v.document_id
      WHERE v.sha256=? AND v.origin='generated' AND v.label=? LIMIT 1
    `).get(createHash("sha256").update(bytes).digest("hex"), CHRONOLOGY_EXPORT_LABEL) as { name: string } | null
  } finally {
    db.close()
  }
  if (known || hasChronologyExportOrigin(bytes)) return { originalName: known?.name }
}

export function hasChronologyExportOrigin(bytes: Uint8Array): boolean {
  try {
    const properties = unzipSync(bytes, {
      filter: (part) => part.name === "docProps/custom.xml" && part.originalSize <= 64 * 1024,
    })["docProps/custom.xml"]
    if (!properties) return false
    const fail = () => { throw new Error("Invalid document properties") }
    const xml = new DOMParser({ errorHandler: { warning: fail, error: fail, fatalError: fail } })
      .parseFromString(strFromU8(properties), "application/xml")
    const found = xml.getElementsByTagNameNS("http://schemas.openxmlformats.org/officeDocument/2006/custom-properties", "property")
    for (let index = 0; index < found.length; index++) {
      const property = found.item(index)!
      if (property.getAttribute("name") !== CHRONOLOGY_EXPORT_PROPERTY) continue
      for (let child = property.firstChild; child; child = child.nextSibling) {
        if (child.nodeType !== 1) continue
        const value = child as Element
        if (
          value.namespaceURI === "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes" &&
          ["lpwstr", "lpstr"].includes(value.localName ?? "") && value.textContent === CHRONOLOGY_EXPORT_ORIGIN
        ) return true
      }
    }
  } catch {
    // This is origin recognition, not DOCX validation. An unrelated malformed
    // document still receives the normal format/extraction error downstream.
  }
  return false
}
