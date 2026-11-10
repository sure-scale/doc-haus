import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, zipSync } from "fflate"
import path from "node:path"
import { parseZip } from "../../../dochaus/lib/zip"
import { readAvailableDocumentVersions, readDocumentVersions, readLinkedDocumentBranch, readVersionBytes, sha256Hex } from "./document-versions"

const CUSTOM = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
const VARIANT = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
const REL = "http://schemas.openxmlformats.org/package/2006/relationships"
const CUSTOM_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"
const CT = "http://schemas.openxmlformats.org/package/2006/content-types"
const WORD = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const WORD_STRICT = "http://purl.oclc.org/ooxml/wordprocessingml/main"
const MARKER = "doc.haus.version"
const UUID = /^[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}$/

export type DocumentReturnMarker = { documentId: string; versionId: string }
export type DocumentExchangeSource = { documentName?: string; attachmentId?: string; sender?: string; date?: string }
export type DocumentReturnSignal = "sha256" | "marker" | "rsid" | "text" | "headings" | "filename" | "email"
export type DocumentReturnCandidate = {
  documentId: string; documentName: string; versionId: string; versionNumber: number
  sha256: string; current: boolean; sent: boolean; signals: DocumentReturnSignal[]; score: number; reason: string
}
export type DocumentReturnInspection = {
  kind: "duplicate" | "matched" | "needs-confirmation" | "new-document"
  sha256: string; candidates: DocumentReturnCandidate[]; selected?: DocumentReturnCandidate
  reason: string; warnings: string[]; complete: boolean
}

function xml(bytes: Uint8Array | string) {
  const text = typeof bytes === "string" ? bytes : strFromU8(bytes)
  if (/<!DOCTYPE|<!ENTITY/i.test(text)) throw new Error("Word metadata cannot contain document type or entity declarations.")
  const fail = () => { throw new Error("The Word package contains malformed XML.") }
  const doc = new DOMParser({ errorHandler: { warning: fail, error: fail, fatalError: fail } }).parseFromString(text, "application/xml")
  if (!doc.documentElement || doc.getElementsByTagName("parsererror").length) fail()
  return doc
}
const serialize = (doc: Document) => strToU8(new XMLSerializer().serializeToString(doc))
const elements = (doc: Document | Element, ns: string, name: string) => Array.from(doc.getElementsByTagNameNS(ns, name))
const remove = (node: Node) => node.parentNode?.removeChild(node)

function partsOf(bytes: Uint8Array) {
  if (bytes.length > 128 * 1024 * 1024) throw new Error("The Word file exceeds the package size limit.")
  const archive = parseZip(bytes)
  if (archive.uncompressed > 256 * 1024 * 1024 || archive.entries.length > 1000) throw new Error("The Word file exceeds the package size limit.")
  const parts: Record<string, Uint8Array> = Object.create(null)
  for (const entry of archive.entries) {
    if (entry.encrypted || entry.unsafePath || /[\\\0]/.test(entry.filename) || entry.filename.split("/").some(p => !p || p === ".") || parts[entry.filename])
      throw new Error("The Word package has an unsafe or duplicate entry.")
    parts[entry.filename] = entry.content()
  }
  if (!parts["word/document.xml"] || !parts["[Content_Types].xml"] || !parts["_rels/.rels"]) throw new Error("A readable Word package is required.")
  return parts
}

function customParts(parts: Record<string, Uint8Array>) {
  const rels = xml(parts["_rels/.rels"]!)
  const links = elements(rels, REL, "Relationship").filter(e => e.getAttribute("Type") === CUSTOM_REL)
  if (links.length > 1) throw new Error("The Word file contains ambiguous custom properties.")
  const link = links[0]
  const target = link?.getAttribute("Target")?.replace(/^\//, "") ?? "docProps/custom.xml"
  if (link?.getAttribute("TargetMode") === "External" || !target || /[\\:\0]/.test(target) || target.split("/").some(p => !p || p === "." || p === ".."))
    throw new Error("The Word custom-properties target is invalid.")
  // An unreferenced part at the standard location is not silently adopted.
  if (!link && parts[target]) throw new Error("The Word custom properties have no package relationship.")
  const custom = parts[target] ? xml(parts[target]) : xml(`<Properties xmlns="${CUSTOM}" xmlns:vt="${VARIANT}"/>`)
  if (custom.documentElement.namespaceURI !== CUSTOM || custom.documentElement.localName !== "Properties") throw new Error("The Word custom properties are invalid.")
  return { rels, link, target, custom }
}

/** Opaque source identity only: never a credential, installation address or
 * permission to write. It is stable for a version, so retries export the same bytes. */
export function stampDocumentReturnMarker(bytes: Uint8Array, marker: DocumentReturnMarker | null) {
  if (marker && (!UUID.test(marker.documentId) || !UUID.test(marker.versionId))) throw new Error("Invalid document version marker.")
  const parts = partsOf(bytes)
  const { rels, link, target, custom } = customParts(parts)
  const owned = elements(custom, CUSTOM, "property").filter(e => e.getAttribute("name") === MARKER)
  if (!marker && !owned.length) return bytes
  for (const node of owned) remove(node)
  if (marker) {
    const pids = elements(custom, CUSTOM, "property").map(e => Number(e.getAttribute("pid"))).filter(Number.isSafeInteger)
    const property = custom.createElementNS(CUSTOM, "property")
    property.setAttribute("fmtid", "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}")
    property.setAttribute("pid", String(Math.max(1, ...pids) + 1))
    property.setAttribute("name", MARKER)
    const value = custom.createElementNS(VARIANT, "vt:lpwstr")
    value.appendChild(custom.createTextNode(JSON.stringify(marker)))
    property.appendChild(value)
    custom.documentElement.appendChild(property)
  }
  if (!link) {
    const ids = new Set(elements(rels, REL, "Relationship").map(e => e.getAttribute("Id")))
    let n = 1
    while (ids.has(`rIdDocHausVersion${n}`)) n++
    const relationship = rels.createElementNS(REL, "Relationship")
    for (const [name, value] of Object.entries({ Id: `rIdDocHausVersion${n}`, Type: CUSTOM_REL, Target: target })) relationship.setAttribute(name, value)
    rels.documentElement.appendChild(relationship)
    parts["_rels/.rels"] = serialize(rels)
  }
  const types = xml(parts["[Content_Types].xml"]!)
  if (!elements(types, CT, "Override").some(e => e.getAttribute("PartName") === `/${target}`)) {
    const override = types.createElementNS(CT, "Override")
    override.setAttribute("PartName", `/${target}`)
    override.setAttribute("ContentType", "application/vnd.openxmlformats-officedocument.custom-properties+xml")
    types.documentElement.appendChild(override)
    parts["[Content_Types].xml"] = serialize(types)
  }
  parts[target] = serialize(custom)
  return zipSync(Object.fromEntries(Object.entries(parts).map(([name, data]) => [name, [data, { mtime: new Date("2000-01-01T00:00:00Z") }]])))
}

export function readDocumentReturnMarker(bytes: Uint8Array): DocumentReturnMarker | undefined {
  const { custom } = customParts(partsOf(bytes))
  const owned = elements(custom, CUSTOM, "property").filter(e => e.getAttribute("name") === MARKER)
  if (!owned.length) return
  if (owned.length !== 1) throw new Error("The Word file contains ambiguous doc.haus version markers.")
  const values = elements(owned[0]!, VARIANT, "lpwstr")
  if (values.length !== 1 || (values[0]!.textContent?.length ?? 0) > 250) throw new Error("The Word version marker is invalid.")
  const value = JSON.parse(values[0]!.textContent ?? "null")
  if (!value || !UUID.test(value.documentId) || !UUID.test(value.versionId) || Object.keys(value).some(k => !["documentId", "versionId"].includes(k)))
    throw new Error("The Word version marker is invalid.")
  return { documentId: value.documentId, versionId: value.versionId }
}

function wordElements(doc: Document | Element, name: string) { return [...elements(doc, WORD, name), ...elements(doc, WORD_STRICT, name)] }
const attr = (element: Element, name: string) => element.getAttributeNS(WORD, name) ?? element.getAttributeNS(WORD_STRICT, name)
const normalized = (text: string) => text.normalize("NFKC").toLocaleLowerCase("en").replace(/[^\p{L}\p{N}]+/gu, " ").trim()
function signalsOf(bytes: Uint8Array) {
  const parts = partsOf(bytes), doc = xml(parts["word/document.xml"]!)
  const paragraphs = wordElements(doc, "p").map(p => {
    const text = wordElements(p, "t").filter(t => {
      for (let n: Node | null = t.parentNode; n && n !== p; n = n.parentNode)
        if (n.nodeType === 1 && [WORD, WORD_STRICT].includes((n as Element).namespaceURI ?? "") && ["del", "moveFrom"].includes((n as Element).localName)) return false
      return true
    }).map(t => t.textContent ?? "").join("")
    const style = wordElements(p, "pStyle")[0]
    return { text: normalized(text), heading: !!style && /^heading\s*\d+$/i.test(attr(style, "val") ?? "") }
  })
  const text = paragraphs.map(p => p.text).join(" ").trim()
  const words = text.split(" ").filter(Boolean)
  const tokens = new Set(words.length < 3 ? words : words.slice(0, -2).map((word, i) => `${word} ${words[i + 1]} ${words[i + 2]}`))
  const headings = new Set(paragraphs.filter(p => p.heading).map(p => p.text).filter(Boolean))
  const settings = parts["word/settings.xml"] ? xml(parts["word/settings.xml"]!) : undefined
  const roots = settings ? wordElements(settings, "rsidRoot").map(e => attr(e, "val")?.toLowerCase()).filter((v): v is string => !!v && /^[a-f0-9]{8}$/.test(v)) : []
  const rsids = new Set(settings ? wordElements(settings, "rsid").map(e => attr(e, "val")?.toLowerCase()).filter((v): v is string => !!v && /^[a-f0-9]{8}$/.test(v)) : [])
  return { tokens, headings, root: roots.length === 1 ? roots[0] : undefined, rsids }
}
const dice = (a: Set<string>, b: Set<string>) => a.size && b.size ? 2 * [...a].filter(v => b.has(v)).length / (a.size + b.size) : 0
const filename = (name: string) => normalized(path.basename(name, path.extname(name)).replace(/\b[A-Z]{1,3}\b/g, " ")
  .replace(/\b(?:draft|final|signed|clean|redline|markup|version|rev|v)\s*\d*\b|\b\d{1,4}(?:[-_.]\d{1,4})*\b/gi, " "))

/** Retrieval only. Similar text, Word session IDs, email and names propose
 * candidates; they never decide that two passages have the same legal meaning.
 * A registry-backed opaque marker or byte-identical version can identify a file.
 * Every other candidate requires explicit confirmation (or a separately consented
 * focused judgment job), even when its text similarity is very high. */
export async function inspectDocumentReturn(matterDir: string, input: { name: string; bytes: Uint8Array; source?: DocumentExchangeSource }): Promise<DocumentReturnInspection> {
  const hash = sha256Hex(input.bytes), warnings: string[] = []
  const rows = readAvailableDocumentVersions(matterDir)
  const histories = new Map<string, ReturnType<typeof readDocumentVersions>>()
  const history = (name: string) => histories.get(name) ?? histories.set(name, readDocumentVersions(matterDir, name)).get(name)!
  const candidate = (row: typeof rows[number], signals: DocumentReturnSignal[], score: number, reason: string): DocumentReturnCandidate => ({
    documentId: row.document_id, documentName: row.name, versionId: row.id, versionNumber: row.number,
    sha256: row.sha256, current: !!row.current, sent: history(row.name).events.some(e => e.kind === "sent" && e.version_id === row.id), signals, score, reason,
  })
  const exact = rows.filter(row => row.sha256 === hash && readVersionBytes(matterDir, row))
  if (exact.length) {
    const candidates = exact.map(row => candidate(row, ["sha256"], 1, "The complete file bytes are identical to this retained version."))
    const documents = new Set(candidates.map(c => c.documentId))
    return { kind: "duplicate", sha256: hash, candidates, ...(documents.size === 1 ? { selected: candidates[0] } : {}),
      reason: documents.size === 1 ? "This file is already retained. Nothing new needs to be saved." : "Identical bytes occur in more than one document. Choose the intended document.", warnings, complete: true }
  }
  if (path.extname(input.name).toLowerCase() !== ".docx") return { kind: "new-document", sha256: hash, candidates: [], reason: "No identical retained file was found.", warnings, complete: true }
  let markerUnavailable = false
  try {
    const marker = readDocumentReturnMarker(input.bytes)
    if (marker) {
      const row = rows.find(r => r.id === marker.versionId && r.document_id === marker.documentId)
      if (row && readVersionBytes(matterDir, row)) {
        const correctedReceipts = history(row.name).events.filter(e => e.kind === "received" && e.version_id === row.id && e.detail?.association)
        if (correctedReceipts.length) {
          // A marker keeps the immutable physical artifact identity. A later
          // association correction must not turn it into proof of ancestry for
          // either the old owner or the corrected target's base.
          const candidates: DocumentReturnCandidate[] = []
          for (const receipt of correctedReceipts) {
            const branch = readLinkedDocumentBranch(matterDir, receipt.id)
            const base = branch?.available ? rows.find(r => r.id === branch.baseVersionId && r.document_id === branch.targetDocumentId) : undefined
            if (branch && base) {
              if (!candidates.some(c => c.versionId === base.id)) candidates.push(candidate(base, [], 0,
                `The marked artifact was explicitly linked to ${branch.targetName}, base version ${base.number}. This is the corrected association, not proof that this later copy descends from that base. Confirm the document and base.`))
            } else warnings.push("The marked artifact has a corrected document association whose target or base is unavailable. Choose a live document and base explicitly.")
          }
          candidates.push(candidate(row, ["marker"], 1,
            `The marker identifies the retained physical source ${row.name}, version ${row.number}. Its received document association was corrected; the marker cannot establish this later copy's document or base.`))
          return { kind: "needs-confirmation", sha256: hash, candidates,
            reason: "The marker identifies a retained artifact whose document association was corrected. Confirm the intended document and base; its original physical owner is not an automatic match.", warnings, complete: true }
        }
        const selected = candidate(row, ["marker"], 1, "The opaque export marker identifies this retained document and base version.")
        return { kind: "matched", sha256: hash, candidates: [selected], selected, reason: selected.reason, warnings, complete: true }
      }
      markerUnavailable = true
      warnings.push("The export marker names an unavailable version. It cannot establish this return's base.")
    }
  } catch { markerUnavailable = true; warnings.push("The Word export marker could not be read safely. Choose the document and base explicitly.") }
  let incoming: ReturnType<typeof signalsOf>
  try { incoming = signalsOf(input.bytes) }
  catch { return { kind: "needs-confirmation", sha256: hash, candidates: [], reason: "This Word file could not be inspected safely. Choose its document explicitly.", warnings, complete: false } }
  const candidates: DocumentReturnCandidate[] = []
  let complete = true
  for (const row of rows.filter(r => r.ext === ".docx")) {
    const bytes = readVersionBytes(matterDir, row)
    if (!bytes) { complete = false; continue }
    try {
      const prior = signalsOf(bytes), text = dice(incoming.tokens, prior.tokens), headings = dice(incoming.headings, prior.headings)
      const rsid = !!incoming.root && incoming.root === prior.root && [...incoming.rsids].some(id => prior.rsids.has(id))
      if (!rsid && text < 0.15 && headings < 0.5) continue
      const signals: DocumentReturnSignal[] = [...rsid ? ["rsid" as const] : [], ...text >= 0.15 ? ["text" as const] : [], ...headings >= 0.5 ? ["headings" as const] : []]
      const sameName = !!filename(input.name) && filename(input.name) === filename(row.name)
      if (sameName) signals.push("filename")
      const sender = input.source?.sender?.trim().toLowerCase()
      const sameSender = !!sender && history(row.name).events.some(e => e.kind === "received" && (e.detail?.source as DocumentExchangeSource | undefined)?.sender?.trim().toLowerCase() === sender)
      if (sameSender) signals.push("email")
      const score = Math.min(0.99, text * 0.8 + headings * 0.1 + (rsid ? 0.08 : 0) + (sameName ? 0.01 : 0) + (sameSender ? 0.005 : 0))
      candidates.push(candidate(row, signals, Math.round(score * 10000) / 10000,
        `${rsid ? "Shared Word editing-session metadata; " : ""}${Math.round(text * 100)}% text overlap${headings ? `, ${Math.round(headings * 100)}% heading overlap` : ""}. Confirm the document and base; similarity does not establish ancestry.`))
    } catch { complete = false }
  }
  // Retrieve likely documents by their best eligible comparison base, then
  // prefer that document's closest sent version over an unsent internal draft.
  // This is a suggestion only; similarity never proves the return's ancestry.
  const documents = new Map<string, DocumentReturnCandidate[]>()
  for (const candidate of candidates) documents.set(candidate.documentId, [...(documents.get(candidate.documentId) ?? []), candidate])
  const groups = [...documents.values()].map(group => group.sort((a, b) => Number(b.sent) - Number(a.sent) || b.score - a.score || b.versionNumber - a.versionNumber))
    .sort((a, b) => b[0]!.score - a[0]!.score || a[0]!.documentName.localeCompare(b[0]!.documentName))
  candidates.splice(0, candidates.length, ...groups.flat())
  if (!complete) warnings.push("Some retained Word versions could not be inspected. Candidate coverage is incomplete.")
  return { kind: candidates.length || markerUnavailable || !complete ? "needs-confirmation" : "new-document", sha256: hash, candidates,
    reason: candidates.length ? "Possible earlier versions were found. Confirm which document and base this file belongs to." : markerUnavailable ? "Choose a document and base, or explicitly keep this as a new document." : "No supported earlier-version match was found.", warnings, complete }
}
