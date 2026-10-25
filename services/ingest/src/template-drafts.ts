import { Hono } from "hono"
import { randomUUID } from "node:crypto"
import { existsSync, lstatSync, mkdirSync, readFileSync, readdirSync, renameSync, rmSync, writeFileSync } from "node:fs"
import path from "node:path"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { applyDocxParagraphEdits, createDocxRevision, DocxTrackedRevisionsError, type DocxParagraphEdit } from "../../../dochaus/lib/docx-amendments"
import { withArtifactMutationLocksAsync } from "../../../dochaus/capabilities/artifacts/mutation-lock"
import { commitDocumentWrite } from "./atomic-document-write"
import { inlineDocumentDisposition } from "./content-disposition"
import { getDocumentRole, openDb } from "./db"
import { resolveMatterDocumentTarget, MatterDocumentIdentityError } from "./document-identity"
import { matterDir, workspaceRoot } from "./matter"
import { deleteTemplateDraftSessions, TemplateDraftSessionDeletionError } from "./template-draft-sessions"
import type { EngineRequestOptions } from "./matter-session"

export type TemplateDraftChange = {
  before_quote: string
  replacement: string
  paragraph_before?: string
  paragraph_after?: string
}

export type TemplateDraft = {
  id: string
  dir: string
  source: { name: string; matterId: string }
  createdAt: number
  revision: number
  name: string
  description: string
  status: "draft" | "published"
  publishedName?: string
  changes: TemplateDraftChange[]
  notes: string[]
}

class TemplateDraftError extends Error {
  constructor(message: string, readonly status: 400 | 404 | 409 | 422) { super(message) }
}

const WORD = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const DRAWING = "http://schemas.openxmlformats.org/drawingml/2006/main"
const UUID = /^[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}$/
const PROPOSAL = "proposal.docx"
const METADATA = "draft.json"

// A source-bearing conversation has its own directory beside the library. A
// shared library chat can never discover another conversation's client copy by
// listing its own directory. Only the reviewed proposal enters .templates.
export function templateDraftRoutes(resolveMatter = matterDir, root = workspaceRoot, engine: EngineRequestOptions = {}) {
  const app = new Hono()
  app.onError((error, c) => {
    if (error instanceof TemplateDraftError) return c.json({ error: error.message }, error.status)
    if (error instanceof TemplateDraftSessionDeletionError) return c.json({ error: error.message }, error.status)
    if (error instanceof MatterDocumentIdentityError)
      return c.json({ error: "Choose an original Word document in this matter, rather than a linked file." }, 400)
    if (error instanceof DocxTrackedRevisionsError)
      return c.json({ error: "Resolve the document's tracked changes in Word or doc.haus before creating a reusable template." }, 422)
    throw error
  })

  const draftDir = (id: string) => {
    if (!UUID.test(id)) throw new TemplateDraftError("Template draft was not found.", 404)
    const base = path.join(root(), ".template-drafts")
    const dir = path.join(base, id)
    requirePlainPath(base, true)
    requirePlainPath(dir, true)
    return dir
  }

  app.get("/template-drafts", (c) => {
    const base = path.join(root(), ".template-drafts")
    if (!existsSync(base)) return c.json({ drafts: [] })
    requirePlainPath(base, true)
    const drafts = readdirSync(base).filter(id => UUID.test(id)).flatMap(id => {
      // A damaged or interrupted draft never hides the other conversations.
      // Its directory stays available for recovery, matching matter listing.
      try { return [readDraft(draftDir(id))] } catch { return [] }
    }).sort((a, b) => b.createdAt - a.createdAt)
    return c.json({ drafts })
  })

  app.post("/matters/:id/template-drafts", async (c) => {
    const body = await c.req.json().catch(() => null) as { name?: unknown } | null
    const name = wordName(body?.name)
    if (!name) throw new TemplateDraftError("Choose one Word document in this matter, without a folder path.", 400)
    const matterId = c.req.param("id")
    if (!/^[a-z0-9][a-z0-9-]*$/.test(matterId)) throw new TemplateDraftError("Matter was not found.", 404)
    const directory = resolveMatter(matterId)
    requirePlainPath(directory, true)
    const matter = plainJson(path.join(directory, "matter.json")) as { id?: unknown } | null
    if (matter?.id !== matterId) throw new TemplateDraftError("Matter was not found.", 404)
    // Reject an alias before it can choose a lock outside this matter; resolve
    // again under the lock before reading the registered bytes.
    resolveMatterDocumentTarget(directory, name)
    const original = await withArtifactMutationLocksAsync(directory, [name], async () => {
      const target = resolveMatterDocumentTarget(directory, name)
      requirePlainPath(target.docPath)
      if (!existsSync(path.join(directory, ".dochaus", "legal.db")))
        throw new TemplateDraftError("Choose a document registered in this matter.", 404)
      const db = openDb(directory)
      try {
        if (!getDocumentRole(db, target.docPath)) throw new TemplateDraftError("Choose a document registered in this matter.", 404)
        return { name: target.name, bytes: await Bun.file(target.docPath).bytes() }
      } finally { db.close() }
    })

    // Validate and clean a copy before making any draft directory discoverable.
    const cleaned = cleanTemplateCopy(original.bytes)
    await project(cleaned.bytes)
    const id = randomUUID()
    const base = path.join(root(), ".template-drafts")
    mkdirSync(base, { recursive: true, mode: 0o700 })
    requirePlainPath(base, true)
    const dir = path.join(base, id)
    const staging = path.join(base, `.new-${id}`)
    const metadata: TemplateDraft = {
      id, dir, source: { name: original.name, matterId }, createdAt: Date.now(), revision: 0,
      name: "reusable-template.docx", description: "", status: "draft", changes: [], notes: cleaned.notes,
    }
    mkdirSync(staging, { mode: 0o700 })
    try {
      writeFileSync(path.join(staging, "source.docx"), original.bytes, { flag: "wx", mode: 0o600 })
      writeFileSync(path.join(staging, PROPOSAL), cleaned.bytes, { flag: "wx", mode: 0o600 })
      writeFileSync(path.join(staging, METADATA), JSON.stringify(metadata, null, 2), { flag: "wx", mode: 0o600 })
      renameSync(staging, dir)
    } finally { rmSync(staging, { recursive: true, force: true }) }
    return c.json(metadata, 201)
  })

  app.get("/template-drafts/:id", (c) => c.json(readDraft(draftDir(c.req.param("id")))))
  app.get("/template-drafts/:id/text", async (c) => {
    const dir = draftDir(c.req.param("id"))
    return withDraftMutationLock(dir, async () => {
      readDraft(dir)
      return c.json(await project(readProposal(dir)))
    })
  })
  app.get("/template-drafts/:id/content", async (c) => {
    const dir = draftDir(c.req.param("id"))
    const view = c.req.query("view") ?? "proposal"
    if (view !== "source" && view !== "proposal") throw new TemplateDraftError("Choose the source or proposed template.", 400)
    return withDraftMutationLock(dir, async () => {
      const metadata = readDraft(dir)
      const file = path.join(dir, view === "source" ? "source.docx" : PROPOSAL)
      requirePlainPath(file)
      return new Response(readFileSync(file), { headers: {
        "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "Content-Disposition": inlineDocumentDisposition(view === "source" ? metadata.source.name : metadata.name),
        "Cache-Control": "no-store",
      } })
    })
  })
  app.post("/template-drafts/:id/edits", async (c) => {
    const body = await c.req.json().catch(() => null) as Record<string, unknown> | null
    if (!body || !validRevision(body.revision) || !Array.isArray(body.edits) ||
      (body.notes !== undefined && (!Array.isArray(body.notes) || body.notes.some(note => typeof note !== "string"))))
      throw new TemplateDraftError("Supply the current revision and an array of exact paragraph edits.", 400)
    const name = body.name === undefined ? undefined : templateName(body.name)
    const description = body.description === undefined ? undefined : templateDescription(body.description)
    if (body.name !== undefined && !name) throw new TemplateDraftError("Use a Word template filename without a folder path.", 400)
    const edits = body.edits as DocxParagraphEdit[]
    const dir = draftDir(c.req.param("id"))
    return withDraftMutationLock(dir, async () => {
      const prior = readDraft(dir)
      requireCurrent(prior, body.revision)
      const bytes = readProposal(dir)
      const edited = edits.length ? await amend(bytes, edits) : undefined
      const metadata: TemplateDraft = {
        ...prior, revision: prior.revision + 1, name: name ?? prior.name, description: description ?? prior.description,
        changes: [...prior.changes, ...(edited?.report.amendments.map(change => ({
          before_quote: change.before_quote, replacement: change.replacement,
          paragraph_before: change.paragraph_before, paragraph_after: change.paragraph_after,
        })) ?? [])],
        notes: [...new Set([...prior.notes, ...(edited?.notes ?? []), ...((body.notes ?? []) as string[]).map(note => note.trim()).filter(Boolean)])],
      }
      commitDocumentWrite(path.join(dir, PROPOSAL), Buffer.from(edited?.bytes ?? bytes), () => writeJson(path.join(dir, METADATA), metadata))
      return c.json(metadata)
    })
  })
  app.post("/template-drafts/:id/publish", async (c) => {
    const body = await c.req.json().catch(() => null) as Record<string, unknown> | null
    const name = templateName(body?.name)
    const description = templateDescription(body?.description)
    if (!body || !validRevision(body.revision) || !name || !description)
      throw new TemplateDraftError("Supply the reviewed revision, a Word filename and a short template description.", 400)
    const dir = draftDir(c.req.param("id"))
    return withDraftMutationLock(dir, async () => {
      const prior = readDraft(dir)
      requireCurrent(prior, body.revision)
      if (prior.revision === 0) throw new TemplateDraftError("Prepare the reusable template in the conversation before saving it to the library.", 409)
      const bytes = readProposal(dir)
      await project(bytes)
      const library = path.join(root(), ".templates")
      mkdirSync(library, { recursive: true })
      requirePlainPath(library, true)
      return withArtifactMutationLocksAsync(library, [name, "manifest.json"], async () => {
        if (readdirSync(library).some(entry => entry.normalize("NFC").toLowerCase() === name.normalize("NFC").toLowerCase()))
          throw new TemplateDraftError(`A template named ${name} already exists. Choose a different name.`, 409)
        const manifest = path.join(library, "manifest.json")
        if (existsSync(manifest)) requirePlainPath(manifest)
        const originalManifest = existsSync(manifest) ? readFileSync(manifest) : undefined
        const descriptions = originalManifest ? JSON.parse(originalManifest.toString("utf8")) as Record<string, string> : {}
        const metadata: TemplateDraft = { ...prior, name, description, revision: prior.revision + 1, status: "published", publishedName: name }
        commitDocumentWrite(path.join(library, name), Buffer.from(bytes), () => {
          writeJson(manifest, { ...descriptions, [name]: description })
          try { writeJson(path.join(dir, METADATA), metadata) } catch (error) {
            if (originalManifest) writeFileSync(manifest, originalManifest)
            else rmSync(manifest, { force: true })
            throw error
          }
        })
        return c.json(metadata)
      })
    })
  })
  app.delete("/template-drafts/:id", async (c) => {
    const id = c.req.param("id")
    const dir = draftDir(id)
    return withDraftMutationLock(dir, async () => {
      readDraft(draftDir(id))
      await deleteTemplateDraftSessions(dir, engine)
      // Keep all draft bytes if the engine is unavailable or refuses deletion.
      // Revalidate after the native calls before removing this exact directory.
      readDraft(draftDir(id))
      rmSync(dir, { recursive: true })
      return c.json({ ok: true })
    })
  })
  return app
}

function withDraftMutationLock<T>(dir: string, run: () => Promise<T>) {
  // These are the same proposal/metadata artifact identities, with locks kept
  // in the surviving parent so a waiting edit cannot recreate a deleted draft.
  const id = path.basename(dir)
  return withArtifactMutationLocksAsync(path.dirname(dir), [path.join(id, METADATA), path.join(id, PROPOSAL)], run)
}

function wordName(value: unknown) {
  if (typeof value !== "string") return
  const name = value.trim()
  return name && !/[\\/\p{Cc}]/u.test(name) && path.basename(name) === name && path.extname(name).toLowerCase() === ".docx" ? name : undefined
}

function templateName(value: unknown) {
  const name = wordName(value)
  if (!name || name.startsWith(".") || Buffer.byteLength(name) > 255) return
  return name.slice(0, -5) + ".docx"
}

function templateDescription(value: unknown) {
  if (typeof value !== "string") throw new TemplateDraftError("Supply a short template description.", 400)
  const description = value.replace(/[\p{Cc}\s]+/gu, " ").trim()
  if (description.length > 1000) throw new TemplateDraftError("Keep the template description under 1,000 characters.", 400)
  return description
}

function validRevision(value: unknown): value is number { return typeof value === "number" && Number.isSafeInteger(value) && value >= 0 }

function requireCurrent(metadata: TemplateDraft, revision: unknown) {
  if (metadata.status === "published") throw new TemplateDraftError("This template has already been saved to the library.", 409)
  if (metadata.revision !== revision) throw new TemplateDraftError("The proposed template changed. Review the current version before continuing.", 409)
}

function requirePlainPath(file: string, directory = false) {
  if (!existsSync(file)) throw new TemplateDraftError("The selected document or template draft was not found.", 404)
  const stat = lstatSync(file)
  if (stat.isSymbolicLink() || (directory ? !stat.isDirectory() : !stat.isFile()))
    throw new TemplateDraftError("Linked files and folders cannot be used for this template conversation.", 400)
}

function plainJson(file: string): unknown {
  requirePlainPath(file)
  try { return JSON.parse(readFileSync(file, "utf8")) } catch { throw new TemplateDraftError("The selected document or template draft could not be read.", 409) }
}

function readDraft(dir: string) {
  const metadata = plainJson(path.join(dir, METADATA)) as TemplateDraft
  if (!metadata || metadata.id !== path.basename(dir) || metadata.dir !== dir || !validRevision(metadata.revision) ||
    !Number.isFinite(metadata.createdAt) || !["draft", "published"].includes(metadata.status) ||
    !metadata.source || typeof metadata.source.name !== "string" || typeof metadata.source.matterId !== "string" ||
    typeof metadata.name !== "string" || typeof metadata.description !== "string" ||
    !Array.isArray(metadata.changes) || !Array.isArray(metadata.notes) || metadata.notes.some(note => typeof note !== "string"))
    throw new TemplateDraftError("The template draft could not be read.", 409)
  return metadata
}

function readProposal(dir: string) {
  const file = path.join(dir, PROPOSAL)
  requirePlainPath(file)
  return new Uint8Array(readFileSync(file))
}

function writeJson(file: string, value: unknown) {
  const temporary = `${file}.${randomUUID()}.tmp`
  try {
    writeFileSync(temporary, JSON.stringify(value, null, 2), { flag: "wx", mode: 0o600 })
    renameSync(temporary, file)
  } finally { rmSync(temporary, { force: true }) }
}

async function project(bytes: Uint8Array) {
  try { return await createDocxRevision(bytes) } catch (error) {
    if (error instanceof DocxTrackedRevisionsError) throw error
    throw new TemplateDraftError("This document contains Word structure that cannot be safely converted to a reusable template. Save a clean Word copy and try again.", 422)
  }
}

async function amend(bytes: Uint8Array, edits: DocxParagraphEdit[]) {
  try {
    const edited = await applyDocxParagraphEdits(bytes, edits)
    const cleaned = cleanChangedLinks(edited.bytes, edits)
    if (cleaned.notes.length) {
      const before = await createDocxRevision(edited.bytes)
      const after = await createDocxRevision(cleaned.bytes)
      const visible = (revision: typeof before) => revision.paragraphs.map(({ text, scope, style_id, formatting, table }) => ({
        text, scope, style_id, table,
        // Removing the sensitive hyperlink is intentional; every displayed run
        // and every other formatting property must still match exactly.
        formatting: formatting.map(({ text, formatting: { hyperlinkUrl: _target, ...formatting } }) => ({ text, formatting })),
      }))
      if (JSON.stringify(visible(before)) !== JSON.stringify(visible(after))) throw new Error("link_cleanup_changed_visible_content")
    }
    return { ...edited, bytes: cleaned.bytes, notes: cleaned.notes }
  } catch {
    throw new TemplateDraftError("The proposed edits could not be safely applied. Read the current template and use its exact paragraph handles and wording; structural or field edits need a clean Word copy.", 409)
  }
}

// Comments and package properties can retain client data outside the visible
// body. Remove them from the proposal only. Bound data, live content and embedded
// documents require a deliberate clean Word copy rather than a lossy conversion.
function cleanTemplateCopy(source: Uint8Array) {
  let parts: Record<string, Uint8Array>
  try { parts = unzipSync(source) } catch { throw new TemplateDraftError("The selected file is not a readable Word document.", 422) }
  if (!parts["word/document.xml"]) throw new TemplateDraftError("The selected file is not a readable Word document.", 422)
  const removed = new Set<string>()
  const notes: string[] = []
  const documents = new Map<string, Document>()
  for (const [name, bytes] of Object.entries(parts)) {
    if (/(?:^|\/)(?:vbaProject\.bin|vbaData\.xml)|^word\/(?:embeddings|activeX)\/|^customUI\//i.test(name))
      throw new TemplateDraftError("This document contains macros or embedded content. Save a Word copy without macros, embedded files or controls before creating a template.", 422)
    if (name.endsWith(".xml") || name.endsWith(".rels")) {
      let document: Document
      try { document = readXml(bytes) }
      catch { throw new TemplateDraftError("The selected Word document contains unreadable XML. Save a clean Word copy and try again.", 422) }
      documents.set(name, document)
      const elements = Array.from(document.getElementsByTagName("*"))
      // A tracked numbering change has no accepted form doc.haus can choose.
      if (elements.some(element => element.namespaceURI === WORD && element.localName === "numberingChange"))
        throw new TemplateDraftError("This document has tracked numbering changes. Accept or reject them in Word before creating a reusable template.", 422)
      if (elements.some(element => ["dataBinding", "altChunk", "subDoc", "mailMerge", "OLEObject", "object", "control", "docPartObj", "docPartList"].includes(element.localName)))
        throw new TemplateDraftError("This document contains linked, bound or embedded content. Save a self-contained Word copy before creating a template.", 422)
      if (elements.some(element => element.namespaceURI === WORD &&
        (["ins", "del", "delText", "moveFrom", "moveTo", "cellIns", "cellDel", "cellMerge", "numberingChange"].includes(element.localName) ||
          /(?:PrChange|move(?:From|To)Range(?:Start|End))$/.test(element.localName))))
        throw new DocxTrackedRevisionsError()
    }
    if (/^docProps\//i.test(name) || /^customXml\//i.test(name) || /^word\/glossary\//i.test(name) || /^word\/(?:comments[^/]*|people)\.xml$/i.test(name)) removed.add(name)
  }
  if ([...removed].some(name => /^docProps\//i.test(name))) notes.push("Removed document properties from the proposed template.")
  if ([...removed].some(name => /^customXml\//i.test(name))) notes.push("Removed unbound custom XML data from the proposed template.")
  if ([...removed].some(name => /^word\/glossary\//i.test(name))) notes.push("Removed unused Word building blocks from the proposed template.")
  if ([...removed].some(name => /^word\/(?:comments[^/]*|people)\.xml$/i.test(name))) notes.push("Removed comments and review authors from the proposed template.")
  for (const name of [...removed]) {
    const rels = path.posix.join(path.posix.dirname(name), "_rels", `${path.posix.basename(name)}.rels`)
    if (parts[rels]) removed.add(rels)
  }
  for (const name of removed) delete parts[name]
  const bookmarks = new Map<string, string>()
  for (const [name, document] of documents) if (!removed.has(name))
    for (const element of Array.from(document.getElementsByTagNameNS(WORD, "bookmarkStart"))) {
      const name = element.getAttributeNS(WORD, "name")
      if (name && !bookmarks.has(name)) bookmarks.set(name, `TemplateBookmark${bookmarks.size + 1}`)
    }
  let removedDescriptors = false
  for (const [name, document] of documents) {
    if (removed.has(name)) continue
    let changed = cleanFields(document, bookmarks)
    for (const element of Array.from(document.getElementsByTagName("*"))) {
      const commentMarker = element.namespaceURI === WORD && ["commentRangeStart", "commentRangeEnd", "commentReference", "docVars", "attachedTemplate", "alias", "tag"].includes(element.localName)
      const removedType = element.localName === "Override" && removed.has((element.getAttribute("PartName") ?? "").replace(/^\//, ""))
      const target = element.getAttribute("Target") ?? ""
      const owner = name === "_rels/.rels" ? "" : path.posix.dirname(path.posix.dirname(name))
      const targetPart = target.startsWith("/") ? target.slice(1) : path.posix.normalize(path.posix.join(owner, target))
      const linkedRemoved = element.localName === "Relationship" && (removed.has(targetPart) || /\/attachedTemplate$/.test(element.getAttribute("Type") ?? ""))
      if (element.localName === "Relationship" && element.getAttribute("TargetMode") === "External" &&
        !/\/(?:hyperlink|attachedTemplate)$/.test(element.getAttribute("Type") ?? ""))
        throw new TemplateDraftError("This document links to external content. Save a self-contained Word copy before creating a template.", 422)
      if (commentMarker || removedType || linkedRemoved) { element.parentNode?.removeChild(element); changed = true; removedDescriptors ||= commentMarker }
      if (["docPr", "cNvPr"].includes(element.localName)) {
        for (const attribute of ["name", "descr", "title"]) if (element.hasAttribute(attribute)) {
          if (attribute === "name") element.setAttribute(attribute, "Template image")
          else element.removeAttribute(attribute)
          changed = removedDescriptors = true
        }
      } else for (const attribute of ["alt", "title"]) if (element.hasAttribute(attribute)) {
        element.removeAttribute(attribute)
        changed = removedDescriptors = true
      }
    }
    const removeXmlComments = (node: Node) => {
      for (const child of Array.from(node.childNodes ?? [])) {
        if (child.nodeType === child.COMMENT_NODE) { node.removeChild(child); changed = true }
        else removeXmlComments(child)
      }
    }
    removeXmlComments(document)
    if (changed) parts[name] = strToU8(new XMLSerializer().serializeToString(document))
  }
  if (bookmarks.size) notes.push("Generalised internal bookmark names while preserving their cross-references.")
  if (removedDescriptors) notes.push("Removed review markers, content-control labels and image description metadata; review image accessibility descriptions before reuse.")
  notes.push("Review the proposed wording, headers, footers, links and images for remaining client details before saving.")
  return { bytes: zipSync(parts), notes }
}

function readXml(bytes: Uint8Array) {
  const fail = () => { throw new Error("invalid XML") }
  return new DOMParser({ errorHandler: { warning: fail, error: fail, fatalError: fail } }).parseFromString(strFromU8(bytes), "application/xml")
}

// Page numbers and bookmark references remain functional. Other live fields can
// conceal client values or reintroduce them on refresh, so ask for an unlinked
// Word copy instead of presenting a misleadingly clean template.
function cleanFields(document: Document, bookmarks: Map<string, string>) {
  let changed = false
  const unsupported = () => { throw new TemplateDraftError("This document contains a live or unsupported Word field. In Word, unlink fields other than page numbers, bookmark cross-references and ordinary tables of contents, then try the clean copy.", 422) }
  const instruction = (original: string) => {
    const plain = original.replace(/\s+\\\*\s+(?:MERGEFORMAT|CHARFORMAT|Arabic|Roman|Alphabetic)\b/gi, "")
    if (/^\s*(?:PAGE|NUMPAGES|SECTION|SECTIONPAGES)\s*$/i.test(plain)) return original
    if (/^\s*TOC(?:\s+\\[hzu]|\s+\\o\s+"[1-9]-[1-9]")*\s*$/i.test(plain)) return original
    const reference = /^\s*(?:REF|PAGEREF|NOTEREF)\s+(?:"([^"]+)"|([^\s\\]+))((?:\s+\\[fhlnprtuw])*)\s*$/i.exec(plain)
    const bookmark = reference && bookmarks.get(reference[1] ?? reference[2])
    if (!bookmark) return unsupported()
    return original.replace(/^(\s*(?:REF|PAGEREF|NOTEREF)\s+)(?:"[^"]+"|[^\s\\]+)/i, `$1${bookmark}`)
  }
  const fields: { nodes: Element[]; complete: boolean }[] = []
  const finish = (field: { nodes: Element[]; complete: boolean }) => {
    if (field.complete) return
    const original = field.nodes.map(node => node.textContent ?? "").join("")
    const clean = instruction(original)
    if (original !== clean) {
      field.nodes.forEach((node, index) => { node.textContent = index === 0 ? clean : "" })
      changed = true
    }
    field.complete = true
  }
  for (const element of Array.from(document.getElementsByTagName("*"))) {
    if (element.namespaceURI !== WORD) continue
    if (element.localName === "bookmarkStart") {
      const name = element.getAttributeNS(WORD, "name")
      if (name && bookmarks.has(name)) { element.setAttributeNS(WORD, `${element.prefix ?? "w"}:name`, bookmarks.get(name)!); changed = true }
    }
    if (element.localName === "hyperlink" && element.hasAttributeNS(WORD, "anchor")) {
      const name = element.getAttributeNS(WORD, "anchor")
      if (name && bookmarks.has(name)) { element.setAttributeNS(WORD, `${element.prefix ?? "w"}:anchor`, bookmarks.get(name)!); changed = true }
      else unsupported()
    }
    if (element.localName === "fldSimple") {
      const original = element.getAttributeNS(WORD, "instr") ?? ""
      const clean = instruction(original)
      if (original !== clean) { element.setAttributeNS(WORD, `${element.prefix ?? "w"}:instr`, clean); changed = true }
    }
    if (element.localName === "fldChar") {
      const type = element.getAttributeNS(WORD, "fldCharType")
      if (type === "begin") fields.push({ nodes: [], complete: false })
      else if (type === "separate") { if (!fields.length) unsupported(); finish(fields.at(-1)!) }
      else if (type === "end") { if (!fields.length) unsupported(); finish(fields.pop()!) }
    }
    if (element.localName === "instrText") {
      const field = fields.at(-1)
      if (!field || field.complete) unsupported()
      field!.nodes.push(element)
    }
  }
  if (fields.length) unsupported()
  return changed
}

function cleanChangedLinks(bytes: Uint8Array, edits: DocxParagraphEdit[]) {
  const parts = unzipSync(bytes)
  const needles = edits.filter(edit => edit.before_quote !== edit.replacement).map(edit => edit.before_quote)
  let removed = 0
  for (const [part, bytes] of Object.entries(parts)) {
    if (!part.endsWith(".rels")) continue
    const document = readXml(bytes)
    const ids = new Set<string>()
    for (const element of Array.from(document.getElementsByTagName("*"))) {
      if (element.localName !== "Relationship" || !/\/hyperlink$/.test(element.getAttribute("Type") ?? "")) continue
      const target = element.getAttribute("Target") ?? ""
      const decoded = target.replace(/(?:%[0-9a-f]{2})+/gi, run => { try { return decodeURIComponent(run) } catch { return run } })
      if (!needles.some(needle => target.includes(needle) || decoded.includes(needle))) continue
      ids.add(element.getAttribute("Id") ?? "")
      element.parentNode?.removeChild(element)
      removed++
    }
    if (!ids.size) continue
    parts[part] = strToU8(new XMLSerializer().serializeToString(document))
    const owner = path.posix.join(path.posix.dirname(path.posix.dirname(part)), path.posix.basename(part, ".rels"))
    if (!parts[owner]) throw new Error("missing_hyperlink_owner")
    const content = readXml(parts[owner])
    for (const hyperlink of Array.from(content.getElementsByTagNameNS(WORD, "hyperlink"))) {
      if (!ids.has(hyperlink.getAttributeNS(RELATIONSHIP, "id") ?? "")) continue
      const parent = hyperlink.parentNode!
      while (hyperlink.firstChild) parent.insertBefore(hyperlink.firstChild, hyperlink)
      parent.removeChild(hyperlink)
    }
    // An image or shape may share the text link's relationship. Remove only
    // its click/hover action, retaining the drawing and its image relationship.
    for (const kind of ["hlinkClick", "hlinkHover"]) {
      for (const hyperlink of Array.from(content.getElementsByTagNameNS(DRAWING, kind))) {
        if (ids.has(hyperlink.getAttributeNS(RELATIONSHIP, "id") ?? "")) hyperlink.parentNode?.removeChild(hyperlink)
      }
    }
    parts[owner] = strToU8(new XMLSerializer().serializeToString(content))
  }
  return { bytes: removed ? zipSync(parts) : bytes, notes: removed ? [`Removed ${removed} hyperlink target(s) containing replaced source text; add any required links when drafting.`] : [] }
}
