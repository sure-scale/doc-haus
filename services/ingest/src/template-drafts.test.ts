import { expect, test } from "bun:test"
import { existsSync, mkdtempSync, mkdirSync, readFileSync, readdirSync, rmSync, symlinkSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { Document, ExternalHyperlink, Footer, Header, HeadingLevel, ImageRun, Packer, Paragraph, Table, TableCell, TableRow, TextRun } from "docx"
import { createDocxRevision, type DocxParagraphEdit, type DocxRevisionPacket } from "../../../dochaus/lib/docx-amendments"
import { openDb, upsertDeliverable } from "./db"
import { templateDraftRoutes, type TemplateDraft } from "./template-drafts"
import type { EngineRequestOptions } from "./matter-session"

function fixture(engine: EngineRequestOptions = {}) {
  const temporary = mkdtempSync(path.join(tmpdir(), "dochaus-template-drafts-"))
  const workspace = path.join(temporary, "workspace")
  mkdirSync(workspace)
  const app = templateDraftRoutes(id => path.join(workspace, id), () => workspace, engine)
  const request = (route: string, method = "GET", body?: unknown) => app.request(`http://localhost${route}`, {
    method, ...(body === undefined ? {} : { headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) }),
  })
  function register(id: string, bytes: Uint8Array, name = "Selected.docx", source = false) {
    const dir = path.join(workspace, id)
    mkdirSync(dir, { recursive: true })
    writeFileSync(path.join(dir, "matter.json"), JSON.stringify({ id, title: "Synthetic matter" }))
    const file = path.join(dir, name)
    writeFileSync(file, bytes)
    const db = openDb(dir)
    try {
      upsertDeliverable(db, file, name, Date.now())
      if (source) db.run("UPDATE documents SET role = 'source' WHERE doc_path = ?", [file])
    } finally { db.close() }
    return { dir, file }
  }
  return { workspace, request, register, close: () => rmSync(temporary, { recursive: true, force: true }) }
}

async function sourceDocument() {
  const bytes = new Uint8Array(await Packer.toBuffer(new Document({
    creator: "Synthetic Reviewer", title: "Acme Example Ltd private terms",
    sections: [{
      headers: { default: new Header({ children: [new Paragraph("Acme Example Ltd engagement")] }) },
      footers: { default: new Footer({ children: [new Paragraph("House wording remains intact")] }) },
      children: [
        new Paragraph({ heading: HeadingLevel.HEADING_1, text: "1. Engagement" }),
        new Paragraph({ children: [new TextRun("Client: "), new TextRun({ text: "Acme Example Ltd", bold: true }), new TextRun(" instructs the adviser.")] }),
        new Table({ rows: [
          new TableRow({ children: [new TableCell({ children: [new Paragraph("Party")] }), new TableCell({ children: [new Paragraph("Procedure")] })] }),
          new TableRow({ children: [new TableCell({ children: [new Paragraph("Acme Example Ltd")] }), new TableCell({ children: [new Paragraph("The specified process uses 30 days.")] })] }),
        ] }),
      ],
    }],
  })))
  const parts = unzipSync(bytes)
  // All names, authors and document content here are reviewed synthetic data.
  parts["customXml/item1.xml"] = strToU8('<private-data>Acme Example Ltd confidential note</private-data>')
  parts["word/comments.xml"] = strToU8('<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="7" w:author="Synthetic Reviewer"><w:p><w:r><w:t>Private negotiation comment</w:t></w:r></w:p></w:comment></w:comments>')
  parts["word/document.xml"] = strToU8(strFromU8(parts["word/document.xml"]!).replace("</w:p>", '<w:commentRangeStart w:id="7"/><w:commentRangeEnd w:id="7"/><w:r><w:commentReference w:id="7"/></w:r></w:p>'))
  parts["word/_rels/document.xml.rels"] = strToU8(strFromU8(parts["word/_rels/document.xml.rels"]!).replace("</Relationships>", '<Relationship Id="rIdSyntheticComment" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/><Relationship Id="rIdSyntheticData" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item1.xml"/></Relationships>'))
  parts["[Content_Types].xml"] = strToU8(strFromU8(parts["[Content_Types].xml"]!).replace("</Types>", '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>'))
  return zipSync(parts)
}

async function capture(f: ReturnType<typeof fixture>, id = "matter-a") {
  const response = await f.request(`/matters/${id}/template-drafts`, "POST", { name: " Selected.docx " })
  expect(response.status).toBe(201)
  return await response.json() as TemplateDraft
}

async function candidate(f: ReturnType<typeof fixture>, draft: TemplateDraft) {
  return new Uint8Array(await (await f.request(`/template-drafts/${draft.id}/content`)).arrayBuffer())
}

function xmlContent(bytes: Uint8Array) {
  return new XMLSerializer().serializeToString(new DOMParser().parseFromString(strFromU8(bytes), "application/xml").documentElement)
}

test("captures one immutable source, preserves styled Word/table structure and saves only after review", async () => {
  const f = fixture()
  try {
    const original = await sourceDocument()
    const matter = f.register("matter-a", original, "Selected.docx", true)
    const draft = await capture(f)
    expect(draft).toMatchObject({ source: { name: "Selected.docx", matterId: "matter-a" }, status: "draft", revision: 0, changes: [] })
    expect(draft.dir).toBe(path.join(f.workspace, ".template-drafts", draft.id))
    expect(readdirSync(f.workspace)).not.toContain(".templates")
    const source = await f.request(`/template-drafts/${draft.id}/content?view=source`)
    expect(new Uint8Array(await source.arrayBuffer())).toEqual(original)
    const clean = await candidate(f, draft)
    const cleanedParts = unzipSync(clean)
    expect(Object.keys(cleanedParts).some(name => /^(docProps|customXml)\//.test(name))).toBe(false)
    expect(cleanedParts["word/comments.xml"]).toBeUndefined()
    expect(strFromU8(cleanedParts["word/document.xml"]!)).not.toContain("commentReference")
    expect(strFromU8(cleanedParts["word/_rels/document.xml.rels"]!)).not.toContain("rIdSynthetic")
    expect(cleanedParts["word/styles.xml"]).toEqual(unzipSync(original)["word/styles.xml"])
    expect(readFileSync(matter.file)).toEqual(Buffer.from(original))
    expect(draft.notes.join(" ")).toContain("Removed comments")
    const revision = await (await f.request(`/template-drafts/${draft.id}/text`)).json() as DocxRevisionPacket
    const matching = revision.paragraphs.filter(paragraph => paragraph.text.includes("Acme Example Ltd"))
    expect(matching).toHaveLength(3)
    const edits: DocxParagraphEdit[] = matching.map((paragraph, index) => ({
      id: `party-${index}`, paragraph_handle: paragraph.paragraph_handle, occurrence: 0,
      before_quote: "Acme Example Ltd", replacement: `[insert client detail ${index + 1}]`,
    }))
    const response = await f.request(`/template-drafts/${draft.id}/edits`, "POST", {
      revision: 0, name: "Firm terms.docx", description: "Reusable engagement terms", edits,
    })
    expect(response.status).toBe(200)
    const proposed = await response.json() as TemplateDraft
    expect(proposed.revision).toBe(1)
    expect(proposed.changes).toHaveLength(3)
    expect(proposed.changes[0]).toMatchObject({ before_quote: "Acme Example Ltd", replacement: "[insert client detail 1]" })
    const preview = await candidate(f, draft)
    const after = await createDocxRevision(preview)
    expect(after.paragraphs.some(paragraph => paragraph.text.includes("Acme Example Ltd"))).toBe(false)
    expect(after.paragraphs.find(paragraph => paragraph.text.includes("30 days"))?.text).toBe("The specified process uses 30 days.")
    expect(xmlContent(unzipSync(preview)["word/styles.xml"]!)).toBe(xmlContent(cleanedParts["word/styles.xml"]!))
    expect(strFromU8(unzipSync(preview)["word/document.xml"]!).match(/<w:tc\b/g)).toHaveLength(4)
    expect(readdirSync(f.workspace)).not.toContain(".templates")
    // A later matter revision neither changes the captured source nor becomes
    // the candidate published from this conversation.
    writeFileSync(matter.file, "Later externally saved matter version")
    const published = await f.request(`/template-drafts/${draft.id}/publish`, "POST", {
      revision: proposed.revision, name: proposed.name, description: proposed.description,
    })
    expect(published.status).toBe(200)
    expect(await published.json()).toMatchObject({ status: "published", publishedName: "Firm terms.docx", revision: 2 })
    expect(readFileSync(path.join(f.workspace, ".templates", proposed.name))).toEqual(Buffer.from(preview))
    expect(JSON.parse(readFileSync(path.join(f.workspace, ".templates", "manifest.json"), "utf8"))).toEqual({ "Firm terms.docx": "Reusable engagement terms" })
    expect(readFileSync(matter.file, "utf8")).toBe("Later externally saved matter version")
    expect(new Uint8Array(await (await f.request(`/template-drafts/${draft.id}/content?view=source`)).arrayBuffer())).toEqual(original)
  } finally { f.close() }
}, 60000)

test("refuses stale or invalid edits, an unprepared save and a colliding library name without mutation", async () => {
  const f = fixture()
  try {
    const original = await sourceDocument()
    const matter = f.register("matter-a", original)
    const draft = await capture(f)
    const before = await candidate(f, draft)
    expect((await f.request(`/template-drafts/${draft.id}/publish`, "POST", { revision: 0, name: "Existing.docx", description: "General terms" })).status).toBe(409)
    const prepared = await f.request(`/template-drafts/${draft.id}/edits`, "POST", { revision: 0, name: "Existing.docx", description: "General terms", edits: [] })
    expect(prepared.status).toBe(200)
    const metadata = await prepared.json() as TemplateDraft
    expect((await f.request(`/template-drafts/${draft.id}/edits`, "POST", { revision: 0, name: "Stale.docx", edits: [] })).status).toBe(409)
    expect((await f.request(`/template-drafts/${draft.id}/edits`, "POST", { revision: 1, edits: [{ id: "bad", paragraph_handle: "foreign-handle", occurrence: 0, before_quote: "Acme", replacement: "[insert client]" }] })).status).toBe(409)
    expect(await (await f.request(`/template-drafts/${draft.id}`)).json()).toEqual(metadata)
    expect(await candidate(f, draft)).toEqual(before)
    const library = path.join(f.workspace, ".templates")
    mkdirSync(library)
    writeFileSync(path.join(library, "existing.docx"), "Existing library bytes")
    writeFileSync(path.join(library, "manifest.json"), JSON.stringify({ "existing.docx": "Original description" }))
    expect((await f.request(`/template-drafts/${draft.id}/publish`, "POST", { revision: 1, name: "Existing.docx", description: "General terms" })).status).toBe(409)
    expect(readFileSync(path.join(library, "existing.docx"), "utf8")).toBe("Existing library bytes")
    expect(await (await f.request(`/template-drafts/${draft.id}`)).json()).toEqual(metadata)
    expect(readFileSync(matter.file)).toEqual(Buffer.from(original))
    expect((await f.request(`/template-drafts/${draft.id}/publish`, "POST", { revision: 1, name: "Different.docx", description: "General terms" })).status).toBe(200)
  } finally { f.close() }
}, 60000)

test("rejects paths, linked or unregistered sources and lists only recoverable draft conversations", async () => {
  const f = fixture()
  try {
    const bytes = await sourceDocument()
    const matter = f.register("matter-a", bytes)
    const other = f.register("matter-b", bytes, "Other.docx")
    writeFileSync(path.join(matter.dir, "Loose.docx"), bytes)
    symlinkSync(other.file, path.join(matter.dir, "Linked.docx"))
    for (const name of ["../matter-b/Other.docx", other.file, "folder\\Other.docx", "Other.pdf"])
      expect((await f.request("/matters/matter-a/template-drafts", "POST", { name })).status).toBe(400)
    expect((await f.request("/matters/matter-a/template-drafts", "POST", { name: "Linked.docx" })).status).toBe(400)
    expect((await f.request("/matters/matter-a/template-drafts", "POST", { name: "Loose.docx" })).status).toBe(404)
    expect((await f.request("/matters/matter-a/template-drafts", "POST", { name: "Other.docx" })).status).toBe(404)
    expect(await (await f.request("/template-drafts")).json()).toEqual({ drafts: [] })
    const first = await capture(f)
    const second = await capture(f)
    const base = path.join(f.workspace, ".template-drafts")
    mkdirSync(path.join(base, ".new-incomplete"))
    const broken = path.join(base, "00000000-0000-0000-0000-000000000000")
    mkdirSync(broken)
    writeFileSync(path.join(broken, "draft.json"), "broken")
    const linked = "11111111-1111-1111-1111-111111111111"
    symlinkSync(first.dir, path.join(base, linked))
    expect((await f.request(`/template-drafts/${linked}`)).status).toBe(400)
    const listed = await (await f.request("/template-drafts")).json() as { drafts: TemplateDraft[] }
    expect(listed.drafts.map(draft => draft.id)).toEqual([second.id, first.id])
    expect(JSON.stringify(listed)).not.toContain("paragraph_handle")
    expect((await f.request("/template-drafts/not-a-draft")).status).toBe(404)
  } finally { f.close() }
}, 60000)

test("unsupported revisions, bindings and embedded content leave no partial draft", async () => {
  const f = fixture()
  try {
    const source = await sourceDocument()
    const matter = f.register("matter-a", source)
    await capture(f)
    const base = path.join(f.workspace, ".template-drafts")
    const prior = readdirSync(base)
    const unsupported = [
      (parts: Record<string, Uint8Array>) => { parts["word/document.xml"] = strToU8(strFromU8(parts["word/document.xml"]!).replace("</w:p>", '<w:ins w:id="1"><w:r><w:t>Unresolved change</w:t></w:r></w:ins></w:p>')) },
      (parts: Record<string, Uint8Array>) => { parts["word/document.xml"] = strToU8(strFromU8(parts["word/document.xml"]!).replace("</w:p>", '<w:dataBinding w:storeItemID="synthetic"/></w:p>')) },
      (parts: Record<string, Uint8Array>) => { parts["word/embeddings/embedded.bin"] = new Uint8Array([1, 2, 3]) },
      (parts: Record<string, Uint8Array>) => { parts["word/document.xml"] = strToU8(strFromU8(parts["word/document.xml"]!).replace("</w:p>", '<w:fldSimple w:instr="MERGEFIELD SyntheticHiddenClient"><w:r><w:t>Generic visible text</w:t></w:r></w:fldSimple></w:p>')) },
    ]
    for (const mutate of unsupported) {
      const parts = unzipSync(source)
      mutate(parts)
      const invalid = zipSync(parts)
      writeFileSync(matter.file, invalid)
      const response = await f.request("/matters/matter-a/template-drafts", "POST", { name: "Selected.docx" })
      expect(response.status).toBe(422)
      expect((await response.json() as { error: string }).error).toMatch(/Word|tracked|embedded|bound/)
      expect(readdirSync(base)).toEqual(prior)
      expect(readFileSync(matter.file)).toEqual(Buffer.from(invalid))
    }
  } finally { f.close() }
}, 60000)

test("removes hidden client labels and building blocks while keeping safe fields and cross-references", async () => {
  const f = fixture()
  try {
    const parts = unzipSync(await sourceDocument())
    parts["word/glossary/document.xml"] = strToU8('<w:glossaryDocument xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docParts><w:docPart><w:docPartPr><w:name w:val="SyntheticHiddenClient"/></w:docPartPr></w:docPart></w:docParts></w:glossaryDocument>')
    const controls = '<w:sdt><w:sdtPr><w:alias w:val="SyntheticHiddenClient"/><w:tag w:val="SyntheticHiddenClient"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Generic controlled text</w:t></w:r></w:p></w:sdtContent></w:sdt>'
    const fields = '<w:p><w:bookmarkStart w:id="99" w:name="SyntheticHiddenClient"/><w:r><w:t>Reference target</w:t></w:r><w:bookmarkEnd w:id="99"/></w:p><w:p><w:fldSimple w:instr="PAGE"><w:r><w:t>1</w:t></w:r></w:fldSimple><w:fldSimple w:instr="REF SyntheticHiddenClient \\h"><w:r><w:t>Reference target</w:t></w:r></w:fldSimple></w:p>'
    parts["word/document.xml"] = strToU8(strFromU8(parts["word/document.xml"]!).replace("<w:sectPr>", `${controls}${fields}<w:sectPr>`))
    const original = zipSync(parts)
    const matter = f.register("matter-a", original)
    const draft = await capture(f)
    const clean = unzipSync(await candidate(f, draft))
    expect(clean["word/glossary/document.xml"]).toBeUndefined()
    const xml = strFromU8(clean["word/document.xml"]!)
    expect(xml).not.toContain("SyntheticHiddenClient")
    expect(xml).toContain('w:instr="PAGE"')
    expect(xml).toContain("REF TemplateBookmark1")
    expect(xml).toContain("Generic controlled text")
    expect((await f.request(`/template-drafts/${draft.id}/edits`, "POST", { revision: 0, name: "Clean fields.docx", description: "Reusable terms with page references", edits: [] })).status).toBe(200)
    expect((await f.request(`/template-drafts/${draft.id}/publish`, "POST", { revision: 1, name: "Clean fields.docx", description: "Reusable terms with page references" })).status).toBe(200)
    const saved = unzipSync(new Uint8Array(readFileSync(path.join(f.workspace, ".templates", "Clean fields.docx"))))
    expect(Object.keys(saved).some(name => name.startsWith("word/glossary/"))).toBe(false)
    expect(Object.entries(saved).filter(([name]) => name.endsWith(".xml") || name.endsWith(".rels")).some(([, bytes]) => strFromU8(bytes).includes("SyntheticHiddenClient"))).toBe(false)
    expect(readFileSync(matter.file)).toEqual(Buffer.from(original))
  } finally { f.close() }
}, 60000)

test("a placeholder edit removes shared text and image links while retaining image bytes and unrelated links", async () => {
  const f = fixture()
  try {
    const image = Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jJWoAAAAASUVORK5CYII=", "base64")
    const source = new Uint8Array(await Packer.toBuffer(new Document({ sections: [{ children: [
      new Paragraph({ children: [new ExternalHyperlink({ link: "https://example.test/contact?email=contact%40example.test", children: [new TextRun({ text: "contact@example.test", bold: true })] })] }),
      new Paragraph({ children: [new ImageRun({ type: "png", data: image, transformation: { width: 16, height: 16 } })] }),
      new Paragraph({ children: [new ExternalHyperlink({ link: "https://example.test/public", children: [new TextRun("Public notice")] })] }),
    ] }] })))
    const sourceParts = unzipSync(source)
    const document = new DOMParser().parseFromString(strFromU8(sourceParts["word/document.xml"]!), "application/xml")
    const word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    const drawing = "http://schemas.openxmlformats.org/drawingml/2006/main"
    const relationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    const textLinks = Array.from(document.getElementsByTagNameNS(word, "hyperlink"))
    const sensitiveId = textLinks[0]!.getAttributeNS(relationship, "id")!
    const publicId = textLinks[1]!.getAttributeNS(relationship, "id")!
    const imageProperties = document.getElementsByTagNameNS("http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing", "docPr")[0]!
    for (const kind of ["hlinkClick", "hlinkHover"]) {
      const link = document.createElementNS(drawing, `a:${kind}`)
      link.setAttributeNS(relationship, "r:id", sensitiveId)
      imageProperties.appendChild(link)
    }
    const pictureProperties = document.getElementsByTagNameNS("http://schemas.openxmlformats.org/drawingml/2006/picture", "cNvPr")[0]!
    const publicImageLink = document.createElementNS(drawing, "a:hlinkClick")
    publicImageLink.setAttributeNS(relationship, "r:id", publicId)
    pictureProperties.appendChild(publicImageLink)
    sourceParts["word/document.xml"] = strToU8(new XMLSerializer().serializeToString(document))
    const original = zipSync(sourceParts)
    const matter = f.register("matter-a", original)
    const draft = await capture(f)
    const revision = await (await f.request(`/template-drafts/${draft.id}/text`)).json() as DocxRevisionPacket
    const response = await f.request(`/template-drafts/${draft.id}/edits`, "POST", { revision: 0, edits: [{
      id: "email", paragraph_handle: revision.paragraphs[0].paragraph_handle, occurrence: 0,
      before_quote: "contact@example.test", replacement: "[insert contact email]",
    }] })
    expect(response.status).toBe(200)
    expect((await response.json() as TemplateDraft).notes.join(" ")).toContain("hyperlink target")
    const bytes = await candidate(f, draft)
    const parts = unzipSync(bytes)
    expect(strFromU8(parts["word/_rels/document.xml.rels"]!)).not.toContain("example.test/contact")
    expect(strFromU8(parts["word/_rels/document.xml.rels"]!)).toContain("example.test/public")
    expect(strFromU8(parts["word/document.xml"]!)).not.toContain("contact@example.test")
    const proposed = new DOMParser().parseFromString(strFromU8(parts["word/document.xml"]!), "application/xml")
    expect(Array.from(proposed.getElementsByTagName("*")).some(element => element.getAttributeNS(relationship, "id") === sensitiveId)).toBe(false)
    expect(proposed.getElementsByTagNameNS(drawing, "hlinkHover").length).toBe(0)
    expect(proposed.getElementsByTagNameNS(drawing, "hlinkClick").length).toBe(1)
    expect(proposed.getElementsByTagNameNS(drawing, "hlinkClick")[0]!.getAttributeNS(relationship, "id")).toBe(publicId)
    expect(proposed.getElementsByTagNameNS("http://schemas.openxmlformats.org/drawingml/2006/picture", "pic").length).toBe(1)
    for (const name of Object.keys(sourceParts).filter(name => name.startsWith("word/media/"))) expect(parts[name]).toEqual(sourceParts[name])
    expect((await createDocxRevision(bytes)).paragraphs[0].formatting).toMatchObject([{ text: "[insert contact email]", formatting: { bold: true } }])
    expect(readFileSync(matter.file)).toEqual(Buffer.from(original))
  } finally { f.close() }
}, 60000)

type NativeTestSession = { id: string; location: { directory: string }; parentID?: string }
function nativeConversations() {
  const sessions = new Map<string, NativeTestSession>()
  const deleted: string[] = []
  const calls: Array<{ path: string; method: string; directory: string | null }> = []
  const active = new Set<string>()
  const queued = new Set<string>()
  const waiting = new Set<string>()
  let failure: "list" | "state" | "delete" | undefined
  let owner: NativeTestSession | undefined
  let beforeList: (() => Promise<void>) | undefined
  const transport = async (input: RequestInfo | URL, init?: RequestInit) => {
    const url = new URL(input instanceof Request ? input.url : String(input))
    const method = init?.method ?? "GET"
    calls.push({ path: url.pathname, method, directory: new Headers(init?.headers).get("x-opencode-directory") })
    if (url.pathname === "/api/session") {
      if (beforeList) await beforeList()
      if (failure === "list") return new Response(null, { status: 500 })
      const all = [...sessions.values()].filter(session =>
        (!url.searchParams.has("directory") || session.location.directory === url.searchParams.get("directory")) &&
        (!url.searchParams.has("parentID") || session.parentID === url.searchParams.get("parentID")))
      const offset = Number(url.searchParams.get("cursor") ?? 0)
      // Deliberately paginate both directory-scoped and unfiltered child lists.
      return Response.json({ data: all.slice(offset, offset + 1), cursor: { next: offset + 1 < all.length ? String(offset + 1) : undefined } })
    }
    if (url.pathname === "/api/session/active") {
      if (failure === "state") return new Response(null, { status: 500 })
      return Response.json({ data: Object.fromEntries([...active].map(id => [id, { type: "running" }])) })
    }
    if (url.pathname === "/api/permission/request") return Response.json({ data: [...waiting].map(sessionID => ({ sessionID })) })
    const match = /^\/api\/session\/([^/]+)(\/inbox|\/message)?$/.exec(url.pathname)
    if (!match || !sessions.has(match[1])) return new Response(null, { status: 404 })
    const id = match[1]
    if (match[2] === "/inbox") return Response.json({ data: queued.has(id) ? [{ id: "msg_queued", sessionID: id }] : [] })
    if (match[2] === "/message") return Response.json({ data: [] })
    if (method === "GET") return Response.json({ data: owner ?? sessions.get(id) })
    if (method === "DELETE") {
      if (failure === "delete") return new Response(null, { status: 500 })
      deleted.push(id)
      const remove = (target: string) => {
        for (const child of [...sessions.values()].filter(session => session.parentID === target)) remove(child.id)
        sessions.delete(target)
      }
      remove(id)
      return new Response(null, { status: 204 })
    }
    return new Response(null, { status: 405 })
  }
  return {
    sessions, deleted, calls, active, queued, waiting,
    options: { engineURL: "http://native.test", fetch: transport as typeof fetch },
    fail: (value?: typeof failure) => { failure = value },
    overrideOwner: (value?: NativeTestSession) => { owner = value },
    pauseList: (run?: () => Promise<void>) => { beforeList = run },
    add: (id: string, directory: string, parentID?: string) => {
      const session = { id, location: { directory }, ...(parentID ? { parentID } : {}) }
      sessions.set(id, session)
      return session
    },
  }
}

test("deletes only the chosen template draft and its native conversation tree, preserving source and saved library", async () => {
  const engine = nativeConversations()
  const f = fixture(engine.options)
  try {
    const original = await sourceDocument()
    const matter = f.register("matter-a", original)
    const draft = await capture(f)
    const other = await capture(f)
    await f.request(`/template-drafts/${draft.id}/edits`, "POST", { revision: 0, edits: [], name: "Saved.docx", description: "Reusable terms" })
    expect((await f.request(`/template-drafts/${draft.id}/publish`, "POST", { revision: 1, name: "Saved.docx", description: "Reusable terms" })).status).toBe(200)
    const saved = readFileSync(path.join(f.workspace, ".templates", "Saved.docx"))
    const manifest = readFileSync(path.join(f.workspace, ".templates", "manifest.json"))
    const otherMetadata = readFileSync(path.join(other.dir, "draft.json"))
    engine.add("ses_one", draft.dir)
    engine.add("ses_child", draft.dir, "ses_one")
    engine.add("ses_child_two", draft.dir, "ses_one")
    engine.add("ses_two", draft.dir)
    engine.add("ses_matter", matter.dir)
    engine.add("ses_other", other.dir)
    symlinkSync(matter.dir, path.join(draft.dir, "linked-source"), "dir")
    const response = await f.request(`/template-drafts/${draft.id}`, "DELETE")
    expect(response.status).toBe(200)
    expect(await response.json()).toEqual({ ok: true })
    expect(existsSync(draft.dir)).toBe(false)
    expect((await f.request(`/template-drafts/${draft.id}`)).status).toBe(404)
    expect((await (await f.request("/template-drafts")).json()).drafts.map((row: TemplateDraft) => row.id)).toEqual([other.id])
    expect(engine.deleted).toEqual(["ses_one", "ses_two"])
    expect([...engine.sessions.keys()]).toEqual(["ses_matter", "ses_other"])
    expect(engine.calls.every(call => call.directory === draft.dir)).toBe(true)
    expect(readFileSync(matter.file)).toEqual(Buffer.from(original))
    expect(readFileSync(path.join(other.dir, "draft.json"))).toEqual(otherMetadata)
    expect(readFileSync(path.join(f.workspace, ".templates", "Saved.docx"))).toEqual(saved)
    expect(readFileSync(path.join(f.workspace, ".templates", "manifest.json"))).toEqual(manifest)
  } finally { f.close() }
})

test("running, queued or permission-waiting native work refuses deletion; engine failures preserve the draft for retry", async () => {
  const engine = nativeConversations()
  const f = fixture(engine.options)
  try {
    const original = await sourceDocument()
    const matter = f.register("matter-a", original)
    const draft = await capture(f)
    engine.add("ses_owner", draft.dir)
    engine.add("ses_child", draft.dir, "ses_owner")
    const metadata = readFileSync(path.join(draft.dir, "draft.json"))
    const proposal = await candidate(f, draft)
    for (const pending of [engine.active, engine.queued, engine.waiting]) {
      pending.add("ses_child")
      const response = await f.request(`/template-drafts/${draft.id}`, "DELETE")
      expect(response.status).toBe(409)
      expect((await response.json()).error).toContain("Stop this template conversation")
      pending.clear()
    }
    for (const failure of ["list", "state", "delete"] as const) {
      engine.fail(failure)
      const response = await f.request(`/template-drafts/${draft.id}`, "DELETE")
      expect(response.status).toBe(503)
      expect((await response.json()).error).toContain("draft has been kept")
      expect(readFileSync(path.join(draft.dir, "draft.json"))).toEqual(metadata)
      expect(await candidate(f, draft)).toEqual(proposal)
      expect(readFileSync(matter.file)).toEqual(Buffer.from(original))
      expect([...engine.sessions.keys()]).toEqual(["ses_owner", "ses_child"])
      expect(engine.deleted).toEqual([])
    }
    engine.fail()
    expect((await f.request(`/template-drafts/${draft.id}`, "DELETE")).status).toBe(200)
  } finally { f.close() }
})

test("draft deletion rejects path aliases, mismatched session identities and descendants moved outside the draft", async () => {
  const engine = nativeConversations()
  const f = fixture(engine.options)
  try {
    f.register("matter-a", await sourceDocument())
    const draft = await capture(f)
    const other = await capture(f)
    const alias = "11111111-1111-4111-8111-111111111111"
    symlinkSync(draft.dir, path.join(f.workspace, ".template-drafts", alias), "dir")
    expect((await f.request(`/template-drafts/${alias}`, "DELETE")).status).toBe(400)
    expect((await f.request("/template-drafts/not-a-draft", "DELETE")).status).toBe(404)
    expect(engine.calls).toEqual([])
    engine.add("ses_owner", draft.dir)
    engine.overrideOwner({ id: "ses_wrong", location: { directory: draft.dir } })
    expect((await f.request(`/template-drafts/${draft.id}`, "DELETE")).status).toBe(409)
    engine.overrideOwner({ id: "ses_owner", location: { directory: `${draft.dir}/../${draft.id}` } })
    expect((await f.request(`/template-drafts/${draft.id}`, "DELETE")).status).toBe(503)
    engine.overrideOwner()
    engine.add("ses_moved", other.dir, "ses_owner")
    expect((await f.request(`/template-drafts/${draft.id}`, "DELETE")).status).toBe(409)
    expect(engine.deleted).toEqual([])
    expect(existsSync(draft.dir)).toBe(true)
    expect(existsSync(other.dir)).toBe(true)
    expect([...engine.sessions.keys()]).toEqual(["ses_owner", "ses_moved"])
  } finally { f.close() }
})

test("deletion serializes with edits and publication without recreating a deleted draft", async () => {
  const engine = nativeConversations()
  const f = fixture(engine.options)
  let release!: () => void
  try {
    f.register("matter-a", await sourceDocument())
    const draft = await capture(f)
    await f.request(`/template-drafts/${draft.id}/edits`, "POST", { revision: 0, edits: [], name: "Waiting.docx", description: "Reusable terms" })
    let entered!: () => void
    const started = new Promise<void>(resolve => { entered = resolve })
    const gate = new Promise<void>(resolve => { release = resolve })
    engine.pauseList(async () => { entered(); await gate })
    const deletion = f.request(`/template-drafts/${draft.id}`, "DELETE")
    await started
    const edits = f.request(`/template-drafts/${draft.id}/edits`, "POST", { revision: 1, edits: [], name: "Later.docx" })
    const publish = f.request(`/template-drafts/${draft.id}/publish`, "POST", { revision: 1, name: "Waiting.docx", description: "Reusable terms" })
    release()
    expect((await deletion).status).toBe(200)
    expect((await edits).status).toBe(404)
    expect((await publish).status).toBe(404)
    expect(existsSync(draft.dir)).toBe(false)
    expect(existsSync(path.join(f.workspace, ".templates", "Waiting.docx"))).toBe(false)
  } finally { release?.(); f.close() }
})
