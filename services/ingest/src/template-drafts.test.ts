import { expect, test } from "bun:test"
import { mkdtempSync, mkdirSync, readFileSync, readdirSync, rmSync, symlinkSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { Document, ExternalHyperlink, Footer, Header, HeadingLevel, ImageRun, Packer, Paragraph, Table, TableCell, TableRow, TextRun } from "docx"
import { createDocxRevision, type DocxParagraphEdit, type DocxRevisionPacket } from "../../../dochaus/lib/docx-amendments"
import { openDb, upsertDeliverable } from "./db"
import { templateDraftRoutes, type TemplateDraft } from "./template-drafts"

function fixture() {
  const temporary = mkdtempSync(path.join(tmpdir(), "dochaus-template-drafts-"))
  const workspace = path.join(temporary, "workspace")
  mkdirSync(workspace)
  const app = templateDraftRoutes(id => path.join(workspace, id), () => workspace)
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
