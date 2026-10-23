import { expect, test } from "bun:test"
import { createElement } from "react"
import { renderToStaticMarkup } from "react-dom/server"
import { DocumentExistsError, MATTER_DOCUMENT_EXTENSIONS, matterDocumentKind, type AttachmentStatus, type Document, type IngestResult } from "../src/api/ingest/documents"
import DocumentUpload, { attachmentsSummary, documentTree, failedUpload, newUpload, uploadedToast, UploadRow } from "../src/components/DocumentUpload"

const file = new File(["theirs"], "Agreement.docx")
const row = (upload: Parameters<typeof UploadRow>[0]["upload"]) => {
  const html = renderToStaticMarkup(createElement(UploadRow, { upload, onRetry: () => {}, onChoose: () => {}, onDismiss: () => {} }))
  return {
    status: html.match(/<span class="upload-status">([^<]*)<\/span>/)?.[1]?.trim(),
    buttons: [...html.matchAll(/<button[^>]*>([\s\S]*?)<\/button>/g)].map(([, label]) => label!.trim()),
    html,
  }
}

test("a same-name upload asks the lawyer to keep both, replace the existing file or cancel, safest first", () => {
  const exists = failedUpload(new DocumentExistsError("Agreement.docx", "source", "Agreement.docx is already in this matter."), file)
  expect(exists).toEqual({ status: "exists", existing: { name: "Agreement.docx", role: "source" } })
  const rendered = row({ key: 1, file, status: "queued", ...exists })
  expect(rendered).toMatchObject({
    status: "Already in this matter. Keep both or replace it?",
    buttons: ["Keep both", "Replace", "Cancel"],
  })
  // Replace overwrites, so it is never the emphasised choice.
  expect(rendered.html).toContain('class="icon-btn upload-keep-both"')
  expect(rendered.html).not.toContain("danger")
})

test("a returned draft named like a generated document can only be kept beside it", () => {
  const exists = failedUpload(
    new DocumentExistsError("Agreement.docx", "deliverable", "Cannot replace generated deliverable Agreement.docx with a source document."),
    file,
  )
  expect(exists).toEqual({ status: "exists", existing: { name: "Agreement.docx", role: "deliverable" } })
  expect(row({ key: 3, file, status: "queued", ...exists })).toMatchObject({
    status: "A document created by Doc.Haus has this name. Keep both?",
    buttons: ["Keep both", "Cancel"],
  })
})

test("any other failed upload stays retryable and offers no replace choice", () => {
  const failed = failedUpload(new Error("The embedding model could not be loaded."), file)
  expect(failed).toEqual({ status: "failed", error: "The embedding model could not be loaded.", retryable: true })
  expect(row({ key: 2, file, status: "queued", ...failed })).toMatchObject({
    status: "The embedding model could not be loaded.",
    buttons: ["Retry", "Remove"],
  })
})

test("every format ingest reads can be added; any other file is refused before upload", () => {
  // Mirrors MATTER_DOCUMENT_EXTENSIONS in services/ingest/src/ingest.ts.
  expect([...MATTER_DOCUMENT_EXTENSIONS]).toEqual([".docx", ".pdf", ".eml", ".xlsx", ".zip"])
  for (const [name, kind] of [["Agreement.docx", "docx"], ["Scan.PDF", "pdf"], ["Client instructions.eml", "eml"], ["Cap table.XLSX", "xlsx"], ["Artwork files.ZIP", "zip"]] as const) {
    expect(matterDocumentKind(name)).toBe(kind)
    expect(newUpload(1, new File(["synthetic"], name))).toMatchObject({ key: 1, status: "queued" })
  }
  // Like path.extname, which ingest checks: a name whose only dot leads it has no extension.
  for (const [name, kind] of [[".hidden.eml", "eml"], ["..pdf", "pdf"]] as const) expect(matterDocumentKind(name)).toBe(kind)
  for (const name of ["Notes.msg", "Old.doc", "Table.csv", "eml", "mail.eml.txt", "Agreement", ".pdf", ".DOCX", " .xlsx"]) {
    expect(matterDocumentKind(name)).toBeUndefined()
    expect(newUpload(2, new File(["synthetic"], name))).toMatchObject({
      status: "failed",
      error: "Only Word (.docx), PDF, email (.eml), Excel (.xlsx) and zip (.zip) files can be added.",
    })
    expect(newUpload(2, new File(["synthetic"], name)).retryable).toBeUndefined()
  }
})

test("the picker offers every format and the list marks emails and spreadsheets with their own glyph", () => {
  const documents: Document[] = ["Agreement.docx", "Scan.pdf", "Client instructions.eml", "Cap table.xlsx"].map((name, index) => ({
    id: index + 1, name, doc_path: `/matters/m/${name}`, created_at: index + 1, role: "source",
  }))
  const html = renderToStaticMarkup(createElement(DocumentUpload, { matterId: "m", documents, onUploaded: () => {}, onView: () => {} }))
  expect(html).toContain('accept=".docx,.pdf,.eml,.xlsx,.zip"')
  expect(html).toContain("Drop Word, PDF, email, Excel or zip files here, or click to browse.")
  const glyphs = [...html.matchAll(/<span class="doc-type is-(\w+)" title="([^"]+)"[^>]*>([\s\S]*?)<\/span>/g)].map(([, kind, title, svg]) => ({ kind, title, svg }))
  expect(glyphs.map(({ kind, title }) => [kind, title])).toEqual([
    ["xlsx", "Excel spreadsheet"], ["eml", "Email"], ["pdf", "PDF document"], ["docx", "Word document"],
  ])
  // Email and spreadsheet glyphs differ from the document glyph Word and PDF share.
  expect(new Set(glyphs.map(({ svg }) => svg)).size).toBe(3)
})

test("an uploaded email's ready message says which attachments were added and which were not", () => {
  expect(attachmentsSummary({})).toBe("")
  expect(attachmentsSummary({ attachments: [
    { name: "Schedule.docx", type: "application/pdf", size: 1, status: "indexed", document: "Schedule.docx" },
    { name: "Draft.docx", type: "application/pdf", size: 1, status: "same", document: "Draft.docx" },
    { name: "bundle.zip", type: "application/zip", size: 1, status: "unsupported" },
    { name: "", type: "application/pdf", size: 1, status: "failed" },
  ] })).toBe(" 1 attachment added. 1 already in the matter. Not added: bundle.zip (not supported), Unnamed attachment (could not be read).")
})

test("an uploaded zip's ready message counts files the same way", () => {
  const result: IngestResult = {
    name: "Artwork files.zip", docPath: "/m/Artwork files.zip", sections: 1, chunks: 1, injection: null, quality: { status: "ok" },
    attachments: [
      { name: "Artwork/Brief.docx", type: "application/pdf", size: 1, status: "indexed", document: "Brief.docx" },
      { name: "Artwork/Proof.pdf", type: "application/pdf", size: 1, status: "indexed", document: "Proof.pdf" },
      { name: "Artwork/logo.png", type: "application/octet-stream", size: 1, status: "unsupported" },
      { name: "Locked.docx", type: "application/pdf", size: 1, status: "encrypted" },
    ],
  }
  expect(uploadedToast(result)).toEqual({
    kind: "success",
    message: "Artwork files.zip is ready. 2 files added. Not added: Artwork/logo.png (not supported), Locked.docx (password-protected).",
  })
})

// The extraction quality gate (#239): a file with no readable text is added but
// never reported as ready; a degraded one is ready with a warning.
test("an upload with no readable text is never reported ready, and a degraded one warns", () => {
  const base: IngestResult = { name: "Scan.pdf", docPath: "/m/Scan.pdf", sections: 0, chunks: 0, injection: null }
  const unreadable = uploadedToast({ ...base, quality: { status: "no_readable_text", signals: [{ code: "empty", detail: "No text could be extracted" }] } })
  expect(unreadable).toEqual({
    kind: "warning",
    message: "Scan.pdf was added but not indexed: no readable text (no text could be extracted). The assistant cannot search or cite it.",
  })
  expect(row({ key: 4, file, status: "not_indexed" }).status).toBe("Added, not indexed: no readable text")

  const signals = [{ code: "little_text_per_page", detail: "40 characters over 2 pages; pages may be images whose text was not read" }]
  expect(uploadedToast({ ...base, sections: 1, chunks: 1, quality: { status: "degraded", signals } })).toEqual({
    kind: "warning",
    message: "Scan.pdf is ready, with a warning — its text may be incomplete or garbled (40 characters over 2 pages; pages may be images whose text was not read).",
  })
  const injection = { findings: [{ rule: "instruction-override", detail: "text that tries to override the assistant's instructions" }], flaggedChunks: 1 }
  expect(uploadedToast({ ...base, sections: 1, injection, quality: { status: "degraded", signals } }).message).toBe(
    "Scan.pdf is ready, with warnings — possible prompt injection: text that tries to override the assistant's instructions. The assistant will treat this content as untrusted. Its text may be incomplete or garbled (40 characters over 2 pages; pages may be images whose text was not read).",
  )
  // An unreadable file that also carries an injection finding shows both.
  expect(uploadedToast({ ...base, injection, quality: { status: "no_readable_text", signals: [{ code: "empty", detail: "No text could be extracted" }] } })).toEqual({
    kind: "warning",
    message:
      "Scan.pdf was added but not indexed: no readable text (no text could be extracted). The assistant cannot search or cite it. It also has a warning — possible prompt injection: text that tries to override the assistant's instructions. The assistant will treat this content as untrusted.",
  })
  // A quality result never reads as an injection warning, and ok adds nothing.
  expect(uploadedToast({ ...base, sections: 2, quality: { status: "ok" } })).toEqual({ kind: "success", message: "Scan.pdf is ready (2 sections)." })
})

test("the list says which sources have no readable text or degraded text, and nothing for the rest", () => {
  const documents: Document[] = [
    { id: 1, name: "Blank.docx", doc_path: "/m/Blank.docx", created_at: 4, role: "source", quality: { status: "no_readable_text", signals: [{ code: "empty", detail: "No text could be extracted" }] } },
    { id: 2, name: "Garbled.docx", doc_path: "/m/Garbled.docx", created_at: 3, role: "source", quality: { status: "degraded", signals: [{ code: "replacement_characters", detail: "40 of 900 characters are replacement or control characters" }] } },
    { id: 3, name: "Agreement.docx", doc_path: "/m/Agreement.docx", created_at: 2, role: "source", quality: { status: "ok" } },
    // Ingested before the gate: no result, nothing shown.
    { id: 4, name: "Older.docx", doc_path: "/m/Older.docx", created_at: 1, role: "source" },
  ]
  const html = renderToStaticMarkup(createElement(DocumentUpload, { matterId: "m", documents, onUploaded: () => {}, onView: () => {} }))
  const notes = [...html.matchAll(/<span class="document-quality is-(\w+)" title="([^"]*)">([^<]*)<\/span>/g)].map(([, status, title, label]) => ({ status, title, label }))
  expect(notes).toEqual([
    { status: "no_readable_text", title: "No text could be extracted", label: "Not indexed: no readable text" },
    { status: "degraded", title: "40 of 900 characters are replacement or control characters", label: "Text may be incomplete or garbled" },
  ])
})

test("an email lists an attachment with no readable text, and one whose text is degraded", () => {
  const attachments = [
    { name: "Scan.pdf", type: "application/pdf", size: 1, status: "no_readable_text" as const, document: "Scan.pdf" },
    { name: "Notes.docx", type: "application/pdf", size: 1, status: "indexed" as const, document: "Notes.docx", degraded: true as const },
  ]
  expect(attachmentsSummary({ attachments })).toBe(" 1 attachment added. Added but not indexed, no readable text: Scan.pdf.")
  // A status from a newer ingest still reads as not added.
  expect(attachmentsSummary({ attachments: [{ name: "Later.bin", type: "", size: 1, status: "from_a_newer_ingest" as unknown as "failed" }] }))
    .toBe(" Not added: Later.bin (not added).")
})

// Each extracted document is listed once, under the document it came from, by
// the parent link whatever the parent's type; attachments that did not become
// one of the parent's documents stay in its list.
test("the documents list nests each extracted document once under its parent, at any depth", () => {
  const doc = (id: number, name: string, extra: Partial<Document> = {}): Document => ({ id, name, doc_path: `/m/${name}`, created_at: id, role: "source", ...extra })
  const entry = (name: string, status: AttachmentStatus, document?: string) =>
    ({ name, type: "", size: 1, status, ...(document ? { document } : {}) })
  const of = (name: string, removed = false) => ({ attachment_of: { name, removed } })
  const tree = documentTree([
    doc(1, "Thread.eml", { attachments: [entry("Bundle.zip", "indexed", "Bundle.zip"), entry("Logo.exe", "unsupported"), entry("Shared.pdf", "same", "Shared.pdf"), entry("Old.docx", "removed", "Old.docx")] }),
    doc(2, "Bundle.zip", of("Thread.eml")),
    // No attachment record on the archive: its files nest by the link alone.
    doc(3, "Inner.docx", of("Bundle.zip")),
    doc(4, "Shared.pdf"),
    doc(5, "Orphan.docx", of("Gone.eml", true)),
    // A link loop is still listed rather than lost.
    doc(6, "A.eml", of("B.eml")),
    doc(7, "B.eml", of("A.eml")),
  ])
  const shape = (node: (typeof tree)[number] | { attachment: { name: string; status: string } }): unknown =>
    "document" in node ? [node.document.name, ...node.children.map(shape)] : `${node.attachment.name}:${node.attachment.status}`
  expect(tree.map(shape)).toEqual([
    ["Thread.eml", ["Bundle.zip", ["Inner.docx"]], "Logo.exe:unsupported", "Shared.pdf:same", "Old.docx:removed"],
    ["Shared.pdf"],
    ["Orphan.docx"],
    ["A.eml", ["B.eml"]],
  ])
})
