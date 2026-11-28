import { expect, test } from "bun:test"
import { DerivedChronologyError, DocumentExistsError, DuplicateUploadError, MATTER_DOCUMENT_EXTENSIONS, matterDocumentKind, type AttachmentStatus, type Document, type EmailAttachment, type IngestResult } from "../src/api/ingest/documents"
import { attachmentsSummary, documentTree, failedUpload, newUpload, sameNamedSource, uploadedToast } from "../src/components/DocumentUpload"

const file = new File(["theirs"], "Agreement.docx")

test("failedUpload classifies chronology, exists, duplicate, deliverable collision, and generic errors", () => {
  expect(failedUpload(new DerivedChronologyError("My corrections.docx", "Chronology.docx", "This is a chronology export."), file))
    .toMatchObject({ status: "chronology", chronology: { name: "My corrections.docx", originalName: "Chronology.docx" } })

  expect(failedUpload(new DocumentExistsError("Agreement.docx", "source", "Agreement.docx is already in this matter."), file))
    .toEqual({ status: "exists", existing: { name: "Agreement.docx", role: "source" } })

  expect(failedUpload(new DocumentExistsError("Agreement.docx", "deliverable", "Cannot replace generated deliverable Agreement.docx with a source document."), file))
    .toEqual({ status: "exists", existing: { name: "Agreement.docx", role: "deliverable" } })

  const older = failedUpload(new DuplicateUploadError("Agreement.docx", { id: "v-2", number: 2, current: false }, "server words"), file)
  expect(older).toMatchObject({ status: "duplicate", match: { name: "Agreement.docx", version: { id: "v-2", number: 2 } } })
  expect(older.retryable).toBeUndefined()

  const current = failedUpload(new DuplicateUploadError("Agreement.docx", { id: "v-4", number: 4, current: true }, ""), file)
  expect(current.status).toBe("duplicate")
  expect(current.match).toBeUndefined()

  const renamed = new File(["identical"], "Renamed.pdf")
  expect(failedUpload(new DuplicateUploadError("Original.pdf", { id: "v1", number: 1, current: true }, "", true), renamed))
    .toMatchObject({ status: "duplicate", canKeepBoth: true })

  expect(failedUpload(new Error("The embedding model could not be loaded."), file))
    .toEqual({ status: "failed", error: "The embedding model could not be loaded.", retryable: true })
})

test("a same-name source is found in the matter's own list before the file is uploaded, whatever its case", () => {
  const documents: Document[] = [
    { id: 1, name: "Agreement.docx", doc_path: "/m/Agreement.docx", created_at: 1, role: "source" },
    { id: 2, name: "Draft.docx", doc_path: "/m/Draft.docx", created_at: 2, role: "deliverable" },
  ]
  expect(sameNamedSource(documents, new File(["x"], "agreement.DOCX"))).toMatchObject({ name: "Agreement.docx", role: "source" })
  expect(sameNamedSource(documents, new File(["x"], "Draft.docx"))).toMatchObject({ name: "Draft.docx", role: "deliverable" })
  expect(sameNamedSource(documents, new File(["x"], "New terms.docx"))).toBeUndefined()
})

test("every format ingest reads can be added; any other file is refused before upload", () => {
  expect([...MATTER_DOCUMENT_EXTENSIONS]).toEqual([".docx", ".pdf", ".eml", ".msg", ".xlsx", ".zip", ".md", ".markdown"])
  for (const [name, kind] of [
    ["Agreement.docx", "docx"], ["Scan.PDF", "pdf"], ["Client instructions.eml", "eml"],
    ["Forwarded thread.MSG", "msg"], ["Cap table.XLSX", "xlsx"], ["Artwork files.ZIP", "zip"],
    ["Chronology.md", "md"], ["Notes.MARKDOWN", "markdown"],
  ] as const) {
    expect(matterDocumentKind(name)).toBe(kind)
    expect(newUpload(1, new File(["synthetic"], name))).toMatchObject({ key: 1, status: "queued" })
  }
  for (const [name, kind] of [[".hidden.eml", "eml"], ["..pdf", "pdf"]] as const) expect(matterDocumentKind(name)).toBe(kind)
  for (const name of ["Notes.mbox", "Old.doc", "Table.csv", "eml", "mail.eml.txt", "Agreement", ".pdf", ".DOCX", " .xlsx"]) {
    expect(matterDocumentKind(name)).toBeUndefined()
    const upload = newUpload(2, new File(["synthetic"], name))
    expect(upload.status).toBe("failed")
    expect(upload.retryable).toBeUndefined()
  }
})

test("uploadedToast and attachmentsSummary classify warnings, tracked-change notices, unreadable files, and attachment outcomes", () => {
  const base: IngestResult = { name: "Scan.pdf", docPath: "/m/Scan.pdf", sections: 0, chunks: 0, injection: null }
  expect(uploadedToast({ ...base, sections: 2, quality: { status: "ok" } }).kind).toBe("success")

  const noticed = uploadedToast({
    ...base, name: "Agreement.docx", sections: 4, chunks: 4, version: { id: "v-3", number: 3, created: true },
    notice: "Saved as version 3, with 2 tracked changes.",
  })
  expect(noticed.kind).toBe("warning")
  expect(noticed.message).toContain("Saved as version 3")

  const unreadable = uploadedToast({ ...base, quality: { status: "no_readable_text", signals: [{ code: "empty", detail: "No text could be extracted" }] } })
  expect(unreadable.kind).toBe("warning")
  expect(unreadable.message).toContain("not indexed")

  const signals = [{ code: "little_text_per_page", detail: "40 characters over 2 pages" }]
  const injection = { findings: [{ rule: "instruction-override", detail: "override attempt" }], flaggedChunks: 1 }
  const combined = uploadedToast({ ...base, sections: 1, injection, quality: { status: "degraded", signals } })
  expect(combined.kind).toBe("warning")
  expect(combined.message).toContain("override attempt")
  expect(combined.message).toContain("40 characters over 2 pages")

  const unreadableInjection = uploadedToast({ ...base, injection, quality: { status: "no_readable_text", signals: [{ code: "empty", detail: "No text could be extracted" }] } })
  expect(unreadableInjection.kind).toBe("warning")
  expect(unreadableInjection.message).toContain("not indexed")
  expect(unreadableInjection.message).toContain("override attempt")

  expect(attachmentsSummary({})).toBe("")
  const mixed: EmailAttachment[] = [
    { name: "Schedule.docx", type: "application/pdf", size: 1, status: "indexed", document: "Schedule.docx" },
    { name: "Draft.docx", type: "application/pdf", size: 1, status: "same", document: "Draft.docx" },
    { name: "Scan.pdf", type: "application/pdf", size: 1, status: "no_readable_text", document: "Scan.pdf" },
    { name: "bundle.zip", type: "application/zip", size: 1, status: "unsupported", note: "linked" },
    { name: "Locked.docx", type: "application/pdf", size: 1, status: "encrypted" },
    { name: "", type: "application/pdf", size: 1, status: "failed" },
  ]
  const summary = attachmentsSummary({ attachments: mixed })
  expect(summary).toContain("1 attachment added")
  expect(summary).toContain("Scan.pdf")
  expect(summary).toContain("bundle.zip")
  expect(summary).toContain("Locked.docx")
  expect(attachmentsSummary({ attachments: [{ name: "Later.bin", type: "", size: 1, status: "from_a_newer_ingest" as unknown as "failed" }] })).toContain("Later.bin")
})

test("the documents list nests each extracted document once under its parent, at any depth, without losing orphans or cycles", () => {
  const doc = (id: number, name: string, extra: Partial<Document> = {}): Document => ({ id, name, doc_path: `/m/${name}`, created_at: id, role: "source", ...extra })
  const entry = (name: string, status: AttachmentStatus, document?: string) => ({ name, type: "", size: 1, status, ...(document ? { document } : {}) })
  const of = (name: string, removed = false) => ({ attachment_of: { name, removed } })
  const tree = documentTree([
    doc(1, "Thread.eml", { attachments: [entry("Bundle.zip", "indexed", "Bundle.zip"), entry("Logo.exe", "unsupported"), entry("Shared.pdf", "same", "Shared.pdf"), entry("Old.docx", "removed", "Old.docx")] }),
    doc(2, "Bundle.zip", of("Thread.eml")),
    doc(3, "Inner.docx", of("Bundle.zip")),
    doc(4, "Shared.pdf"),
    doc(5, "Orphan.docx", of("Gone.eml", true)),
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
