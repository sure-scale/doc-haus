import { expect, test } from "bun:test"
import { createElement } from "react"
import { renderToStaticMarkup } from "react-dom/server"
import { DocumentExistsError, MATTER_DOCUMENT_EXTENSIONS, matterDocumentKind, type Document } from "../src/api/ingest/documents"
import DocumentUpload, { failedUpload, newUpload, UploadRow } from "../src/components/DocumentUpload"

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
  expect([...MATTER_DOCUMENT_EXTENSIONS]).toEqual([".docx", ".pdf", ".eml", ".xlsx"])
  for (const [name, kind] of [["Agreement.docx", "docx"], ["Scan.PDF", "pdf"], ["Client instructions.eml", "eml"], ["Cap table.XLSX", "xlsx"]] as const) {
    expect(matterDocumentKind(name)).toBe(kind)
    expect(newUpload(1, new File(["synthetic"], name))).toMatchObject({ key: 1, status: "queued" })
  }
  for (const name of ["Notes.msg", "Old.doc", "Table.csv", "eml", "mail.eml.txt", "Agreement"]) {
    expect(matterDocumentKind(name)).toBeUndefined()
    expect(newUpload(2, new File(["synthetic"], name))).toMatchObject({
      status: "failed",
      error: "Only Word (.docx), PDF, email (.eml) and Excel (.xlsx) files can be added.",
    })
    expect(newUpload(2, new File(["synthetic"], name)).retryable).toBeUndefined()
  }
})

test("the picker offers every format and the list marks emails and spreadsheets with their own glyph", () => {
  const documents: Document[] = ["Agreement.docx", "Scan.pdf", "Client instructions.eml", "Cap table.xlsx"].map((name, index) => ({
    id: index + 1, name, doc_path: `/matters/m/${name}`, created_at: index + 1, role: "source",
  }))
  const html = renderToStaticMarkup(createElement(DocumentUpload, { matterId: "m", documents, onUploaded: () => {}, onView: () => {} }))
  expect(html).toContain('accept=".docx,.pdf,.eml,.xlsx"')
  expect(html).toContain("Drop Word, PDF, email or Excel files here, or click to browse.")
  const glyphs = [...html.matchAll(/<span class="doc-type is-(\w+)" title="([^"]+)"[^>]*>([\s\S]*?)<\/span>/g)].map(([, kind, title, svg]) => ({ kind, title, svg }))
  expect(glyphs.map(({ kind, title }) => [kind, title])).toEqual([
    ["xlsx", "Excel spreadsheet"], ["eml", "Email"], ["pdf", "PDF document"], ["docx", "Word document"],
  ])
  // Email and spreadsheet glyphs differ from the document glyph Word and PDF share.
  expect(new Set(glyphs.map(({ svg }) => svg)).size).toBe(3)
})
