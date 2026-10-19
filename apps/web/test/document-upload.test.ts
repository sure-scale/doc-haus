import { expect, test } from "bun:test"
import { createElement } from "react"
import { renderToStaticMarkup } from "react-dom/server"
import { SourceExistsError } from "../src/api/ingest/documents"
import { failedUpload, UploadRow } from "../src/components/DocumentUpload"

const file = new File(["theirs"], "Agreement.docx")
const row = (upload: Parameters<typeof UploadRow>[0]["upload"]) => {
  const html = renderToStaticMarkup(createElement(UploadRow, { upload, onRetry: () => {}, onChoose: () => {}, onDismiss: () => {} }))
  return {
    status: html.match(/<span class="upload-status">([^<]*)<\/span>/)?.[1],
    buttons: [...html.matchAll(/<button[^>]*>([\s\S]*?)<\/button>/g)].map(([, label]) => label!.trim()),
  }
}

test("a same-name upload asks the lawyer to replace the existing file, keep both or cancel", () => {
  const exists = failedUpload(new SourceExistsError("Agreement.docx", "Agreement.docx is already in this matter."), file)
  expect(exists).toEqual({ status: "exists", existing: "Agreement.docx" })
  expect(row({ key: 1, file, status: "queued", ...exists })).toEqual({
    status: "Already in this matter. Replace it or keep both?",
    buttons: ["Replace", "Keep both", "Cancel"],
  })
})

test("any other failed upload stays retryable and offers no replace choice", () => {
  const failed = failedUpload(new Error("The embedding model could not be loaded."), file)
  expect(failed).toEqual({ status: "failed", error: "The embedding model could not be loaded.", retryable: true })
  expect(row({ key: 2, file, status: "queued", ...failed })).toEqual({
    status: "The embedding model could not be loaded.",
    buttons: ["Retry", "Remove"],
  })
})
