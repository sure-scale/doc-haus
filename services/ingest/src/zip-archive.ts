import { parseZip, ZipFormatError } from "../../../dochaus/lib/zip"
import { ingestContainerDocument, type ContainerOptions, type ContainerPart } from "./email-attachments"

// Zip files. A zip uploaded to a matter, or found as an email's attachment (or
// inside another zip), is a container like an email (email-attachments.ts): it
// is kept as a matter document whose text lists the files inside
// (dochaus/lib/zip.ts), and each file a matter document can take is ingested
// after it as a document that names the zip, with the same statuses, safe
// names, de-duplication, depth limit and removal as an email's attachments.
// Other files are listed as not supported; a password-protected file, or one
// whose path is absolute or climbs out of the archive, is listed and not added.
// Folders and macOS metadata are left out. Nothing is unpacked to disk except
// each added document, through the usual ingest into the matter.

// What one upload may unpack from zip files, counted over every zip in it
// (nested zips and zips attached to emails included), so nesting cannot
// multiply it. The byte limit is twice the web server's default request size
// (128 MB), room for an upload near that size that compresses well; the file
// limit is far above any real bundle of matter documents.
export const MAX_ZIP_UNCOMPRESSED_BYTES = 256 * 1024 * 1024
export const MAX_ZIP_FILES = 1000

export type ArchiveBudget = { bytes: number; files: number }

export function archiveBudget(): ArchiveBudget {
  return { bytes: MAX_ZIP_UNCOMPRESSED_BYTES, files: MAX_ZIP_FILES }
}

// A zip that would take the upload past the limits. Nothing of it is added:
// uploaded directly, the upload is refused with this message; inside another
// document, it is listed as too large.
export class ZipTooLargeError extends Error {
  constructor(name: string, reason: "bytes" | "files") {
    super(
      reason === "bytes"
        ? `${name} was not added: its files would unpack to more than ${MAX_ZIP_UNCOMPRESSED_BYTES / (1024 * 1024)} MB. Split it into smaller zip files and upload those.`
        : `${name} was not added: it holds more than ${MAX_ZIP_FILES.toLocaleString("en")} files. Split it into smaller zip files and upload those.`,
    )
  }
}

export { ZipFormatError }

export async function ingestZipDocument(matterDir: string, fileName: string, buffer: Buffer, options: ContainerOptions = {}) {
  const zip = parseZip(buffer)
  const budget = options.budget ?? archiveBudget()
  // Checked against the sizes the archive states before anything is added; each
  // file is then unpacked only up to its stated size (dochaus/lib/zip.ts).
  if (zip.uncompressed > budget.bytes) throw new ZipTooLargeError(fileName, "bytes")
  if (zip.entries.length > budget.files) throw new ZipTooLargeError(fileName, "files")
  budget.bytes -= zip.uncompressed
  budget.files -= zip.entries.length
  const parts = zip.entries.map(({ encrypted, unsafePath, ...entry }): ContainerPart => ({
    ...entry,
    ...(encrypted ? { refused: "encrypted" as const } : unsafePath ? { refused: "unsafe_path" as const } : {}),
  }))
  return ingestContainerDocument(matterDir, fileName, buffer, parts, 0, { ...options, budget })
}
