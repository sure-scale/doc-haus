import { inflateSync } from "fflate"
import { EMAIL_ATTACHMENT_EXTENSIONS, extension, formatSize, oneLine, safeAttachmentName } from "./email"

// Zip sources (.zip) are containers, like an email's attachments: each file
// inside can become a matter document of its own (services/ingest/src/zip-archive.ts).
// This is the one zip reader: ingest indexes the archive's text and takes its
// entries from it, and liveText (extract.ts) re-extracts the same text to check
// citation offsets, so the text depends on the archive's bytes alone. The text
// is a listing of the files inside, one per line with its size, the zip's own
// counterpart of an email's Attachments line. It carries no status wording.
//
// The central directory is read here rather than through fflate's unzip, which
// does not say whether an entry is encrypted and stops at the first entry it
// cannot inflate. Each entry is inflated on its own, only when asked for.

export class ZipFormatError extends Error {}

export type ZipEntry = {
  // The entry's path in the archive, on one line.
  filename: string
  // A safe matter-document name derived from the path's last segment.
  name: string
  mimeType: string
  // The uncompressed size the archive states.
  size: number
  // Whether `name` is a format a matter document can take.
  supported: boolean
  // Password-protected: its bytes cannot be read.
  encrypted: boolean
  // An absolute path or one that climbs out of the archive (`..`).
  unsafePath: boolean
  // The entry's bytes. Throws ZipFormatError when they cannot be read.
  content: () => Uint8Array
}

export type ParsedZip = {
  text: string
  // Every file in the archive, in order. Folders and macOS metadata
  // (__MACOSX/, .DS_Store) are left out.
  entries: ZipEntry[]
  // The stated uncompressed size of all `entries`.
  uncompressed: number
}

const ENTRY_TYPE: Record<string, string> = {
  ".pdf": "application/pdf",
  ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  ".eml": "message/rfc822",
  ".zip": "application/zip",
}

const EOCD = 0x06054b50
const CENTRAL = 0x02014b50
const LOCAL = 0x04034b50

export function parseZip(input: Buffer | Uint8Array): ParsedZip {
  const bytes = input instanceof Uint8Array ? input : new Uint8Array(input)
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  const fail = (): never => {
    throw new ZipFormatError("The file is not a readable zip archive.")
  }
  const u16 = (at: number) => (at + 2 <= bytes.length ? view.getUint16(at, true) : fail())
  const u32 = (at: number) => (at + 4 <= bytes.length ? view.getUint32(at, true) : fail())
  // The end-of-central-directory record sits in the last 22 bytes plus at most
  // a 64 KB comment.
  let end = -1
  for (let at = bytes.length - 22; at >= Math.max(0, bytes.length - 22 - 0xffff); at--) {
    if (view.getUint32(at, true) === EOCD) {
      end = at
      break
    }
  }
  if (end < 0) fail()
  const count = u16(end + 10)
  let at = u32(end + 16)
  const entries: ZipEntry[] = []
  for (let index = 0; index < count; index++) {
    if (u32(at) !== CENTRAL) fail()
    const flags = u16(at + 8)
    const method = u16(at + 10)
    let compressed = u32(at + 20)
    let size = u32(at + 24)
    const nameLength = u16(at + 28)
    const extraLength = u16(at + 30)
    const commentLength = u16(at + 32)
    let local = u32(at + 42)
    if (at + 46 + nameLength + extraLength > bytes.length) fail()
    const rawName = bytes.subarray(at + 46, at + 46 + nameLength)
    // Zip64 sizes and offset, for an entry whose 32-bit fields are saturated.
    for (let extra = at + 46 + nameLength; extra + 4 <= at + 46 + nameLength + extraLength; ) {
      const id = u16(extra)
      const length = u16(extra + 2)
      if (id === 0x0001) {
        let field = extra + 4
        const next = () => {
          const value = Number(field + 8 <= bytes.length ? view.getBigUint64(field, true) : fail())
          field += 8
          return value
        }
        if (size === 0xffffffff) size = next()
        if (compressed === 0xffffffff) compressed = next()
        if (local === 0xffffffff) local = next()
      }
      extra += 4 + length
    }
    at += 46 + nameLength + extraLength + commentLength
    const path = decodeName(rawName, (flags & 0x800) !== 0)
    const segments = path.split(/[\\/]/)
    const folder = /[\\/]$/.test(path)
    const metadata = segments.includes("__MACOSX") || segments.at(-1) === ".DS_Store"
    if (folder || metadata) continue
    const filename = oneLine(path)
    const name = safeAttachmentName(segments.at(-1), ENTRY_TYPE[extension(segments.at(-1) ?? "")] ?? "", index + 1)
    entries.push({
      filename,
      name,
      mimeType: ENTRY_TYPE[extension(name)] ?? "application/octet-stream",
      size,
      supported: (EMAIL_ATTACHMENT_EXTENSIONS as readonly string[]).includes(extension(name)),
      encrypted: (flags & 0x1) !== 0,
      unsafePath: /^([\\/]|[a-z]:)/i.test(path) || segments.includes(".."),
      content: () => entryContent(bytes, view, local, method, compressed, size),
    })
  }
  return {
    text: entries.length ? `Contents:\n${entries.map((entry) => `${entry.filename} (${formatSize(entry.size)})`).join("\n")}` : "",
    entries,
    uncompressed: entries.reduce((sum, entry) => sum + entry.size, 0),
  }
}

export function zipToText(buffer: Buffer | Uint8Array) {
  return parseZip(buffer).text
}

// Inflate one entry into a buffer one byte longer than its stated size, so an
// entry that expands past what the archive states is caught, not followed.
function entryContent(bytes: Uint8Array, view: DataView, local: number, method: number, compressed: number, size: number) {
  const unreadable = () => new ZipFormatError("A file in the zip archive could not be read.")
  if (local + 30 > bytes.length || view.getUint32(local, true) !== LOCAL) throw unreadable()
  const start = local + 30 + view.getUint16(local + 26, true) + view.getUint16(local + 28, true)
  if (start + compressed > bytes.length) throw unreadable()
  const data = bytes.subarray(start, start + compressed)
  if (method === 0) {
    if (compressed !== size) throw unreadable()
    return data.slice()
  }
  if (method !== 8) throw unreadable()
  let out: Uint8Array
  try {
    out = inflateSync(data, { out: new Uint8Array(size + 1) })
  } catch {
    throw unreadable()
  }
  if (out.length !== size) throw unreadable()
  return out
}

// UTF-8 when the archive says so or the bytes decode as it (macOS writes UTF-8
// without the flag); otherwise the legacy single-byte encoding, read as Latin-1.
function decodeName(raw: Uint8Array, utf8: boolean) {
  if (utf8) return new TextDecoder().decode(raw)
  try {
    return new TextDecoder("utf-8", { fatal: true }).decode(raw)
  } catch {
    return new TextDecoder("latin1").decode(raw)
  }
}
