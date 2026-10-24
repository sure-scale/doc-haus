import { decompressRTF } from "@kenjiuno/decompressrtf"
import MsgReader, { type FieldsData } from "@kenjiuno/msgreader"
import iconv from "iconv-lite"
import { deEncapsulateSync, Tokenize } from "rtf-stream-parser"
import {
  emailAttachment,
  emailText,
  EXTENSION_TYPE,
  extension,
  htmlBodyText,
  safeAttachmentName,
  type EmailAddress,
  type EmailAttachment,
  type ParsedEmail,
} from "./email"

// Outlook emails (.msg) are OLE compound files holding MAPI properties, not MIME.
// Outlook on Windows saves a dragged or "Save as" email this way. This reads one
// into the same shape parseEmail gives an .eml (email.ts), and writes its text
// with the same emailText, so the header block, the attachments line, child
// documents, de-duplication, safe names, the nesting limit and removal work as
// they do for .eml. Like parseEmail it is the one extractor for the format:
// ingest indexes its text and liveText (extract.ts) re-extracts it to check
// citation offsets, so the text depends on the file's bytes alone.
//
// The body is the plain-text body when the message has one, else its HTML body,
// else its compressed RTF body: decompressed, and either the HTML or text it
// encapsulates or, for a native RTF message, its text. HTML gets the same tag
// strip as an .eml's HTML body. An attached Outlook message is listed and read
// as an email of its own (one level per parse, like a forwarded .eml); an OLE
// embedded object, or an attachment held only as a link, is listed and never
// read.

export class MsgFormatError extends Error {}

// PidTagAttachMethod (MS-OXCMSG 2.2.2.9): how an attachment's content is held.
const ATTACH_METHOD = "37050003"
const ATTACH_OLE = 6
const OUTLOOK_TYPE = EXTENSION_TYPE[".msg"]!

export async function parseMsg(buffer: Buffer | Uint8Array): Promise<ParsedEmail> {
  const { reader, fields } = readMsg(buffer)
  const attachments: EmailAttachment[] = []
  let embedded = 0
  for (const attachment of fields.attachments ?? []) {
    const part = msgAttachment(reader, attachment, attachments.length + 1)
    if (part) attachments.push(part)
    else embedded++
  }
  const body = msgBody(fields)
  const headers = {
    subject: fields.subject,
    from: address(fields.senderName, fields.senderSmtpAddress, fields.senderEmail, fields.senderAddressType),
    to: recipients(fields, "to"),
    cc: recipients(fields, "cc"),
    date: isoDate(fields.clientSubmitTime ?? fields.messageDeliveryTime),
  }
  return { text: emailText(headers, attachments, body), body, attachments, embedded }
}

export async function msgToText(buffer: Buffer | Uint8Array) {
  return (await parseMsg(buffer)).text
}

// A file that is not an Outlook item (not a compound file, a compound file with
// no message in it, or one cut short) is refused with MsgFormatError.
function readMsg(buffer: Buffer | Uint8Array) {
  if (!compoundFileEnds(buffer)) throw new MsgFormatError("The file is not a readable Outlook email.")
  // A copy of exactly the file's bytes: the reader views the whole buffer it is
  // given, and a Buffer can be a slice of a larger pool.
  const reader = new MsgReader(Uint8Array.from(buffer).buffer)
  reader.parserConfig = { includeRawProps: true }
  let fields: FieldsData
  try {
    fields = reader.getFileData()
  } catch {
    throw new MsgFormatError("The file is not a readable Outlook email.")
  }
  // Every Outlook item states its class (IPM.Note for an email).
  if (fields.dataType !== "msg" || !fields.messageClass) throw new MsgFormatError("The file is not a readable Outlook email.")
  return { reader, fields }
}

// The reader follows the compound file's sector chains and directory tree with
// no guard against a loop, so a damaged file whose chain or tree points back on
// itself would hold the parse, and the ingest service, forever. This walks the
// same chains and tree as the reader (its Reader.parse and readProperty, which
// it mirrors) and refuses a file where one revisits a sector or an entry, runs
// longer than the file allows or reads past its end (MS-CFB 2.3 to 2.6).
const END_OF_CHAIN = -2
const NO_ENTRY = -1
const DIRECTORY = 1
const DOCUMENT = 2
const ROOT = 5
const MINI_SECTOR = 64
const MINI_STREAM_CUTOFF = 4096

function compoundFileEnds(bytes: Uint8Array) {
  try {
    return chainsEnd(bytes)
  } catch {
    // A sector or entry the reader would read from outside the file.
    return false
  }
}

function chainsEnd(bytes: Uint8Array) {
  if (bytes.length < 512) return false
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  // A read the reader would fail on is undefined here; the reader clamps the
  // position to the file first.
  const int = (offset: number) => {
    const at = Math.max(0, Math.min(bytes.length, offset))
    return at + 4 <= bytes.length ? view.getInt32(at, true) : undefined
  }
  const sectorSize = view.getInt8(30) === 12 ? 4096 : 512
  const perSector = sectorSize / 4
  const sectors = Math.ceil(bytes.length / sectorSize)
  const at = (sector: number) => (sector + 1) * sectorSize
  // A chain from `start` to its end through `next`, as the sectors it names.
  const chain = (start: number | undefined, next: (sector: number) => number | undefined, limit: number) => {
    const seen = new Set<number>()
    for (let sector = start; sector !== END_OF_CHAIN; sector = next(sector)) {
      if (sector === undefined || seen.has(sector) || seen.size >= limit) return
      seen.add(sector)
    }
    return [...seen]
  }
  const nextIn = (table: number[]) => (sector: number) => {
    const holder = table[Math.floor(sector / perSector)]
    return holder === undefined ? END_OF_CHAIN : int(at(holder) + 4 * (sector % perSector))
  }

  // The FAT's sectors: up to 109 named in the header, then those in the DIFAT chain.
  const fatCount = int(0x2c)!
  if (fatCount < 0) return false
  const fat = Array.from({ length: Math.min(fatCount, 109) }, (_, index) => int(0x4c + 4 * index)!)
  let remaining = fatCount - fat.length
  const difat = new Set<number>()
  for (let index = 0, sector = int(0x44)!; index < int(0x48)!; index++) {
    if (sector < 0 || sector >= sectors || difat.has(sector)) return false
    difat.add(sector)
    const count = Math.min(remaining, perSector - 1)
    for (let entry = 0; entry < count; entry++) {
      const holder = int(at(sector) + 4 * entry)
      if (holder === undefined) return false
      if (holder === NO_ENTRY || holder === END_OF_CHAIN) break
      fat.push(holder)
    }
    remaining -= count
    const next = int(at(sector) + 4 * (perSector - 1))
    if (next === undefined) return false
    if (next === NO_ENTRY || next === END_OF_CHAIN) break
    sector = next
  }
  const nextSector = nextIn(fat)

  // The mini FAT's sectors, as many as the header states (the reader also stops
  // at sector 0).
  const miniFat = new Set<number>()
  const miniFatCount = int(0x40)!
  for (let sector = int(0x3c); miniFat.size < miniFatCount && sector !== 0 && sector !== END_OF_CHAIN; sector = nextSector(sector)) {
    if (sector === undefined || miniFat.has(sector) || miniFat.size >= sectors) return false
    miniFat.add(sector)
  }

  // The directory's entries, from every sector of its chain.
  const directory = chain(int(0x30), nextSector, sectors)
  if (!directory) return false
  type Entry = { type: number; previous: number; next: number; child: number; start: number; size: number }
  const entries: Entry[] = []
  for (const sector of directory) {
    for (let offset = at(sector); offset < at(sector) + sectorSize && offset + 0x42 <= bytes.length; offset += 128) {
      if (offset + 0x42 === bytes.length) return false
      const type = view.getInt8(offset + 0x42)
      if (type !== DIRECTORY && type !== DOCUMENT && type !== ROOT) {
        entries.push({ type, previous: NO_ENTRY, next: NO_ENTRY, child: NO_ENTRY, start: 0, size: 0 })
        continue
      }
      const [previous, next, child, start, size] = [0x44, 0x48, 0x4c, 0x74, 0x78].map((field) => int(offset + field))
      if (previous === undefined || next === undefined || child === undefined || start === undefined || size === undefined) return false
      entries.push({ type, previous, next, child, start, size })
    }
  }

  // Every entry the tree reaches, each once: a storage's children are the
  // siblings around its child, and a child storage holds its own.
  const root = entries[0]
  if (!root) return false
  const reached = new Set<number>([0])
  const pending = [root]
  while (pending.length) {
    const storage = pending.pop()!
    const siblings = storage.child === NO_ENTRY ? [] : [storage.child]
    while (siblings.length) {
      const index = siblings.pop()!
      const entry = entries[index]
      if (!entry || reached.has(index)) return false
      reached.add(index)
      if (entry.type === DIRECTORY) pending.push(entry)
      if (entry.next !== NO_ENTRY) siblings.push(entry.next)
      if (entry.previous !== NO_ENTRY) siblings.push(entry.previous)
    }
  }

  // The mini stream, and each stream: a small one through the mini FAT, a large
  // one no longer than the file.
  if (!chain(root.start, nextSector, sectors)) return false
  const nextMini = nextIn([...miniFat])
  for (const index of reached) {
    const entry = entries[index]!
    if (entry.type !== DOCUMENT || !entry.size) continue
    if (entry.size < MINI_STREAM_CUTOFF ? !chain(entry.start, nextMini, MINI_STREAM_CUTOFF / MINI_SECTOR) : entry.size > bytes.length) return false
  }
  return true
}

// One attachment as the message states it, or undefined for an image the HTML
// body shows inline (a logo, a signature), which is counted, not listed. An OLE
// object or a linked file is listed and not read.
function msgAttachment(reader: MsgReader, attachment: FieldsData, position: number): EmailAttachment | undefined {
  // An integer property; the reader's type only lists its string and binary values.
  const method: unknown = attachment.rawProps?.find((prop) => prop.propertyTag === ATTACH_METHOD)?.value
  const filename = attachment.fileName ?? attachment.fileNameShort ?? attachment.name
  const type = attachment.attachMimeTag || EXTENSION_TYPE[extension(filename ?? "")]
  if (attachment.attachmentHidden && attachment.pidContentId && type?.toLowerCase().startsWith("image/")) return
  // Bytes held in the attachment, or a storage (an attached message or an OLE
  // object) written out as a compound file of its own; a link has neither.
  const content = attachment.innerMsgContent || attachment.dataId !== undefined ? readAttachment(reader, attachment) : undefined
  if (attachment.innerMsgContent && method !== ATTACH_OLE) {
    // An attached Outlook message is named by its display name (its subject),
    // and saved, as Outlook saves it, as a .msg file.
    const part = emailAttachment(attachment.name, OUTLOOK_TYPE, content ?? new Uint8Array(), position)
    const name = safeAttachmentName(attachment.name ? `${attachment.name}.msg` : "", OUTLOOK_TYPE, position)
    return { ...part, name, supported: content !== undefined }
  }
  const part = emailAttachment(filename, type, content ?? new Uint8Array(), position)
  return content && method !== ATTACH_OLE ? part : { ...part, supported: false }
}

function readAttachment(reader: MsgReader, attachment: FieldsData) {
  try {
    return reader.getAttachment(attachment).content
  } catch {
    return undefined
  }
}

function recipients(fields: FieldsData, type: "to" | "cc") {
  return (fields.recipients ?? [])
    .filter((recipient) => recipient.recipType === type)
    .map((recipient) => address(recipient.name, recipient.smtpAddress, recipient.email, recipient.addressType))
    .filter((recipient): recipient is EmailAddress => recipient !== undefined)
}

// A name and an SMTP address, as an .eml's From or To states them. An Exchange
// address (a directory path) is not an email address; the name alone stands.
// A name that only repeats the address is left out, as a bare address in an
// .eml has none.
function address(name: string | undefined, smtp: string | undefined, email: string | undefined, type: string | undefined): EmailAddress | undefined {
  const addr = smtp || (type?.toUpperCase() === "SMTP" ? email : undefined) || undefined
  const shown = name && name.toLowerCase() !== addr?.toLowerCase() ? name : undefined
  if (!shown && !addr) return
  return { name: shown, address: addr }
}

// MAPI times read as UTC strings; an .eml's Date reads as ISO 8601.
function isoDate(value: string | undefined) {
  if (!value) return
  const date = new Date(value)
  return Number.isNaN(date.getTime()) ? undefined : date.toISOString()
}

function msgBody(fields: FieldsData) {
  if (fields.body?.trim()) return fields.body
  const html = fields.bodyHtml ?? (fields.html ? decode(Buffer.from(fields.html), codepage(fields.internetCodepage)) : undefined)
  if (html?.trim()) return htmlBodyText(html)
  if (fields.compressedRtf) return rtfBody(fields.compressedRtf)
  return ""
}

// The compressed RTF body's text, or "" when it cannot be read, which leaves the
// email to the extraction quality gate as one with no readable body.
function rtfBody(compressed: Uint8Array) {
  // The stated compressed size must fit the property (MS-OXRTFCP 2.1.3.1.1), or
  // the decompressor reads past it.
  if (compressed.length < 16 || new DataView(compressed.buffer, compressed.byteOffset).getUint32(0, true) + 4 > compressed.length) return ""
  let rtf: Buffer
  try {
    rtf = Buffer.from(decompressRTF(compressed as unknown as number[]))
  } catch {
    return ""
  }
  try {
    const { mode, text } = deEncapsulateSync(rtf, { decode, warn: () => {}, allowCp0: true })
    return mode === "html" ? htmlBodyText(String(text)) : String(text)
  } catch {
    // Not encapsulated HTML or text: a native RTF body.
  }
  try {
    return rtfText(rtf)
  } catch {
    return ""
  }
}

// RTF codepages (\ansicpg, "cp1252") and MAPI ones (1252) by iconv-lite's names.
function codepage(value: number | undefined) {
  return value ? `cp${value}` : "cp65001"
}

function decode(bytes: Buffer, encoding: string) {
  const name = encoding === "cp65001" ? "utf8" : iconv.encodingExists(encoding) ? encoding : "cp1252"
  return iconv.decode(bytes, name)
}

// Groups whose text is not the message's: tables, metadata and pictures.
const HIDDEN_DESTINATIONS = new Set([
  "fonttbl",
  "colortbl",
  "stylesheet",
  "info",
  "pict",
  "object",
  "fldinst",
  "listtable",
  "listoverridetable",
  "revtbl",
  "rsidtbl",
  "filetbl",
  "xmlnstbl",
  "themedata",
  "colorschememapping",
  "datastore",
  "latentstyles",
])

const SYMBOLS: Record<string, string> = {
  par: "\n",
  line: "\n",
  sect: "\n",
  page: "\n",
  row: "\n",
  tab: "\t",
  cell: "\t",
  lquote: "‘",
  rquote: "’",
  ldblquote: "“",
  rdblquote: "”",
  bullet: "•",
  endash: "–",
  emdash: "—",
  emspace: " ",
  enspace: " ",
  "~": " ",
  _: "‑",
  "{": "{",
  "}": "}",
  "\\": "\\",
}

type RtfToken = { type: number; word?: string; param?: number; data?: Buffer }

// The text of a native RTF body: paragraphs as lines, special characters as
// themselves, tables and metadata left out. Formatting is dropped.
function rtfText(rtf: Buffer) {
  const tokens: RtfToken[] = []
  const tokenizer = new Tokenize()
  const fail = (error?: unknown) => {
    if (error) throw error
  }
  tokenizer.push = (token: RtfToken) => tokens.push(token) > 0
  tokenizer._transform(rtf, undefined, fail)
  tokenizer._flush(fail)

  type Group = { hidden: boolean; uc: number }
  const stack: Group[] = []
  let group: Group = { hidden: false, uc: 1 }
  let encoding = "cp1252"
  let out = ""
  let bytes: number[] = []
  // Characters still to skip after a \u character: its stand-ins for readers
  // without Unicode.
  let skip = 0
  let previous: RtfToken | undefined
  const flush = () => {
    if (bytes.length) out += decode(Buffer.from(bytes), encoding)
    bytes = []
  }
  const emit = (text: string) => {
    if (group.hidden) return
    flush()
    out += text
  }
  for (const token of tokens) {
    const opensGroup = previous?.type === 0
    previous = token
    if (token.type === 0 || token.type === 1) {
      skip = 0
      if (token.type === 0) {
        stack.push(group)
        group = { ...group }
      } else {
        flush()
        group = stack.pop() ?? group
      }
      continue
    }
    if (token.type === 3) {
      const data = token.data ?? Buffer.alloc(0)
      const dropped = Math.min(skip, data.length)
      skip -= dropped
      if (!group.hidden) bytes.push(...data.subarray(dropped))
      continue
    }
    const word = token.word ?? ""
    if (opensGroup && (word === "*" || HIDDEN_DESTINATIONS.has(word))) group.hidden = true
    if (skip > 0 && word === "'") {
      skip--
      continue
    }
    if (word === "ansicpg" && token.param) encoding = `cp${token.param}`
    else if (word === "uc") group.uc = token.param ?? 1
    else if (word === "u" && token.param !== undefined) {
      emit(String.fromCharCode(token.param < 0 ? token.param + 0x10000 : token.param))
      skip = group.uc
    } else if (word === "'" && token.data && !group.hidden) bytes.push(...token.data)
    else if (SYMBOLS[word] !== undefined) emit(SYMBOLS[word]!)
  }
  flush()
  return out
}
