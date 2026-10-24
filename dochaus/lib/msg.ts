import MsgReader, { type FieldsData } from "@kenjiuno/msgreader"
import iconv from "iconv-lite"
import PostalMime, { type Address } from "postal-mime"
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
//
// Ingest reads uploads in its own process, so a damaged or hostile file must
// not hold it: the compound file's structure is checked before the reader
// follows it, and an RTF body is read only within the limits below. A body past
// them, or one that cannot be decoded, is marked unreadable, which the
// extraction quality gate reports; a property of a type it never has is left
// out, never read as another.

export class MsgFormatError extends Error {}

const UNREADABLE = "The file is not a readable Outlook email."

// PidTagAttachMethod (MS-OXCMSG 2.2.2.9): how an attachment's content is held.
// By reference (2, 3, 4) or by web reference (7) is a link to a file elsewhere.
const ATTACH_OLE = 6
const ATTACH_LINKED = new Set([2, 3, 4, 7])
// PidTagAttachFlags: ATT_MHTML_REF, an attachment the HTML body shows.
const ATT_MHTML_REF = 4
// PidTagMessageFlags: MSGFLAG_UNSENT, a draft never sent.
const MSGFLAG_UNSENT = 8
const OUTLOOK_TYPE = EXTENSION_TYPE[".msg"]!
const IMAGE_NAME = /\.(png|jpe?g|gif|bmp|webp|tiff?)$/i

export async function parseMsg(buffer: Buffer | Uint8Array): Promise<ParsedEmail> {
  const { reader, fields } = readMsg(buffer)
  try {
    return await emailOf(reader, fields)
  } catch {
    // The properties are checked as they are read; a structure the reader
    // returned in some other shape is still a file that cannot be read.
    throw new MsgFormatError(UNREADABLE)
  }
}

export async function msgToText(buffer: Buffer | Uint8Array) {
  return (await parseMsg(buffer)).text
}

async function emailOf(reader: MsgReader, fields: FieldsData): Promise<ParsedEmail> {
  const body = msgBody(fields)
  const shown = cidReferences(body.html)
  const attachments: EmailAttachment[] = []
  let embedded = 0
  for (const attachment of fields.attachments ?? []) {
    const part = msgAttachment(reader, attachment, attachments.length + 1, shown)
    if (part) attachments.push(part)
    else embedded++
  }
  const known = await transportAddresses(str(fields.headers))
  // A draft never sent has no date it was sent unless it states one; its
  // delivery time is when it was saved.
  const unsent = ((num(fields.messageFlags) ?? 0) & MSGFLAG_UNSENT) !== 0
  const headers = {
    subject: str(fields.subject),
    from: sender(fields, known),
    to: recipients(fields, "to", known),
    cc: recipients(fields, "cc", known),
    date: isoDate(str(fields.clientSubmitTime) ?? (unsent ? undefined : str(fields.messageDeliveryTime))),
  }
  return {
    text: emailText(headers, attachments, body.text),
    body: body.text,
    attachments,
    embedded,
    ...(body.unreadable ? { bodyUnreadable: true as const } : {}),
  }
}

// A file that is not an Outlook item (not a compound file, a compound file with
// no message in it, or one cut short) is refused with MsgFormatError.
function readMsg(buffer: Buffer | Uint8Array) {
  if (!compoundFileEnds(buffer)) throw new MsgFormatError(UNREADABLE)
  // Exactly the file's bytes: the reader views the whole buffer it is given, and
  // a Buffer can be a slice of a larger pool.
  const bytes = buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength) as ArrayBuffer
  const read = openMsg(bytes)
  if (!read) throw new MsgFormatError(UNREADABLE)
  // A string stored as ANSI (PT_STRING8, type 001E) is in the message's
  // codepage, which the reader learns only from the message: read it again with
  // that codepage. One the reader cannot decode keeps the first reading.
  if (!hasAnsiStrings(buffer)) return read
  const codepage = num(read.fields.messageCodepage) ?? num(read.fields.internetCodepage)
  return openMsg(bytes, iconvName(codepage) ?? "cp1252") ?? read
}

function openMsg(bytes: ArrayBuffer, ansiEncoding?: string) {
  const reader = new MsgReader(bytes)
  reader.parserConfig = { includeRawProps: true, ...(ansiEncoding ? { ansiEncoding } : {}) }
  try {
    const fields = reader.getFileData()
    // Every Outlook item states its class (IPM.Note for an email).
    if (fields.dataType === "msg" && str(fields.messageClass)) return { reader, fields }
  } catch {
    // Not a compound file the reader can follow.
  }
}

// A property stream named for an ANSI string, in any storage of the file.
function hasAnsiStrings(buffer: Buffer | Uint8Array) {
  const bytes = Buffer.from(buffer.buffer, buffer.byteOffset, buffer.byteLength)
  return ["001E", "001e"].some((type) => bytes.includes(Buffer.from(type, "utf16le")))
}

// The reader follows the compound file's sector chains and directory tree with
// no guard against a loop, so a damaged file whose chain or tree points back on
// itself would hold the parse, and the ingest service, forever; and it reads
// each stream an entry names, so entries naming one stream's sectors many times
// over would make a small file read as a huge one. This walks the same chains
// and tree as the reader (its Reader.parse and readProperty, which it mirrors)
// and refuses a file where one revisits a sector or an entry, runs longer than
// the file allows or reads past its end, or whose streams together are larger
// than the file holds (MS-CFB 2.3 to 2.6).
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

  // The mini stream (the root's stream), and each stream: a small one through
  // the mini FAT, a large one no longer than the file. In a sound file no two
  // streams share a sector, so the small ones together fit in the mini stream
  // and the large ones in the file.
  if (root.size < 0 || root.size > bytes.length || !chain(root.start, nextSector, sectors)) return false
  const nextMini = nextIn([...miniFat])
  let small = 0
  let large = 0
  for (const index of reached) {
    const entry = entries[index]!
    if (entry.type !== DOCUMENT || !entry.size) continue
    if (entry.size < 0) return false
    if (entry.size >= MINI_STREAM_CUTOFF) large += entry.size
    else if (!chain(entry.start, nextMini, MINI_STREAM_CUTOFF / MINI_SECTOR)) return false
    else small += entry.size
  }
  return small <= root.size && large <= bytes.length
}

// One attachment as the message states it, or undefined for an image the HTML
// body shows inline (a logo, a signature), which is counted, not listed. An OLE
// object or a linked file is listed and not read; an attachment whose content
// cannot be read out of the message is listed as failed.
function msgAttachment(reader: MsgReader, attachment: FieldsData, position: number, shown: ReadonlySet<string>): EmailAttachment | undefined {
  const method = num(rawProp(attachment, "3705"))
  const filename = str(attachment.fileName) ?? str(attachment.fileNameShort) ?? str(attachment.name)
  const type = str(attachment.attachMimeTag) || EXTENSION_TYPE[extension(filename ?? "")]
  if (shownInline(attachment, filename, type, shown)) return
  const linked = method !== undefined && ATTACH_LINKED.has(method)
  // Bytes held in the attachment, or a storage (an attached message or an OLE
  // object) written out as a compound file of its own; a link has neither.
  const storage = attachment.innerMsgContent === true
  const content = !linked && (storage || attachment.dataId !== undefined) ? readAttachment(reader, attachment) : undefined
  // A storage's size is the one the message states (PidTagAttachSize) when it
  // states one: the written-out copy is the reader's own, not the sender's.
  const size = storage ? num(rawProp(attachment, "0e20")) : undefined
  if (storage && method !== ATTACH_OLE) {
    // An attached Outlook message is named by its display name (its subject),
    // and saved, as Outlook saves it, as a .msg file. A slash in a subject
    // ("RE: 12/03 call") is not a folder.
    const subject = str(attachment.name)
    const part = emailAttachment(subject, OUTLOOK_TYPE, content ?? new Uint8Array(), position)
    const name = safeAttachmentName(subject ? `${subject.replace(/[\\/]/g, "-")}.msg` : "", OUTLOOK_TYPE, position)
    return { ...part, name, size: size ?? part.size, supported: true, ...(content ? {} : { refused: "failed" as const }) }
  }
  const part = emailAttachment(filename, type, content ?? new Uint8Array(), position)
  if (linked) return { ...part, supported: false, note: "linked" }
  if (method === ATTACH_OLE) return { ...part, size: size ?? part.size, supported: false, note: "ole" }
  return content ? part : { ...part, refused: "failed" }
}

// An image the HTML body shows: one with a Content-ID that Outlook hides from
// the attachment list, marks as shown by the HTML (ATT_MHTML_REF), or that the
// HTML names by that Content-ID (one of `shown`).
function shownInline(attachment: FieldsData, filename: string | undefined, type: string | undefined, shown: ReadonlySet<string>) {
  const id = str(attachment.pidContentId)
  if (!id || !(type?.toLowerCase().startsWith("image/") || IMAGE_NAME.test(filename ?? ""))) return false
  return attachment.attachmentHidden === true || ((num(rawProp(attachment, "3714")) ?? 0) & ATT_MHTML_REF) !== 0 || shown.has(id)
}

// The Content-IDs the HTML body names (src="cid:..."), read once for every
// attachment.
function cidReferences(html: string | undefined) {
  return new Set(Array.from(html?.matchAll(/cid:([^\s"'<>()]+)/gi) ?? [], (match) => match[1]!))
}

function readAttachment(reader: MsgReader, attachment: FieldsData) {
  try {
    const content: unknown = reader.getAttachment(attachment).content
    return content instanceof Uint8Array ? content : undefined
  } catch {
    return undefined
  }
}

// From: who the message is from (PidTagSentRepresenting*: the person a
// delegate sent it for, as an .eml's From names them), else its sender.
// Exchange states its own sender as the person sent for too, often with the
// SMTP address on the sender only: the same name or directory path takes it.
function sender(fields: FieldsData, known: KnownAddresses) {
  const name = str(rawProp(fields, "0042"))
  const directory = str(rawProp(fields, "0065"))
  const representing = address(name, str(fields.sentRepresentingSmtpAddress), directory, str(rawProp(fields, "0064")), known)
  const from = address(str(fields.senderName), str(fields.senderSmtpAddress), str(fields.senderEmail), str(fields.senderAddressType), known)
  const same = (one: string | undefined, other: string | undefined) => !!one?.trim() && one.trim().toLowerCase() === other?.trim().toLowerCase()
  if (representing && !representing.address && from?.address && (same(name, str(fields.senderName)) || same(directory, str(fields.senderEmail))))
    return address(name, from.address, undefined, undefined, known)
  return representing ?? from
}

function recipients(fields: FieldsData, type: "to" | "cc", known: KnownAddresses) {
  return (fields.recipients ?? [])
    .filter((recipient) => recipient.recipType === type)
    .map((recipient) => address(str(recipient.name), str(recipient.smtpAddress), str(recipient.email), str(recipient.addressType), known))
    .filter((recipient): recipient is EmailAddress => recipient !== undefined)
}

// A name and an SMTP address, as an .eml's From or To states them. An Exchange
// address (a directory path) is not an email address: the address the
// message's transport headers give the same name stands in, else the name
// alone. A name that only repeats the address is left out, as a bare address in
// an .eml has none.
function address(name: string | undefined, smtp: string | undefined, email: string | undefined, type: string | undefined, known: KnownAddresses): EmailAddress | undefined {
  const addr = smtp || (type?.toUpperCase() === "SMTP" ? email : undefined) || (name ? known.get(name.trim().toLowerCase()) : undefined) || undefined
  const shown = name && name.toLowerCase() !== addr?.toLowerCase() ? name : undefined
  if (!shown && !addr) return
  return { name: shown, address: addr }
}

// The address of each name in the message's transport headers
// (PidTagTransportMessageHeaders), which a message that came in over SMTP
// keeps, by its trimmed lower-case name. A name listed with two addresses has
// none: which is whose is not stated.
type KnownAddresses = ReadonlyMap<string, string | undefined>

async function transportAddresses(headers: string | undefined): Promise<KnownAddresses> {
  const known = new Map<string, string | undefined>()
  if (!headers?.trim()) return known
  try {
    const parsed = await PostalMime.parse(`${headers.replace(/\s+$/, "")}\r\n\r\n`)
    const listed: Address[] = [parsed.from, parsed.sender, ...(parsed.to ?? []), ...(parsed.cc ?? []), ...(parsed.bcc ?? [])].filter(
      (entry): entry is Address => entry !== undefined,
    )
    for (const entry of listed.flatMap((entry) => (entry.group ? entry.group : [entry]))) {
      const name = entry.name?.trim().toLowerCase()
      if (!name || !entry.address) continue
      if (!known.has(name)) known.set(name, entry.address)
      else if (known.get(name)?.toLowerCase() !== entry.address.toLowerCase()) known.set(name, undefined)
    }
  } catch {
    // Headers it cannot read give no addresses.
  }
  return known
}

// MAPI times read as UTC strings; an .eml's Date reads as ISO 8601.
function isoDate(value: string | undefined) {
  if (!value) return
  const date = new Date(value)
  return Number.isNaN(date.getTime()) ? undefined : date.toISOString()
}

// The message body as text, the HTML it came from (for the images the HTML
// shows), and whether a body it has could not be read.
type MsgBody = { text: string; html?: string; unreadable?: true }

function msgBody(fields: FieldsData): MsgBody {
  const plain: unknown = fields.body
  const stated: unknown = fields.bodyHtml ?? fields.html
  const html = typeof stated === "string" ? stated : stated instanceof Uint8Array ? decode(stated, num(fields.internetCodepage) ?? 65001) : undefined
  if (typeof plain === "string" && plain.trim()) return { text: plain, html }
  if (html?.trim()) return { text: htmlBodyText(html), html }
  const compressed: unknown = fields.compressedRtf
  if (compressed instanceof Uint8Array) return rtfBody(compressed) ?? { text: "", unreadable: true }
  // A body stored as a type no body has (a number, say) is one that cannot be read.
  const misstated = (plain !== undefined && typeof plain !== "string") || (stated !== undefined && html === undefined) || compressed !== undefined
  return misstated ? { text: "", unreadable: true } : { text: "" }
}

// Compressed RTF (MS-OXRTFCP) inflates up to eight times its size, so a small
// file can hold a huge body: one that would inflate past this is not read.
const MAX_RTF_BYTES = 32 * 1024 * 1024
// The RTF readers below keep one state per open group, and the encapsulated
// reader's cost grows with the depth each group opens at, so a body nested
// deep enough would hold the parse for minutes. Each group counts its depth;
// past this total (a second or so of reading) the body is not read. A real
// message stays far below it: its groups nest a few levels deep.
const MAX_RTF_NESTING = 32_000_000

// The compressed RTF body's text, and the HTML it encapsulates, or undefined
// when it cannot be read.
function rtfBody(compressed: Uint8Array): MsgBody | undefined {
  const rtf = inflateRtf(compressed)
  if (!rtf) return
  const shape = rtfShape(rtf)
  if (!shape) return
  if (shape.encapsulated) {
    try {
      const { mode, text } = deEncapsulateSync(rtf, { decode: decodeNamed, warn: () => {}, allowCp0: true, maxTokenSize: rtf.length })
      return mode === "html" ? { text: htmlBodyText(String(text)), html: String(text) } : { text: String(text) }
    } catch {
      // Not encapsulated HTML or text after all: a native RTF body.
    }
  }
  try {
    return { text: rtfText(rtf) }
  } catch {
    return
  }
}

// The dictionary LZFu compression starts from (MS-OXRTFCP 2.1.3.1.1).
const RTF_DICTIONARY = Buffer.from(
  "{\\rtf1\\ansi\\mac\\deff0\\deftab720{\\fonttbl;}{\\f0\\fnil \\froman \\fswiss \\fmodern \\fscript \\fdecor MS Sans SerifSymbolArialTimes New RomanCourier{\\colortbl\\red0\\green0\\blue0\r\n\\par \\pard\\plain\\f0\\fs20\\b\\i\\u\\tab\\tx",
  "latin1",
)
const LZFU = 0x75465a4c
const UNCOMPRESSED = 0x414c454d

// The RTF a compressed RTF property holds (MS-OXRTFCP 2.1.3.1), or undefined
// when its header is wrong or it would inflate past MAX_RTF_BYTES. The data is
// a run of control bytes, each saying which of the next eight tokens are
// literal bytes and which are references (12-bit offset, 4-bit length) into
// the last 4096 bytes written; a reference to the write position ends it.
function inflateRtf(compressed: Uint8Array) {
  if (compressed.length < 16) return
  const view = new DataView(compressed.buffer, compressed.byteOffset, compressed.byteLength)
  // The stated compressed size counts the bytes after its own field, and must
  // fit the property.
  const end = view.getUint32(0, true) + 4
  const rawSize = view.getUint32(4, true)
  const type = view.getUint32(8, true)
  if (end < 16 || end > compressed.length) return
  if (type === UNCOMPRESSED) return rawSize <= MAX_RTF_BYTES ? Buffer.from(compressed.subarray(16, Math.min(end, 16 + rawSize))) : undefined
  if (type !== LZFU) return
  const recent = new Uint8Array(4096)
  recent.set(RTF_DICTIONARY)
  let position = RTF_DICTIONARY.length
  // Room for the size the header states, within what the data can inflate to.
  let out = new Uint8Array(Math.min(Math.max(rawSize, 4096), 8 * end, MAX_RTF_BYTES))
  let written = 0
  const put = (byte: number) => {
    if (written === out.length) {
      if (out.length >= MAX_RTF_BYTES) return false
      const grown = new Uint8Array(Math.min(out.length * 2, MAX_RTF_BYTES))
      grown.set(out)
      out = grown
    }
    out[written++] = byte
    recent[position++ & 4095] = byte
    return true
  }
  for (let at = 16; at < end; ) {
    const control = compressed[at++]!
    for (let bit = 0; bit < 8 && at < end; bit++) {
      if (!(control & (1 << bit))) {
        if (!put(compressed[at++]!)) return
        continue
      }
      if (at + 2 > end) break
      const reference = (compressed[at]! << 8) | compressed[at + 1]!
      at += 2
      const offset = reference >> 4
      if (offset === (position & 4095)) return Buffer.from(out.buffer, 0, written)
      const length = (reference & 15) + 2
      for (let index = 0; index < length; index++) if (!put(recent[(offset + index) & 4095]!)) return
    }
  }
  return Buffer.from(out.buffer, 0, written)
}

type RtfToken = { type: number; word?: string; param?: number; data?: Buffer }

// Runs the tokenizer the RTF readers use over the whole body, handing each
// token to `handle` as it is read.
function tokenize(rtf: Buffer, handle: (token: RtfToken) => void) {
  const tokenizer = new Tokenize({ maxTokenSize: rtf.length })
  const fail = (error?: unknown) => {
    if (error) throw error
  }
  tokenizer.push = (token: RtfToken) => {
    handle(token)
    return true
  }
  tokenizer._transform(rtf, undefined, fail)
  tokenizer._flush(fail)
}

class TooDeep extends Error {}

// Whether the body wraps HTML or text (deEncapsulateSync reads it as such only
// when \fromhtml or \fromtext comes among its first tokens), or undefined when
// it nests past MAX_RTF_NESTING.
function rtfShape(rtf: Buffer) {
  let depth = 0
  let work = 0
  let count = 0
  let encapsulated = false
  try {
    tokenize(rtf, (token) => {
      count++
      if (token.type === 0) {
        work += ++depth
        if (work > MAX_RTF_NESTING) throw new TooDeep()
      } else if (token.type === 1) depth = Math.max(0, depth - 1)
      else if (count <= 12 && (token.word === "fromhtml" || token.word === "fromtext")) encapsulated = true
    })
  } catch (error) {
    if (error instanceof TooDeep) return
    // A body the tokenizer stops on is left to the readers, which stop on it too.
  }
  return { encapsulated }
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
  emspace: " ",
  enspace: " ",
  "~": " ",
  _: "‑",
  "{": "{",
  "}": "}",
  "\\": "\\",
}

// The codepage of a font's \fcharset (the Windows character sets). ANSI (0),
// default (1) and symbol (2) fonts use the document's \ansicpg.
const CHARSET_CODEPAGE: Record<number, number> = {
  77: 10000,
  128: 932,
  129: 949,
  134: 936,
  136: 950,
  161: 1253,
  162: 1254,
  163: 1258,
  177: 1255,
  178: 1256,
  186: 1257,
  204: 1251,
  222: 874,
  238: 1250,
  255: 437,
}

// The text of a native RTF body: paragraphs as lines, special characters as
// themselves, tables and metadata left out. Formatting is dropped. Bytes
// written as \'xx or as text are in the codepage of the font they are set in
// (its \cpg or \fcharset in the font table), else the document's \ansicpg.
function rtfText(rtf: Buffer) {
  type Group = { hidden: boolean; uc: number; font?: number; fontTable: boolean; defining?: number }
  const stack: Group[] = []
  let group: Group = { hidden: false, uc: 1, fontTable: false }
  let ansi = 1252
  let defaultFont: number | undefined
  const fontCodepage = new Map<number, number>()
  const fontCharset = new Map<number, number>()
  const parts: string[] = []
  // Bytes not yet decoded, all in one codepage.
  let bytes: Buffer[] = []
  let bytesCodepage = ansi
  // Characters still to skip after a \u character: its stand-ins for readers
  // without Unicode.
  let skip = 0
  let opensGroup = false
  const codepage = () => {
    const font = group.font ?? defaultFont
    return (font === undefined ? undefined : fontCodepage.get(font) ?? fontCharset.get(font)) ?? ansi
  }
  const flush = () => {
    if (bytes.length) parts.push(decode(Buffer.concat(bytes), bytesCodepage))
    bytes = []
  }
  const add = (data: Buffer) => {
    if (group.hidden || !data.length) return
    const current = codepage()
    if (bytes.length && current !== bytesCodepage) flush()
    bytesCodepage = current
    bytes.push(data)
  }
  const emit = (text: string) => {
    if (group.hidden) return
    flush()
    parts.push(text)
  }
  tokenize(rtf, (token) => {
    const first = opensGroup
    opensGroup = token.type === 0
    if (token.type === 0 || token.type === 1) {
      skip = 0
      flush()
      if (token.type === 0) {
        stack.push(group)
        group = { ...group }
      } else group = stack.pop() ?? group
      return
    }
    if (token.type === 3) {
      const data = token.data ?? Buffer.alloc(0)
      const dropped = Math.min(skip, data.length)
      skip -= dropped
      add(data.subarray(dropped))
      return
    }
    const word = token.word ?? ""
    if (first && (word === "*" || HIDDEN_DESTINATIONS.has(word))) group.hidden = true
    if (first && word === "fonttbl") group.fontTable = true
    if (skip > 0 && word === "'") {
      skip--
      return
    }
    if (group.fontTable) {
      // A font table entry: \fN starts font N, its \fcharset or \cpg its codepage.
      if (word === "f" && token.param !== undefined) group.defining = token.param
      else if (word === "fcharset" && group.defining !== undefined && token.param !== undefined && CHARSET_CODEPAGE[token.param])
        fontCharset.set(group.defining, CHARSET_CODEPAGE[token.param]!)
      else if (word === "cpg" && group.defining !== undefined && token.param) fontCodepage.set(group.defining, token.param)
      return
    }
    if (word === "ansicpg" && token.param) ansi = token.param
    else if (word === "deff" && token.param !== undefined) defaultFont = token.param
    else if (word === "f" && token.param !== undefined) group.font = token.param
    else if (word === "plain") group.font = undefined
    else if (word === "uc") group.uc = token.param ?? 1
    else if (word === "u" && token.param !== undefined) {
      emit(String.fromCharCode(token.param < 0 ? token.param + 0x10000 : token.param))
      skip = group.uc
    } else if (word === "'" && token.data) add(token.data)
    else if (SYMBOLS[word] !== undefined) emit(SYMBOLS[word]!)
  })
  flush()
  return parts.join("")
}

// MAPI and RTF codepages by number, under the names iconv-lite gives them; one
// it lacks is read with the WHATWG decoder when that has it, else as
// Windows-1252.
const CODEPAGE_NAMES: Record<number, string> = {
  1200: "utf16le",
  1201: "utf16be",
  10000: "macroman",
  20127: "ascii",
  20932: "eucjp",
  51932: "eucjp",
  54936: "gb18030",
  65000: "utf7",
  65001: "utf8",
}
const WEB_CODEPAGES: Record<number, string> = { 50220: "iso-2022-jp", 50221: "iso-2022-jp", 50222: "iso-2022-jp" }

function iconvName(codepage: number | undefined) {
  if (!codepage) return
  const name = CODEPAGE_NAMES[codepage] ?? `cp${codepage}`
  return iconv.encodingExists(name) ? name : undefined
}

function decode(bytes: Uint8Array, codepage: number | undefined) {
  const web = codepage ? WEB_CODEPAGES[codepage] : undefined
  if (web) return new TextDecoder(web).decode(bytes)
  return iconv.decode(Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength), iconvName(codepage) ?? "cp1252")
}

// The encapsulated reader names a codepage as "cp1252".
function decodeNamed(bytes: Buffer, encoding: string) {
  return decode(bytes, Number(encoding.replace(/^cp/, "")) || undefined)
}

// A value the reader returned, if it has the type the property has.
function str(value: unknown) {
  return typeof value === "string" ? value : undefined
}

function num(value: unknown) {
  return typeof value === "number" && Number.isFinite(value) ? value : undefined
}

// A property the reader keeps only among its raw properties, by its id (the
// tag's first four hex digits).
function rawProp(node: FieldsData, id: string): unknown {
  return node.rawProps?.find((prop) => prop.propertyTag?.startsWith(id))?.value
}
