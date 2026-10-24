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
