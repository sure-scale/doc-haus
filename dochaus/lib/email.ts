import PostalMime, { type Attachment } from "postal-mime"

// Email sources (.eml) are MIME documents, not zip-based Office files, so they go
// through a real RFC822 parser. This is the one email extractor: ingest indexes
// its text and records its attachments, and liveText (extract.ts) re-extracts the
// same text to check citation offsets, so the text must depend on the email's
// bytes alone, never on what became of an attachment. Outlook emails (.msg) are
// read into the same shape by msg.ts and written out by the same emailText, so
// both formats read alike everywhere downstream.
//
// The headers that carry legal meaning in a chain (who wrote, to whom, when,
// about what) sit above the body, followed by one line listing each attachment's
// name, type and size as the email states them. A forwarded message is kept as
// an attachment (forceRfc822Attachments) rather than flattened into the body, so
// it can become its own document with its own attachments, and so each parse
// reads one level of nesting only.

// The formats a matter document can take (ingest's MATTER_DOCUMENT_EXTENSIONS).
// An email or a zip file is a container: its parts become documents of their own.
export const EMAIL_ATTACHMENT_EXTENSIONS = [".docx", ".pdf", ".xlsx", ".eml", ".msg", ".zip"] as const

// The type of each format a matter document can take, for a file whose type is
// not stated (a zip entry, an Outlook attachment without one).
export const EXTENSION_TYPE: Record<string, string> = {
  ".pdf": "application/pdf",
  ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  ".eml": "message/rfc822",
  ".msg": "application/vnd.ms-outlook",
  ".zip": "application/zip",
}

const TYPE_EXTENSION: Record<string, string> = {
  ...Object.fromEntries(Object.entries(EXTENSION_TYPE).map(([ext, type]) => [type, ext])),
  "application/x-zip-compressed": ".zip",
}

// An attached email: a forwarded MIME message or an Outlook message.
const MESSAGE_TYPES = new Set(["message/rfc822", "application/vnd.ms-outlook"])

// Leaves room under the usual 255-byte file-name limit for a keep-both suffix.
const MAX_NAME_BYTES = 200

export type EmailAttachment = {
  // The name as the email states it, on one line; "" when it has none.
  filename: string
  // A safe matter-document name derived from the stated name and type.
  name: string
  mimeType: string
  size: number
  content: Uint8Array
  // Whether `name` is a format a matter document can take.
  supported: boolean
  // Outlook's TNEF wrapper, which can carry attachments of its own.
  tnef: boolean
  // An Outlook attachment held only as a link, or an OLE object: listed, never
  // read, whatever its name.
  note?: "linked" | "ole"
  // An Outlook attachment whose content could not be read out of the message.
  refused?: "failed"
}

export type ParsedEmail = {
  text: string
  // The message body alone, the end of `text` below its headers; "" when the
  // email has none (or its MIME structure hid it). The extraction quality gate
  // reads it apart from the headers.
  body: string
  attachments: EmailAttachment[]
  // Images embedded in the HTML body (logos, signatures): counted, not listed.
  embedded: number
  // The email has a body that could not be read (an Outlook RTF body that is
  // damaged or past the reading limits); `body` is then "".
  bodyUnreadable?: true
}

export type EmailAddress = { name?: string; address?: string }

// What an email's text is made of, whatever its file format.
export type EmailHeaders = {
  subject?: string
  from?: EmailAddress
  to?: EmailAddress[]
  cc?: EmailAddress[]
  // ISO 8601, as postal-mime gives a parseable Date header.
  date?: string
}

export async function parseEmail(buffer: Buffer | Uint8Array): Promise<ParsedEmail> {
  const email = await PostalMime.parse(buffer, { forceRfc822Attachments: true })
  // Only images the HTML body shows inline are left out; any other part,
  // even one the body references by Content-ID, is listed.
  const parts = email.attachments.filter((part) => !(part.related && part.contentId && part.mimeType?.toLowerCase().startsWith("image/")))
  const attachments = parts.map((part, index) => emailAttachment(part.filename, part.mimeType, partContent(part), index + 1))
  const body = email.text ?? (email.html ? htmlBodyText(email.html) : "")
  return { text: emailText(email, attachments, body), body, attachments, embedded: email.attachments.length - parts.length }
}

export async function emlToText(buffer: Buffer | Uint8Array) {
  return (await parseEmail(buffer)).text
}

// The email's text: the headers that carry legal meaning, one line listing the
// attachments as the email states them, then the body.
export function emailText(headers: EmailHeaders, attachments: EmailAttachment[], body: string) {
  const header = [
    headers.subject && `Subject: ${headers.subject}`,
    headers.from && `From: ${formatAddress(headers.from)}`,
    headers.to?.length && `To: ${headers.to.map(formatAddress).join(", ")}`,
    headers.cc?.length && `Cc: ${headers.cc.map(formatAddress).join(", ")}`,
    headers.date && `Date: ${headers.date}`,
    attachments.length &&
      `Attachments: ${attachments
        .map((attachment) => `${attachment.filename || "unnamed"} (${attachment.mimeType}, ${formatSize(attachment.size)})`)
        .join("; ")}`,
  ]
    .filter(Boolean)
    .join("\n")
  return header ? `${header}\n\n${body}` : body
}

// An HTML-only body keeps the plain tag strip until a real converter lands (#103):
// each "<" with at least one character before the next ">" becomes one space,
// through that ">". The same output as html.replace(/<[^>]+>/g, " "), whose
// retries after an unclosed "<" make it quadratic, so stored offsets into
// earlier text still hold.
export function htmlBodyText(html: string) {
  const parts: string[] = []
  let from = 0
  for (let open = html.indexOf("<"); open >= 0; open = html.indexOf("<", from)) {
    const close = html.indexOf(">", open + 1)
    // No ">" after this "<", so none after any later one: the rest stays.
    if (close < 0) break
    if (close === open + 1) {
      // "<>" is not a tag; the search goes on after its "<".
      parts.push(html.slice(from, open + 1))
      from = open + 1
      continue
    }
    parts.push(html.slice(from, open), " ")
    from = close + 1
  }
  parts.push(html.slice(from))
  return parts.join("")
}

function partContent(part: Attachment) {
  return typeof part.content === "string"
    ? new TextEncoder().encode(part.content)
    : part.content instanceof Uint8Array
      ? part.content
      : new Uint8Array(part.content)
}

// One attachment as the email states it: its name (null when it has none), its
// type (null when not stated) and its bytes.
export function emailAttachment(filename: string | null | undefined, type: string | null | undefined, content: Uint8Array, position: number): EmailAttachment {
  const mimeType = (type || "application/octet-stream").toLowerCase()
  const name = safeAttachmentName(filename, mimeType, position)
  return {
    filename: oneLine(filename ?? ""),
    name,
    mimeType,
    size: content.byteLength,
    content,
    supported: (EMAIL_ATTACHMENT_EXTENSIONS as readonly string[]).includes(extension(name)),
    tnef: mimeType === "application/ms-tnef" || /^winmail\.dat$/i.test(name),
  }
}

// The stated name is the sender's: take its last path segment, drop control and
// format characters (bidi overrides, zero-width marks) and leading dots, keep its extension, add one from the type when it
// has none, and fall back to a name from the type and position when nothing is
// left. Never refused, so no attachment is lost to its name.
export function safeAttachmentName(filename: string | null | undefined, mimeType: string, position: number) {
  const base = (filename ?? "").split(/[\\/]/).pop() ?? ""
  const clean = oneLine(base.replace(/\p{Cf}/gu, "").replace(/[\u0000-\u001f\u007f-\u009f]/g, " "))
    .replace(/^[.\s]+/, "")
    .replace(/[.\s]+$/, "")
  const typed = TYPE_EXTENSION[mimeType] ?? ""
  if (!clean) return `${MESSAGE_TYPES.has(mimeType) ? "message" : "attachment"}-${position}${typed}`
  const own = extension(clean)
  const ext = own && /^\.[a-z0-9]{1,10}$/.test(own) ? clean.slice(-own.length) : typed
  const stem = own && ext === clean.slice(-own.length) ? clean.slice(0, -own.length) : clean
  return `${truncateBytes(stem || "attachment", MAX_NAME_BYTES - Buffer.byteLength(ext))}${ext}`
}

export function extension(name: string) {
  const dot = name.lastIndexOf(".")
  return dot > 0 ? name.slice(dot).toLowerCase() : ""
}

function truncateBytes(text: string, limit: number) {
  let out = ""
  for (const character of text) {
    if (Buffer.byteLength(out + character) > limit) break
    out += character
  }
  return out.trimEnd()
}

export function oneLine(text: string) {
  return text.replace(/[\u0000-\u001f\u007f-\u009f\s]+/g, " ").trim()
}

export function formatSize(bytes: number) {
  if (bytes < 1024) return `${bytes} bytes`
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`
}

function formatAddress(addr: EmailAddress) {
  if (addr.name && addr.address) return `${addr.name} <${addr.address}>`
  return addr.address ?? addr.name ?? ""
}
