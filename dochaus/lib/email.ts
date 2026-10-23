import PostalMime, { type Attachment } from "postal-mime"

// Email sources (.eml) are MIME documents, not zip-based Office files, so they go
// through a real RFC822 parser. This is the one email extractor: ingest indexes
// its text and records its attachments, and liveText (extract.ts) re-extracts the
// same text to check citation offsets, so the text must depend on the email's
// bytes alone, never on what became of an attachment.
//
// The headers that carry legal meaning in a chain (who wrote, to whom, when,
// about what) sit above the body, followed by one line listing each attachment's
// name, type and size as the email states them. A forwarded message is kept as
// an attachment (forceRfc822Attachments) rather than flattened into the body, so
// it can become its own document with its own attachments, and so each parse
// reads one level of nesting only.

// The formats a matter document can take (ingest's MATTER_DOCUMENT_EXTENSIONS).
export const EMAIL_ATTACHMENT_EXTENSIONS = [".docx", ".pdf", ".xlsx", ".eml"] as const

const TYPE_EXTENSION: Record<string, string> = {
  "application/pdf": ".pdf",
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document": ".docx",
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": ".xlsx",
  "message/rfc822": ".eml",
}

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
}

export async function parseEmail(buffer: Buffer | Uint8Array): Promise<ParsedEmail> {
  const email = await PostalMime.parse(buffer, { forceRfc822Attachments: true })
  const parts = email.attachments.filter((part) => !(part.related && part.contentId))
  const attachments = parts.map((part, index) => emailAttachment(part, index + 1))
  const header = [
    email.subject && `Subject: ${email.subject}`,
    email.from && `From: ${formatAddress(email.from)}`,
    email.to?.length && `To: ${email.to.map(formatAddress).join(", ")}`,
    email.cc?.length && `Cc: ${email.cc.map(formatAddress).join(", ")}`,
    email.date && `Date: ${email.date}`,
    attachments.length &&
      `Attachments: ${attachments
        .map((attachment) => `${attachment.filename || "unnamed"} (${attachment.mimeType}, ${formatSize(attachment.size)})`)
        .join("; ")}`,
  ]
    .filter(Boolean)
    .join("\n")
  // An HTML-only body keeps the plain tag strip until a real converter lands (#103).
  const body = email.text ?? (email.html ? email.html.replace(/<[^>]+>/g, " ") : "")
  return {
    text: header ? `${header}\n\n${body}` : body,
    body,
    attachments,
    embedded: email.attachments.length - parts.length,
  }
}

export async function emlToText(buffer: Buffer | Uint8Array) {
  return (await parseEmail(buffer)).text
}

function emailAttachment(part: Attachment, position: number): EmailAttachment {
  const content =
    typeof part.content === "string"
      ? new TextEncoder().encode(part.content)
      : part.content instanceof Uint8Array
        ? part.content
        : new Uint8Array(part.content)
  const mimeType = (part.mimeType || "application/octet-stream").toLowerCase()
  const name = safeAttachmentName(part.filename, mimeType, position)
  return {
    filename: oneLine(part.filename ?? ""),
    name,
    mimeType,
    size: content.byteLength,
    content,
    supported: (EMAIL_ATTACHMENT_EXTENSIONS as readonly string[]).includes(extension(name)),
    tnef: mimeType === "application/ms-tnef" || /^winmail\.dat$/i.test(name),
  }
}

// The stated name is the sender's: take its last path segment, drop control
// characters and leading dots, keep its extension, add one from the type when it
// has none, and fall back to a name from the type and position when nothing is
// left. Never refused, so no attachment is lost to its name.
export function safeAttachmentName(filename: string | null | undefined, mimeType: string, position: number) {
  const base = (filename ?? "").split(/[\\/]/).pop() ?? ""
  const clean = oneLine(base.replace(/[\u0000-\u001f\u007f-\u009f]/g, " "))
    .replace(/^[.\s]+/, "")
    .replace(/[.\s]+$/, "")
  const typed = TYPE_EXTENSION[mimeType] ?? ""
  if (!clean) return `${mimeType === "message/rfc822" ? "message" : "attachment"}-${position}${typed}`
  const own = extension(clean)
  const ext = own && /^\.[a-z0-9]{1,10}$/.test(own) ? clean.slice(-own.length) : typed
  const stem = own && ext === clean.slice(-own.length) ? clean.slice(0, -own.length) : clean
  return `${truncateBytes(stem || "attachment", MAX_NAME_BYTES - Buffer.byteLength(ext))}${ext}`
}

function extension(name: string) {
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

function oneLine(text: string) {
  return text.replace(/[\u0000-\u001f\u007f-\u009f\s]+/g, " ").trim()
}

function formatSize(bytes: number) {
  if (bytes < 1024) return `${bytes} bytes`
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`
}

function formatAddress(addr: { name?: string; address?: string }) {
  if (addr.name && addr.address) return `${addr.name} <${addr.address}>`
  return addr.address ?? addr.name ?? ""
}
