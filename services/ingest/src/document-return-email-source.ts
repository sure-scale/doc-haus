import { readFileSync } from "node:fs"
import { parseEmailFile } from "./ingest"
import { readDocumentTarget } from "./registered-document"
import { DocumentExchangeError, type DocumentExchangeSource } from "./document-exchanges"
import { sha256Hex } from "./document-versions"

/** Email metadata is useful provenance only when that registered message
 * actually carries these exact attachment bytes. A chosen filename or a From
 * string supplied by the browser is not evidence of receipt or ancestry. */
export async function readDocumentReturnEmailSource(directory: string, emailName: string, bytes: Uint8Array): Promise<DocumentExchangeSource> {
  if (!/\.(eml|msg)$/i.test(emailName) || /[\/\\\0]/.test(emailName)) throw new DocumentExchangeError("Choose an email saved in this matter.", "invalid_receipt_email", 400)
  const target = readDocumentTarget(directory, emailName)
  if (!target) throw new DocumentExchangeError("The receipt email is no longer available.", "receipt_email_missing", 404)
  const messageBytes = readFileSync(target.docPath), hash = sha256Hex(messageBytes)
  const parsed = await parseEmailFile(target.name, messageBytes)
  if (sha256Hex(readFileSync(target.docPath)) !== hash) throw new DocumentExchangeError("The email changed while its attachment was being checked. Inspect it again.", "receipt_email_changed")
  const attachmentHash = sha256Hex(bytes)
  const indexes = parsed.attachments.flatMap((attachment, index) => sha256Hex(attachment.content) === attachmentHash ? [index] : [])
  if (!indexes.length) throw new DocumentExchangeError("This file's exact bytes are not an attachment of that email. Choose its actual source, or leave the email source blank.", "receipt_attachment_mismatch", 400)
  const from = parsed.headers?.from
  const sender = from?.address ? (from.name ? `${from.name} <${from.address}>` : from.address) : from?.name
  return { documentName: target.name, attachmentId: `${hash}:${indexes.join(",")}`, ...(sender ? { sender } : {}), ...(parsed.headers?.date ? { date: parsed.headers.date } : {}) }
}
