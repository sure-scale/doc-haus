import { burn, type Entry } from "@kenjiuno/msgreader/lib/Burner"
import { TypeEnum } from "@kenjiuno/msgreader/lib/Reader"

// Synthetic Outlook (.msg) emails for tests, written as OLE compound files in
// the MS-OXMSG layout with the compound-file writer of the reader ingest uses
// (@kenjiuno/msgreader). Every value comes from the test that builds it: no
// real email, person or organization. Nothing is read from outside the test.

// MAPI property tags (MS-OXPROPS): the id in the high word, the type in the low
// (001F a Unicode string, 001E an ANSI one, 0102 binary, 0003 an integer).
export const TAG = {
  messageClass: 0x001a001f,
  subject: 0x0037001f,
  clientSubmitTime: 0x00390040,
  sentRepresentingName: 0x0042001f,
  sentRepresentingAddressType: 0x0064001f,
  sentRepresentingEmail: 0x0065001f,
  transportHeaders: 0x007d001f,
  messageDeliveryTime: 0x0e060040,
  messageFlags: 0x0e070003,
  senderName: 0x0c1a001f,
  senderAddressType: 0x0c1e001f,
  senderEmail: 0x0c1f001f,
  senderSmtpAddress: 0x5d01001f,
  sentRepresentingSmtpAddress: 0x5d02001f,
  body: 0x1000001f,
  rtfCompressed: 0x10090102,
  html: 0x10130102,
  internetCodepage: 0x3fde0003,
  messageCodepage: 0x3ffd0003,
  recipientType: 0x0c150003,
  displayName: 0x3001001f,
  addressType: 0x3002001f,
  emailAddress: 0x3003001f,
  smtpAddress: 0x39fe001f,
  attachSize: 0x0e200003,
  attachMethod: 0x37050003,
  attachDataBinary: 0x37010102,
  attachLongFilename: 0x3707001f,
  attachMimeTag: 0x370e001f,
  attachContentId: 0x3712001f,
  attachFlags: 0x37140003,
  attachmentHidden: 0x7ffe000b,
} as const

export type PropValue = string | Uint8Array | number | boolean | Date
export type MsgNode = {
  props: [tag: number, value: PropValue][]
  recipients?: MsgNode[]
  attachments?: MsgAttachment[]
}
// An attachment's own properties, and for an attached message (method 5) or an
// OLE object (method 6) the storage that holds it.
export type MsgAttachment = MsgNode & { message?: MsgNode; oleStreams?: Record<string, Uint8Array> }

export type Address = { name?: string; address: string }

// A message with the properties an Outlook email has, from plain values.
export function outlookMessage(input: {
  subject?: string
  from?: Address
  to?: Address[]
  cc?: Address[]
  date?: Date
  body?: string
  html?: string
  rtf?: string
  attachments?: MsgAttachment[]
}): MsgNode {
  const props: MsgNode["props"] = [[TAG.messageClass, "IPM.Note"]]
  if (input.subject !== undefined) props.push([TAG.subject, input.subject])
  if (input.from) {
    props.push([TAG.senderAddressType, "SMTP"], [TAG.senderEmail, input.from.address], [TAG.senderSmtpAddress, input.from.address])
    if (input.from.name) props.push([TAG.senderName, input.from.name])
  }
  if (input.date) props.push([TAG.clientSubmitTime, input.date])
  if (input.body !== undefined) props.push([TAG.body, input.body])
  if (input.html !== undefined) props.push([TAG.html, new Uint8Array(Buffer.from(input.html, "utf8"))], [TAG.internetCodepage, 65001])
  if (input.rtf !== undefined) props.push([TAG.rtfCompressed, compressRtf(input.rtf)])
  const recipient = (type: number) => (address: Address): MsgNode => ({
    props: [
      [TAG.recipientType, type],
      [TAG.displayName, address.name ?? address.address],
      [TAG.addressType, "SMTP"],
      [TAG.emailAddress, address.address],
      [TAG.smtpAddress, address.address],
    ],
  })
  return {
    props,
    recipients: [...(input.to ?? []).map(recipient(1)), ...(input.cc ?? []).map(recipient(2))],
    attachments: input.attachments ?? [],
  }
}

// A file attached by value (method 1).
export function fileAttachment(name: string, type: string | undefined, bytes: Uint8Array): MsgAttachment {
  return {
    props: [
      [TAG.attachMethod, 1],
      [TAG.attachLongFilename, name],
      ...(type ? [[TAG.attachMimeTag, type] as [number, PropValue]] : []),
      [TAG.attachDataBinary, bytes],
    ],
  }
}

// An attached Outlook message (method 5), shown under its display name.
export function messageAttachment(displayName: string, message: MsgNode): MsgAttachment {
  return { props: [[TAG.attachMethod, 5], [TAG.displayName, displayName]], message }
}

export function msgFile(message: MsgNode) {
  const entries: Entry[] = [{ name: "Root Entry", type: TypeEnum.ROOT, length: 0, children: [] }]
  writeStorage(entries, 0, message, "message")
  return Buffer.from(burn(entries))
}

// The header of a storage's property stream: 32 bytes for the top-level
// message, 24 for an attached message, 8 for a recipient or an attachment.
function writeStorage(entries: Entry[], parent: number, node: MsgNode, kind: "message" | "attached" | "part") {
  const add = (entry: Entry) => {
    entries.push(entry)
    entries[parent]!.children!.push(entries.length - 1)
    return entries.length - 1
  }
  const stream = (name: string, data: Uint8Array) =>
    add({ name, type: TypeEnum.DOCUMENT, length: data.length, binaryProvider: () => data })
  const fixed: Uint8Array[] = []
  for (const [tag, value] of node.props) {
    const type = tag & 0xffff
    if (type === 0x001f) stream(`__substg1.0_${hex(tag)}`, new Uint8Array(Buffer.from(String(value), "utf16le")))
    // An ANSI string: the bytes the test encoded in the message's codepage.
    else if (type === 0x001e) stream(`__substg1.0_${hex(tag)}`, value instanceof Uint8Array ? value : new Uint8Array(Buffer.from(String(value), "latin1")))
    else if (type === 0x0102) stream(`__substg1.0_${hex(tag)}`, value as Uint8Array)
    else fixed.push(fixedProperty(tag, value))
  }
  const header = new Uint8Array(kind === "message" ? 32 : kind === "attached" ? 24 : 8)
  if (kind !== "part") {
    const view = new DataView(header.buffer)
    const recipients = node.recipients?.length ?? 0
    const attachments = node.attachments?.length ?? 0
    view.setUint32(8, recipients, true)
    view.setUint32(12, attachments, true)
    view.setUint32(16, recipients, true)
    view.setUint32(20, attachments, true)
  }
  stream("__properties_version1.0", Buffer.concat([header, ...fixed]))
  node.recipients?.forEach((recipient, index) => {
    const storage = add({ name: `__recip_version1.0_#${hex(index)}`, type: TypeEnum.DIRECTORY, length: 0, children: [] })
    writeStorage(entries, storage, recipient, "part")
  })
  node.attachments?.forEach((attachment, index) => {
    const storage = add({ name: `__attach_version1.0_#${hex(index)}`, type: TypeEnum.DIRECTORY, length: 0, children: [] })
    writeStorage(entries, storage, attachment, "part")
    if (!attachment.message && !attachment.oleStreams) return
    const object = entries.length
    entries.push({ name: "__substg1.0_3701000D", type: TypeEnum.DIRECTORY, length: 0, children: [] })
    entries[storage]!.children!.push(object)
    if (attachment.message) writeStorage(entries, object, attachment.message, "attached")
    for (const [name, data] of Object.entries(attachment.oleStreams ?? {})) {
      entries.push({ name, type: TypeEnum.DOCUMENT, length: data.length, binaryProvider: () => data })
      entries[object]!.children!.push(entries.length - 1)
    }
  })
}

function fixedProperty(tag: number, value: PropValue) {
  const out = new Uint8Array(16)
  const view = new DataView(out.buffer)
  view.setUint32(0, tag, true)
  // Readable and writable (MS-OXMSG 2.4.2.1).
  view.setUint32(4, 6, true)
  if (value instanceof Date) view.setBigUint64(8, BigInt(value.getTime()) * 10_000n + 116_444_736_000_000_000n, true)
  else view.setUint32(8, Number(value), true)
  return out
}

const hex = (value: number) => value.toString(16).toUpperCase().padStart(8, "0")

// Compressed RTF (MS-OXRTFCP, "LZFu"): a leading "{\rtf1\ansi" taken from the
// format's initial dictionary, every other byte as a literal, then the end
// marker (a reference to the current write position). The CRC is left 0; the
// reader does not check it.
export function compressRtf(rtf: string) {
  const raw = Buffer.from(rtf, "latin1")
  const prefix = "{\\rtf1\\ansi"
  const dictionarySize = 207
  const tokens: (number | { offset: number; length: number })[] = []
  let start = 0
  if (rtf.startsWith(prefix)) {
    tokens.push({ offset: 0, length: prefix.length })
    start = prefix.length
  }
  for (const byte of raw.subarray(start)) tokens.push(byte)
  tokens.push({ offset: (dictionarySize + raw.length) % 4096, length: 2 })
  const body: number[] = []
  for (let run = 0; run < tokens.length; run += 8) {
    const group = tokens.slice(run, run + 8)
    body.push(group.reduce<number>((flags, token, bit) => (typeof token === "number" ? flags : flags | (1 << bit)), 0))
    for (const token of group) {
      if (typeof token === "number") body.push(token)
      else {
        const reference = (token.offset << 4) | (token.length - 2)
        body.push(reference >> 8, reference & 0xff)
      }
    }
  }
  const header = Buffer.alloc(16)
  header.writeUInt32LE(body.length + 12, 0)
  header.writeUInt32LE(raw.length, 4)
  header.writeUInt32LE(0x75465a4c, 8)
  return new Uint8Array(Buffer.concat([header, Buffer.from(body)]))
}
