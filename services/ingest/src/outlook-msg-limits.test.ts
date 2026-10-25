import { expect, test } from "bun:test"
import { MsgFormatError, parseMsg } from "../../../dochaus/lib/msg"
import { extractSourceText } from "./ingest"
import { compressRtf, fileAttachment, msgFile, outlookMessage, TAG, type MsgAttachment, type PropValue } from "./fixtures/synthetic-msg"

// Damaged or hostile Outlook files, built in the test from invented content
// (fixtures/synthetic-msg.ts): ingest parses uploads in its own process, so
// none may hold it or swell it, and none may read as an email with an empty
// body when its body could not be read.

const SENDER = { name: "Sender Example", address: "sender@example.test" }
const UNREAD = "The message body could not be read; only its headers were indexed"

const withRtf = (rtf: Uint8Array, attachments: MsgAttachment[] = []) => {
  const message = outlookMessage({ subject: "Synthetic RTF", from: SENDER, attachments })
  message.props.push([TAG.rtfCompressed, rtf])
  return msgFile(message)
}

const timed = async <T>(work: () => Promise<T>) => {
  const started = performance.now()
  const result = await work()
  return { result, ms: performance.now() - started }
}

test("a file whose directory entries all name one stream is refused, not read once per entry", async () => {
  // Many properties naming one stream: a 64 KB one (a 70 KB file that read as
  // 6.4 MB), or one under 4096 bytes, held in the mini stream.
  const aliased = (size: number) => {
    const message = outlookMessage({ subject: "Synthetic shared stream", from: SENDER, body: "Body." })
    message.props.push([0x60000102, new Uint8Array(size).fill(0x41)])
    for (let index = 0; index < 100; index++) message.props.push([((0x6100 + index) << 16) | 0x0102, new Uint8Array([1])])
    const file = msgFile(message)
    const entries = directory(file)
    const named = entries.find((entry) => entry.name === "__substg1.0_60000102")!
    for (const entry of entries.filter((entry) => /^__substg1\.0_61/.test(entry.name))) share(file, entry, named)
    return file
  }

  // Many attachments naming one attachment's content.
  const attachments = [fileAttachment("Big.pdf", "application/pdf", new Uint8Array(64 * 1024).fill(0x42))]
  for (let index = 0; index < 20; index++) attachments.push(fileAttachment(`Copy ${index}.pdf`, "application/pdf", new Uint8Array([index])))
  const shared = msgFile(outlookMessage({ subject: "Synthetic shared attachments", from: SENDER, body: "Body.", attachments }))
  const data = directory(shared).filter((entry) => entry.name === "__substg1.0_37010102")
  const content = data.find((entry) => entry.size === 64 * 1024)!
  for (const entry of data) if (entry !== content) share(shared, entry, content)

  for (const file of [aliased(64 * 1024), aliased(4000), shared]) await expect(parseMsg(file)).rejects.toBeInstanceOf(MsgFormatError)

  // A sound file with as many streams, large and small, still reads.
  const sound = Array.from({ length: 40 }, (_, index) =>
    fileAttachment(`Part ${index}.pdf`, "application/pdf", new Uint8Array(index % 2 ? 5000 : 100).fill(index)),
  )
  expect((await parseMsg(msgFile(outlookMessage({ subject: "Synthetic sound", from: SENDER, body: "Body.", attachments: sound })))).attachments).toHaveLength(40)
})

test("an RTF body nested deeper than any message is not read, and says so", async () => {
  const deep = 12_000
  const html = compressRtf(String.raw`{\rtf1\ansi\ansicpg1252\fromhtml1 {\*\htmltag <p>}` + `${"{".repeat(deep)}x${"}".repeat(deep)}}`)
  const native = compressRtf(String.raw`{\rtf1\ansi ` + `${"{".repeat(deep)}x${"}".repeat(deep)}}`)
  for (const rtf of [html, native]) {
    // Reading it took seconds, and minutes ten times deeper.
    const { result, ms } = await timed(() => parseMsg(withRtf(rtf)))
    expect(result).toMatchObject({ body: "", bodyUnreadable: true })
    expect(ms).toBeLessThan(2_000)
  }
  const quality = (await extractSourceText("Deep.msg", withRtf(html))).quality
  expect(quality).toEqual({ status: "degraded", signals: [{ code: "no_email_body", detail: UNREAD }] })

  // Nesting a real message could have still reads.
  const nested = compressRtf(String.raw`{\rtf1\ansi\ansicpg1252\fromhtml1 {\*\htmltag <p>}` + `${"{".repeat(100)}Nested text${"}".repeat(100)}}`)
  expect((await parseMsg(withRtf(nested))).body.trim()).toBe("Nested text")
})

test("an RTF body that would inflate past the limit is not read; a long one within it reads in full", async () => {
  // 1 MB of compressed RTF that inflates to 9 MB, past the 8 MB limit.
  const bomb = withRtf(rtfBomb(9 * 1024 * 1024))
  const before = process.memoryUsage().rss
  const { result, ms } = await timed(() => parseMsg(bomb))
  expect(result).toMatchObject({ body: "", bodyUnreadable: true })
  // Inflated into an array of numbers, it took over a gigabyte.
  expect(process.memoryUsage().rss - before).toBeLessThan(400 * 1024 * 1024)
  expect(ms).toBeLessThan(5_000)

  // One text run of a million characters, within the limit, reads in full
  // (a run this long once came back empty).
  const long = await parseMsg(withRtf(rtfBomb(1_000_000)))
  expect(long.bodyUnreadable).toBeUndefined()
  expect(long.body.length).toBeGreaterThanOrEqual(1_000_000)
  expect(long.body).toMatch(/^a+$/)
})

test("a property stored with a type it never has is left out, never a server error", async () => {
  const bytes = new Uint8Array([0x41, 0x42])
  const message = (props: [number, PropValue][], input: Parameters<typeof outlookMessage>[0] = { subject: "Synthetic", from: SENDER }) => {
    const node = outlookMessage(input)
    node.props.push(...props)
    return msgFile(node)
  }
  const attached = (props: [number, PropValue][]) => msgFile(outlookMessage({ subject: "Synthetic", from: SENDER, body: "Body.", attachments: [{ props }] }))

  // A body as binary or as a number, or compressed RTF as text: a body that
  // cannot be read.
  for (const file of [message([[0x10000102, bytes]]), message([[0x10000003, 7]]), message([[0x1009001f, "x".repeat(40)]])])
    expect(await parseMsg(file)).toMatchObject({ body: "", bodyUnreadable: true })
  // A sender name as binary: the address stands alone.
  expect((await parseMsg(message([[0x0c1a0102, bytes], [TAG.senderSmtpAddress, "sender@example.test"]], { subject: "Synthetic", body: "Body." }))).text).toContain(
    "From: sender@example.test\n",
  )
  // A file name as binary: named from its type and place; a MIME type as a
  // number: typed from its name.
  const unnamed = await parseMsg(attached([[TAG.attachMethod, 1], [0x37070102, bytes], [TAG.attachDataBinary, bytes]]))
  expect(unnamed.attachments.map(({ name, mimeType }) => [name, mimeType])).toEqual([["attachment-1", "application/octet-stream"]])
  const typed = await parseMsg(attached([[TAG.attachMethod, 1], [TAG.attachLongFilename, "Terms.pdf"], [0x370e0003, 5], [TAG.attachDataBinary, bytes]]))
  expect(typed.attachments.map(({ name, mimeType }) => [name, mimeType])).toEqual([["Terms.pdf", "application/pdf"]])
})

test("an RTF body that cannot be read leaves the email degraded, even with attachments", async () => {
  // A compressed RTF whose stated size runs past the property.
  const damaged = new Uint8Array(32)
  new DataView(damaged.buffer).setUint32(0, 1000, true)
  const note = fileAttachment("Note.pdf", "application/pdf", new Uint8Array(Buffer.from("%PDF-1.4 synthetic")))
  const file = withRtf(damaged, [note])
  expect(await parseMsg(file)).toMatchObject({ body: "", bodyUnreadable: true })
  expect((await extractSourceText("Damaged.msg", file)).quality).toEqual({ status: "degraded", signals: [{ code: "no_email_body", detail: UNREAD }] })
  // An email with no body that only carries attachments stays ok.
  const carrier = msgFile(outlookMessage({ subject: "Synthetic carrier", from: SENDER, attachments: [note] }))
  expect((await extractSourceText("Carrier.msg", carrier)).quality).toEqual({ status: "ok" })
})

// Compressed RTF that inflates to at least `size` bytes: a native RTF opening,
// then references that each repeat the byte before them 17 times, then the
// reference to the write position that ends the data.
function rtfBomb(size: number) {
  const literals = Buffer.from(String.raw`{\rtf1\ansi a`, "latin1")
  const references = Math.ceil(size / 17)
  const tokens = literals.length + references + 1
  const out = new Uint8Array(16 + Math.ceil(tokens / 8) + literals.length + 2 * (references + 1))
  let at = 16
  let control = 0
  let position = 207
  for (let token = 0; token < tokens; token++) {
    if (token % 8 === 0) control = at++
    if (token < literals.length) {
      out[at++] = literals[token]!
      position++
      continue
    }
    out[control] = out[control]! | (1 << (token % 8))
    const reference = token === tokens - 1 ? (position & 4095) << 4 : (((position - 1) & 4095) << 4) | 15
    out[at++] = reference >> 8
    out[at++] = reference & 0xff
    position += 17
  }
  const view = new DataView(out.buffer)
  view.setUint32(0, at - 4, true)
  view.setUint32(4, position - 207, true)
  view.setUint32(8, 0x75465a4c, true)
  return out.subarray(0, at)
}

// The compound file's directory (MS-CFB 2.6) as the generator lays it out:
// 512-byte sectors, the FAT in the sectors the header names, 128-byte entries.
function directory(file: Buffer) {
  const at = (sector: number) => (sector + 1) * 512
  const fat = Array.from({ length: Math.min(file.readInt32LE(0x2c), 109) }, (_, index) => file.readInt32LE(0x4c + 4 * index))
  const next = (sector: number) => file.readInt32LE(at(fat[Math.floor(sector / 128)]!) + 4 * (sector % 128))
  const entries: { offset: number; name: string; start: number; size: number }[] = []
  for (let sector = file.readInt32LE(0x30); sector !== -2; sector = next(sector))
    for (let offset = at(sector); offset < at(sector) + 512; offset += 128) {
      const length = file.readUInt16LE(offset + 0x40)
      entries.push({
        offset,
        name: file.toString("utf16le", offset, offset + Math.max(0, length - 2)),
        start: file.readInt32LE(offset + 0x74),
        size: file.readInt32LE(offset + 0x78),
      })
    }
  return entries
}

// Point a directory entry at another entry's stream.
function share(file: Buffer, entry: { offset: number }, target: { start: number; size: number }) {
  file.writeInt32LE(target.start, entry.offset + 0x74)
  file.writeInt32LE(target.size, entry.offset + 0x78)
}
