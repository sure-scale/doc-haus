import { afterAll, expect, spyOn, test } from "bun:test"
import { mkdtempSync, readFileSync, realpathSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { Document, Packer, Paragraph } from "docx"
import iconv from "iconv-lite"
import { findQuote, liveText } from "../../../dochaus/lib/extract"
import * as msgReader from "../../../dochaus/lib/msg"
import { verifyQuote } from "../../../dochaus/lib/quote-verify"
import { openDb } from "./db"
import { deleteMatterDocument } from "./document-delete"
import { emailRemovalPlan, ingestEmailDocument, MAX_EMAIL_DEPTH, reingestLegacyEmails, withEmailLinks } from "./email-attachments"
import { extractDocumentText, extractSourceText } from "./ingest"
import { sourceArtifactsMatchRegistry } from "./source-artifacts"
import { fileAttachment, messageAttachment, msgFile, outlookMessage, TAG, type MsgAttachment, type MsgNode } from "./fixtures/synthetic-msg"

const { MsgFormatError, parseMsg } = msgReader

// Every email, name and attachment here is invented test content built in the
// test (fixtures/synthetic-msg.ts writes the .msg files): example.test
// addresses, no actual person, organization or transaction.

const directories: string[] = []
afterAll(() => directories.forEach((directory) => rmSync(directory, { recursive: true, force: true })))

function matter() {
  const directory = realpathSync(mkdtempSync(path.join(tmpdir(), "dochaus-outlook-msg-")))
  directories.push(directory)
  return directory
}

const DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
const SENDER = { name: "Sender Example", address: "sender@example.test" }
const RECIPIENT = { name: "Recipient Example", address: "recipient@example.test" }
const COPY = { address: "copy@example.test" }
const SENT = new Date("2024-03-04T09:30:00Z")
// The size listed for an attached Outlook message that states none: the length
// of the reader's own written-out copy. It is part of the parent email's text,
// so a reader upgrade that changes it moves the parent's citation offsets.
const REWRITTEN_MESSAGE_SIZE = 3072

const docx = async (text: string) =>
  new Uint8Array(await Packer.toBuffer(new Document({ sections: [{ children: [new Paragraph(text)] }] })))

// A one-page PDF with one line of text.
function pdf(wording: string) {
  const stream = `BT /F1 12 Tf 40 700 Td (${wording}) Tj ET`
  const objects = [
    "<< /Type /Catalog /Pages 2 0 R >>",
    "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
    "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>",
    `<< /Length ${stream.length} >>\nstream\n${stream}\nendstream`,
    "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
  ]
  let text = "%PDF-1.4\n"
  const offsets = objects.map((object, index) => {
    const offset = text.length
    text += `${index + 1} 0 obj\n${object}\nendobj\n`
    return offset
  })
  const xref = text.length
  text += `xref\n0 6\n0000000000 65535 f \n${offsets.map((offset) => `${String(offset).padStart(10, "0")} 00000 n \n`).join("")}trailer\n<< /Size 6 /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`
  return new Uint8Array(Buffer.from(text))
}

function registry(directory: string) {
  const db = openDb(directory)
  try {
    return db.query("SELECT name, parent_name, attachments FROM documents WHERE role = 'source' ORDER BY created_at, id").all() as {
      name: string
      parent_name: string | null
      attachments: string | null
    }[]
  } finally {
    db.close()
  }
}

// Each stored chunk reads back from the live text at its offsets, which is what
// cite and the citation verifier check.
async function expectCitable(directory: string, name: string) {
  const db = openDb(directory)
  const chunks = db.query("SELECT text, char_start, char_end FROM chunks WHERE doc_name = ?").all(name) as {
    text: string
    char_start: number
    char_end: number
  }[]
  db.close()
  expect(chunks.length).toBeGreaterThan(0)
  const text = await liveText(path.join(directory, name))
  for (const chunk of chunks) expect(text.slice(chunk.char_start, chunk.char_end)).toBe(chunk.text)
}

test("an Outlook email reads like the same .eml, and its .docx attachment becomes a document that names it", async () => {
  const schedule = await docx("Synthetic delivery schedule marker")
  const body = "Please find the delivery schedule attached.\r\n\r\nRegards,\r\nSender Example\r\n"
  const msg = msgFile(outlookMessage({
    subject: "Synthetic delivery schedule",
    from: SENDER,
    to: [RECIPIENT],
    cc: [COPY],
    date: SENT,
    body,
    attachments: [fileAttachment("Schedule.docx", DOCX, schedule)],
  }))
  // The same message as MIME.
  const eml = Buffer.from([
    `From: ${SENDER.name} <${SENDER.address}>`,
    `To: ${RECIPIENT.name} <${RECIPIENT.address}>`,
    `Cc: ${COPY.address}`,
    "Subject: Synthetic delivery schedule",
    "Date: Mon, 04 Mar 2024 09:30:00 +0000",
    "MIME-Version: 1.0",
    'Content-Type: multipart/mixed; boundary="B"',
    "",
    "--B",
    "Content-Type: text/plain; charset=utf-8",
    "",
    body.replace(/\r\n$/, ""),
    "--B",
    `Content-Type: ${DOCX}; name="Schedule.docx"`,
    'Content-Disposition: attachment; filename="Schedule.docx"',
    "Content-Transfer-Encoding: base64",
    "",
    Buffer.from(schedule).toString("base64"),
    "--B--",
    "",
  ].join("\r\n"))

  // The same header block and body text, as indexed.
  const text = await extractDocumentText("Thread.msg", msg)
  expect(text).toBe(await extractDocumentText("Thread.eml", eml))
  expect(text).toBe([
    "Subject: Synthetic delivery schedule",
    "From: Sender Example <sender@example.test>",
    "To: Recipient Example <recipient@example.test>",
    "Cc: copy@example.test",
    "Date: 2024-03-04T09:30:00.000Z",
    `Attachments: Schedule.docx (${DOCX}, ${(schedule.length / 1024).toFixed(1)} KB)`,
    "",
    "Please find the delivery schedule attached.",
    "",
    "Regards,",
    "Sender Example",
    "",
  ].join("\n"))

  const directory = matter()
  const result = await ingestEmailDocument(directory, "Thread.msg", msg)
  expect(result.quality).toEqual({ status: "ok" })
  expect(result.attachments).toEqual([{ name: "Schedule.docx", type: DOCX, size: schedule.length, document: "Schedule.docx", status: "indexed" }])
  expect(registry(directory).map((row) => [row.name, row.parent_name])).toEqual([
    ["Schedule.docx", "Thread.msg"],
    ["Thread.msg", null],
  ])
  // Searchable and citable like an upload: stored offsets verify against the
  // live text, cite anchors in it and verify-quote finds it in the index.
  for (const name of ["Thread.msg", "Schedule.docx"]) await expectCitable(directory, name)
  const live = await liveText(path.join(directory, "Thread.msg"))
  expect(findQuote(live, "Please find the delivery schedule attached.")).toBeDefined()
  expect(findQuote(live, "From: Sender Example")).toBeDefined()
  expect(verifyQuote({ directory, quote: "Please find the delivery schedule attached.", document: "Thread.msg" }).found).toBe(true)
  expect(verifyQuote({ directory, quote: "Synthetic delivery schedule marker" }).hits.map((hit) => hit.doc)).toEqual(["Schedule.docx"])
  expect(sourceArtifactsMatchRegistry(directory)).toBe(true)
  const db = openDb(directory)
  const listed = withEmailLinks(db, db.query("SELECT name FROM documents ORDER BY created_at, id").all() as { name: string }[])
  db.close()
  expect(listed.find((row) => row.name === "Thread.msg")).toMatchObject({
    attachments: [{ status: "indexed", document: "Schedule.docx" }],
    removal: { attachments: ["Schedule.docx"], kept: [] },
  })
}, 120_000)

test("an RTF-only body reads as text, whether native RTF or HTML or text wrapped in RTF", async () => {
  const native = String.raw`{\rtf1\ansi\ansicpg1252\deff0{\fonttbl{\f0\fswiss Hidden Font Name;}}{\colortbl;\red0\green0\blue0;}{\*\generator Hidden generator;}{\info{\author Hidden author}}\uc1\pard\f0\fs20 Dear Recipient,\par\par The caf\'e9 terms run for \b five\b0  years\u8217?s length.\tab End.\par {\field{\*\fldinst HYPERLINK "https://example.test/hidden"}{\fldrslt the portal}}\line Sender\par}`
  const wrappedHtml = String.raw`{\rtf1\ansi\ansicpg1252\fromhtml1 \deff0{\fonttbl{\f0\fswiss Hidden Font Name;}}{\*\htmltag19 <html>}{\*\htmltag50 <body>}\htmlrtf {\htmlrtf0 {\*\htmltag64 <p>}Wrapped HTML caf\'e9 body{\*\htmltag72 </p>}\htmlrtf \par\htmlrtf0}{\*\htmltag58 </body>}{\*\htmltag27 </html>}}`
  const wrappedText = String.raw`{\rtf1\ansi\ansicpg1252\fromtext \deff0{\fonttbl{\f0\fswiss Hidden Font Name;}}\uc1\pard Wrapped plain body\par second line\par}`
  const read = async (rtf: string) => (await parseMsg(msgFile(outlookMessage({ subject: "Synthetic RTF", from: SENDER, rtf })))).body

  const nativeBody = await read(native)
  expect(nativeBody).toBe("Dear Recipient,\n\nThe café terms run for five years’s length.\tEnd.\nthe portal\nSender\n")
  expect(nativeBody).not.toMatch(/Hidden/)
  expect((await read(wrappedHtml)).replace(/\s+/g, " ").trim()).toBe("Wrapped HTML café body")
  expect(await read(wrappedText)).toBe("Wrapped plain body\r\nsecond line\r\n")

  // Indexed as readable text, and citable.
  const directory = matter()
  const msg = msgFile(outlookMessage({ subject: "Synthetic RTF", from: SENDER, date: SENT, rtf: native }))
  expect((await extractSourceText("Rich.msg", msg)).quality).toEqual({ status: "ok" })
  await ingestEmailDocument(directory, "Rich.msg", msg)
  await expectCitable(directory, "Rich.msg")
  expect(verifyQuote({ directory, quote: "The café terms run for five years", document: "Rich.msg" }).found).toBe(true)
}, 120_000)

test("an attached Outlook message becomes a child email with its own attachments, down to the depth limit", async () => {
  const directory = matter()
  const inner = outlookMessage({
    subject: "Synthetic forwarded",
    from: SENDER,
    to: [RECIPIENT],
    body: "Forwarded body.\r\n",
    attachments: [fileAttachment("Inner.pdf", "application/pdf", pdf("Synthetic inner marker"))],
  })
  const outer = msgFile(outlookMessage({ subject: "Synthetic outer", from: SENDER, body: "See forwarded.\r\n", attachments: [messageAttachment("Synthetic forwarded", inner)] }))

  const result = await ingestEmailDocument(directory, "Outer.msg", outer)

  expect(result.attachments).toEqual([
    expect.objectContaining({ name: "Synthetic forwarded", type: "application/vnd.ms-outlook", status: "indexed", document: "Synthetic forwarded.msg" }),
  ])
  const rows = registry(directory)
  expect(rows.find((row) => row.name === "Synthetic forwarded.msg")?.parent_name).toBe("Outer.msg")
  expect(rows.find((row) => row.name === "Inner.pdf")?.parent_name).toBe("Synthetic forwarded.msg")
  // The forwarded text is the child's, not the outer email's.
  expect(await liveText(path.join(directory, "Outer.msg"))).not.toContain("Forwarded body.")
  expect(await liveText(path.join(directory, "Synthetic forwarded.msg"))).toContain("Forwarded body.")
  await expectCitable(directory, "Synthetic forwarded.msg")

  // Nested past the limit: the deepest message is listed on its parent, not ingested.
  const deep = matter()
  let nested = outlookMessage({ subject: `Level ${MAX_EMAIL_DEPTH + 1}`, body: "Deepest." })
  for (let level = MAX_EMAIL_DEPTH; level >= 1; level--)
    nested = outlookMessage({ subject: `Level ${level}`, body: `Level ${level} body.`, attachments: [messageAttachment(`level-${level + 1}`, nested)] })
  await ingestEmailDocument(deep, "level-0.msg", msgFile(outlookMessage({ subject: "Level 0", body: "Level 0 body.", attachments: [messageAttachment("level-1", nested)] })))
  const deepRows = registry(deep)
  expect(deepRows).toHaveLength(MAX_EMAIL_DEPTH + 1)
  const last = deepRows.find((row) => row.name === `level-${MAX_EMAIL_DEPTH}.msg`)!
  expect(JSON.parse(last.attachments!).attachments).toEqual([
    expect.objectContaining({ name: `level-${MAX_EMAIL_DEPTH + 1}`, outcome: "nested" }),
  ])

  // An .eml carrying an Outlook message brings it in as a child email too.
  const mixed = matter()
  const attached = msgFile(inner)
  const eml = Buffer.from([
    "From: sender@example.test",
    "Subject: Synthetic carrier",
    "MIME-Version: 1.0",
    'Content-Type: multipart/mixed; boundary="B"',
    "",
    "--B",
    "Content-Type: text/plain",
    "",
    "Outlook message attached.",
    "--B",
    'Content-Type: application/vnd.ms-outlook; name="Forwarded.msg"',
    'Content-Disposition: attachment; filename="Forwarded.msg"',
    "Content-Transfer-Encoding: base64",
    "",
    attached.toString("base64"),
    "--B--",
    "",
  ].join("\r\n"))
  const carried = await ingestEmailDocument(mixed, "Carrier.eml", eml)
  expect(carried.attachments[0]).toMatchObject({ name: "Forwarded.msg", status: "indexed", document: "Forwarded.msg" })
  expect(registry(mixed).find((row) => row.name === "Inner.pdf")?.parent_name).toBe("Forwarded.msg")
}, 180_000)

test("an OLE object and a linked file are listed as not supported, and an inline image is counted", async () => {
  const ole: MsgAttachment = {
    // An OLE object whose display name looks like a document is still not read.
    props: [[TAG.attachMethod, 6], [TAG.displayName, "Embedded chart.docx"]],
    oleStreams: { CONTENTS: new Uint8Array(Buffer.from("synthetic ole object")) },
  }
  const linked: MsgAttachment = { props: [[TAG.attachMethod, 2], [TAG.attachLongFilename, "Shared.pdf"], [TAG.attachMimeTag, "application/pdf"]] }
  const inline = fileAttachment("logo.png", "image/png", new Uint8Array(Buffer.from("synthetic png")))
  inline.props.push([TAG.attachContentId, "logo@example.test"], [TAG.attachmentHidden, true])
  const msg = msgFile(outlookMessage({
    subject: "Synthetic mixed",
    from: SENDER,
    html: '<p>Logo below</p><img src="cid:logo@example.test">',
    attachments: [ole, linked, inline],
  }))

  const parsed = await parseMsg(msg)
  expect(parsed.embedded).toBe(1)
  expect(parsed.body.replace(/\s+/g, " ").trim()).toBe("Logo below")
  const directory = matter()
  const result = await ingestEmailDocument(directory, "Mixed.msg", msg)
  expect(result.attachments.map((entry) => [entry.name, entry.status, entry.note])).toEqual([
    ["Embedded chart.docx", "unsupported", "ole"],
    ["Shared.pdf", "unsupported", "linked"],
  ])
  expect(result.injection).toBeNull()
  expect(registry(directory).map((row) => row.name)).toEqual(["Mixed.msg"])
  // Their names are a document's, but they are noted as never read, so a
  // restart neither takes the email in again nor parses it to find out.
  const parses = spyOn(msgReader, "parseMsg")
  try {
    expect(await reingestLegacyEmails(directory)).toBe(0)
    expect(parses).not.toHaveBeenCalled()
  } finally {
    parses.mockRestore()
  }
}, 120_000)

test("a malformed .msg is refused with a clear message, and never indexed as an empty success", async () => {
  const good = msgFile(outlookMessage({ subject: "Synthetic", from: SENDER, body: "Body." }))
  const notOutlook = Buffer.from("This is not an Outlook email.")
  const cut = good.subarray(0, 1500)
  for (const bytes of [notOutlook, cut]) await expect(parseMsg(bytes)).rejects.toBeInstanceOf(MsgFormatError)
  // A compound file with no Outlook item in it.
  await expect(parseMsg(msgFile({ props: [] }))).rejects.toThrow("not a readable Outlook email")
  // A damaged file whose directory sectors or entries loop back on themselves
  // is refused, not followed forever.
  for (const bytes of [loopedDirectorySectors(good), loopedDirectoryEntries(good)])
    await expect(parseMsg(bytes)).rejects.toBeInstanceOf(MsgFormatError)

  const directory = matter()
  await expect(ingestEmailDocument(directory, "Broken.msg", cut)).rejects.toBeInstanceOf(MsgFormatError)
  expect(registry(directory)).toEqual([])
  // An Outlook item with nothing in it is kept with nothing indexed.
  expect((await extractSourceText("Empty.msg", msgFile(outlookMessage({})))).quality.status).toBe("no_readable_text")
  // Attached to an email, a malformed .msg leaves the email ingested and is marked failed.
  const carrier = msgFile(outlookMessage({ subject: "Synthetic carrier", from: SENDER, body: "See attached.", attachments: [fileAttachment("Broken.msg", "application/vnd.ms-outlook", cut)] }))
  const result = await ingestEmailDocument(directory, "Carrier.msg", carrier)
  expect(result.attachments).toEqual([expect.objectContaining({ name: "Broken.msg", status: "failed" })])
  expect(registry(directory).map((row) => row.name)).toEqual(["Carrier.msg"])
  const index = JSON.parse(readFileSync(path.join(directory, "source-index.json"), "utf8")) as { sources: { document: string }[] }
  expect(index.sources.map((source) => source.document)).toEqual(["Carrier.msg"])
}, 120_000)

test("an ANSI Outlook email reads in its own codepage, an RTF body in each font's, an HTML body in the one it states", async () => {
  const encoded = (text: string, codepage: string) => new Uint8Array(iconv.encode(text, codepage))
  // Strings stored as ANSI (type 001E) in the message's codepage.
  const western = outlookMessage({})
  western.props.push(
    [0x0037001e, encoded("Fee €5m “final”", "cp1252")],
    [0x0c1a001e, encoded("Renée Example", "cp1252")],
    [TAG.senderSmtpAddress, "renee@example.test"],
    [0x1000001e, encoded("The fee is €5m – “final”; it’s agreed.\r\n", "cp1252")],
    [TAG.messageCodepage, 1252],
  )
  expect((await parseMsg(msgFile(western))).text).toBe(
    "Subject: Fee €5m “final”\nFrom: Renée Example <renee@example.test>\n\nThe fee is €5m – “final”; it’s agreed.\r\n",
  )
  const central = outlookMessage({})
  central.props.push([0x0037001e, encoded("Zażółć gęślą jaźń", "cp1250")], [0x1000001e, encoded("Łódź – umowa", "cp1250")], [TAG.messageCodepage, 1250])
  expect((await parseMsg(msgFile(central))).text).toBe("Subject: Zażółć gęślą jaźń\n\nŁódź – umowa")
  // With no message codepage, the internet codepage stands in only when it is
  // a Windows ANSI one; any other (UTF-8 here) leaves Windows-1252.
  const internet = (text: string, codepage: string, internetCodepage: number) => {
    const node = outlookMessage({})
    node.props.push([0x0037001e, encoded(text, codepage)], [TAG.internetCodepage, internetCodepage])
    return msgFile(node)
  }
  expect((await parseMsg(internet("Renée’s fee", "cp1252", 65001))).text).toBe("Subject: Renée’s fee\n\n")
  expect((await parseMsg(internet("Łódź – umowa", "cp1250", 1250))).text).toBe("Subject: Łódź – umowa\n\n")

  // An RTF font's character set decides its bytes, else the document's \ansicpg.
  const rtf = String.raw`{\rtf1\ansi\ansicpg1252\deff0{\fonttbl{\f0 Arial;}{\f1\fcharset238 Arial CE;}}\pard\f0 Price \'80 5m\par\f1 Miasto \'b3\'f3d\'9f\par\plain Total \'80\par}`
  expect((await parseMsg(msgFile(outlookMessage({ subject: "Synthetic RTF", rtf })))).body).toBe("Price € 5m\nMiasto łódź\nTotal €\n")

  // HTML in a codepage iconv-lite lacks (ISO-2022-JP) or names otherwise (UTF-16).
  const html = (bytes: Uint8Array, codepage: number) => {
    const node = outlookMessage({ subject: "Synthetic HTML" })
    node.props.push([TAG.html, bytes], [TAG.internetCodepage, codepage])
    return msgFile(node)
  }
  const japanese = new Uint8Array([0x3c, 0x70, 0x3e, 0x1b, 0x24, 0x42, 0x46, 0x7c, 0x4b, 0x5c, 0x1b, 0x28, 0x42, 0x3c, 0x2f, 0x70, 0x3e])
  expect((await parseMsg(html(japanese, 50220))).body.trim()).toBe("日本")
  expect((await parseMsg(html(new Uint8Array(Buffer.from("<p>Fee €5m</p>", "utf16le")), 1200))).body.trim()).toBe("Fee €5m")
})

test("From, To and Date read as an .eml's: the person a delegate sent for, Exchange addresses, no date on a draft", async () => {
  // Sent by an assistant on a partner's behalf: From is the partner.
  const delegated = outlookMessage({ subject: "Synthetic delegated", body: "Body." })
  delegated.props.push(
    [TAG.senderName, "Assistant Example"],
    [TAG.senderAddressType, "EX"],
    [TAG.senderEmail, "/O=EXAMPLE/OU=EXCHANGE ADMINISTRATIVE GROUP/CN=RECIPIENTS/CN=ASSISTANT"],
    [TAG.sentRepresentingName, "Partner Example"],
    [TAG.sentRepresentingAddressType, "SMTP"],
    [TAG.sentRepresentingEmail, "partner@example.test"],
    [TAG.sentRepresentingSmtpAddress, "partner@example.test"],
  )
  expect((await parseMsg(msgFile(delegated))).text).toStartWith("Subject: Synthetic delegated\nFrom: Partner Example <partner@example.test>\n")

  // Exchange addresses are directory paths: the transport headers give the
  // address for the same name, else the name stands alone.
  const exchange: MsgNode = outlookMessage({ subject: "Synthetic exchange", body: "Body." })
  exchange.props.push(
    [TAG.sentRepresentingName, "Partner Example"],
    [TAG.sentRepresentingAddressType, "EX"],
    [TAG.sentRepresentingEmail, "/O=EXAMPLE/CN=PARTNER"],
    [TAG.transportHeaders, "From: Partner Example <partner@example.test>\r\nTo: Client Example <client@example.test>\r\nSubject: Synthetic exchange\r\n"],
  )
  const exchangeRecipient = (type: number, name: string): MsgNode => ({
    props: [[TAG.recipientType, type], [TAG.displayName, name], [TAG.addressType, "EX"], [TAG.emailAddress, `/O=EXAMPLE/CN=${name.toUpperCase()}`]],
  })
  exchange.recipients = [exchangeRecipient(1, "Client Example"), exchangeRecipient(2, "Colleague Example")]
  expect((await parseMsg(msgFile(exchange))).text).toStartWith(
    "Subject: Synthetic exchange\nFrom: Partner Example <partner@example.test>\nTo: Client Example <client@example.test>\nCc: Colleague Example\n\n",
  )
  // Two people of one name in the headers: neither address stands in for it.
  const namesakes = outlookMessage({ subject: "Synthetic namesakes", body: "Body." })
  namesakes.props.push([TAG.transportHeaders, "From: sender@example.test\r\nTo: Sam Example <sam@one.example.test>\r\nCc: Sam Example <sam@two.example.test>\r\n"])
  namesakes.recipients = [exchangeRecipient(1, "Sam Example"), exchangeRecipient(2, "Sam Example")]
  expect((await parseMsg(msgFile(namesakes))).text).toStartWith("Subject: Synthetic namesakes\nTo: Sam Example\nCc: Sam Example\n\n")

  // Exchange mail from its own sender states the sender twice, its SMTP
  // address once: From has the address. A delegate's is never borrowed.
  const internal = (name: string, directory: string) => {
    const node = outlookMessage({ subject: "Synthetic internal", body: "Body." })
    node.props.push(
      [TAG.senderName, "Pat Example"],
      [TAG.senderAddressType, "EX"],
      [TAG.senderEmail, "/O=EXAMPLE/CN=PAT"],
      [TAG.senderSmtpAddress, "pat@example.test"],
      [TAG.sentRepresentingName, name],
      [TAG.sentRepresentingAddressType, "EX"],
      [TAG.sentRepresentingEmail, directory],
    )
    return msgFile(node)
  }
  expect((await parseMsg(internal("Pat Example", "/O=EXAMPLE/CN=PAT"))).text).toContain("\nFrom: Pat Example <pat@example.test>\n")
  expect((await parseMsg(internal("Example, Pat", "/o=example/cn=pat"))).text).toContain("\nFrom: Example, Pat <pat@example.test>\n")
  expect((await parseMsg(internal("Partner Example", "/O=EXAMPLE/CN=PARTNER"))).text).toContain("\nFrom: Partner Example\n")

  // A draft never sent has no Date unless it states when it was sent; a
  // received message dates from delivery.
  const saved = (flags: number, submitted?: Date) => {
    const node = outlookMessage({ subject: "Synthetic draft", body: "Draft.", date: submitted })
    node.props.push([TAG.messageFlags, flags], [TAG.messageDeliveryTime, SENT])
    return msgFile(node)
  }
  expect((await parseMsg(saved(8))).text).toBe("Subject: Synthetic draft\n\nDraft.")
  expect((await parseMsg(saved(8, new Date("2024-03-01T08:00:00Z")))).text).toBe("Subject: Synthetic draft\nDate: 2024-03-01T08:00:00.000Z\n\nDraft.")
  expect((await parseMsg(saved(1))).text).toBe("Subject: Synthetic draft\nDate: 2024-03-04T09:30:00.000Z\n\nDraft.")
})

test("attachments are named and sized as the message states them, inline images are counted however marked, an unreadable one is failed", async () => {
  const reply = messageAttachment("RE: 12/03 call", outlookMessage({ subject: "RE: 12/03 call", body: "Reply." }))
  const sized = messageAttachment("Sized reply", outlookMessage({ subject: "Sized reply", body: "Sized." }))
  sized.props.push([TAG.attachSize, 4321])
  const image = (name: string, type: string | undefined, id: string, ...props: [number, number][]) => {
    const part = fileAttachment(name, type, new Uint8Array(Buffer.from(`synthetic ${name}`)))
    part.props.push([TAG.attachContentId, id], ...props)
    return part
  }
  // Shown by the HTML, as its flags say or as the HTML names it; one it does
  // not show is an attachment like any other.
  const banner = image("banner.png", "image/png", "banner@example.test", [TAG.attachFlags, 4])
  const signature = image("signature.gif", undefined, "signature@example.test")
  const photo = image("photo.jpg", "image/jpeg", "photo@example.test")
  // Held by value, but with no content to read.
  const missing: MsgAttachment = { props: [[TAG.attachMethod, 1], [TAG.attachLongFilename, "Missing.docx"], [TAG.attachMimeTag, DOCX]] }
  const msg = msgFile(outlookMessage({
    subject: "Synthetic attachments",
    from: SENDER,
    html: '<p>See below.</p><img src="cid:signature@example.test">',
    attachments: [reply, sized, banner, signature, photo, missing],
  }))

  const parsed = await parseMsg(msg)
  expect(parsed.embedded).toBe(2)
  expect(parsed.attachments.map(({ name, size }) => [name, size])).toEqual([
    // A slash in a subject is not a folder.
    ["RE: 12-03 call.msg", REWRITTEN_MESSAGE_SIZE],
    ["Sized reply.msg", 4321],
    ["photo.jpg", 19],
    ["Missing.docx", 0],
  ])
  const directory = matter()
  const result = await ingestEmailDocument(directory, "Attachments.msg", msg)
  expect(result.attachments.map((entry) => [entry.name, entry.status, entry.document])).toEqual([
    ["RE: 12/03 call", "indexed", "RE: 12-03 call.msg"],
    ["Sized reply", "indexed", "Sized reply.msg"],
    ["photo.jpg", "unsupported", undefined],
    ["Missing.docx", "failed", undefined],
  ])
  // Listed as failed, not as a format it can now read: a restart leaves it.
  expect(await reingestLegacyEmails(directory)).toBe(0)
}, 120_000)

test("each email is parsed once, and a restart that takes one in again keeps its parent", async () => {
  // An Outlook email with an attached Outlook message: one parse each.
  const parses = spyOn(msgReader, "parseMsg")
  try {
    const inner = outlookMessage({ subject: "Synthetic inner", from: SENDER, body: "Inner body." })
    await ingestEmailDocument(matter(), "Outer.msg", msgFile(outlookMessage({ subject: "Synthetic outer", from: SENDER, body: "Outer body.", attachments: [messageAttachment("Synthetic inner", inner)] })))
    expect(parses).toHaveBeenCalledTimes(2)
  } finally {
    parses.mockRestore()
  }

  // An .eml forwarded inside an .eml carries an Outlook message, recorded by a
  // build that did not read .msg: a restart takes the inner .eml in again.
  const directory = matter()
  const mime = (lines: string[]) => Buffer.from(lines.join("\r\n"))
  const outlook = msgFile(outlookMessage({ subject: "Synthetic outlook", from: SENDER, body: "Outlook body." }))
  const child = mime([
    "From: child@example.test", "Subject: Synthetic child", "MIME-Version: 1.0", 'Content-Type: multipart/mixed; boundary="C"', "",
    "--C", "Content-Type: text/plain", "", "Child body.",
    "--C", 'Content-Type: application/vnd.ms-outlook; name="Outlook.msg"', 'Content-Disposition: attachment; filename="Outlook.msg"',
    "Content-Transfer-Encoding: base64", "", outlook.toString("base64"), "--C--", "",
  ])
  const outer = mime([
    "From: outer@example.test", "Subject: Synthetic outer", "MIME-Version: 1.0", 'Content-Type: multipart/mixed; boundary="O"', "",
    "--O", "Content-Type: text/plain", "", "Outer body.",
    "--O", "Content-Type: message/rfc822", 'Content-Disposition: attachment; filename="Child.eml"', "", child.toString("latin1"), "--O--", "",
  ])
  await ingestEmailDocument(directory, "Outer.eml", outer)
  await deleteMatterDocument(directory, "Outlook.msg")
  const db = openDb(directory)
  const childRecord = JSON.parse(registry(directory).find((row) => row.name === "Child.eml")!.attachments!)
  childRecord.attachments = childRecord.attachments.map(({ name, type, size }: { name: string; type: string; size: number }) => ({ name, type, size, outcome: "unsupported" }))
  db.query("UPDATE documents SET attachments = ? WHERE name = 'Child.eml'").run(JSON.stringify(childRecord))
  db.close()

  expect(await reingestLegacyEmails(directory)).toBe(1)
  const parents = Object.fromEntries(registry(directory).map((row) => [row.name, row.parent_name]))
  expect(parents).toEqual({ "Outer.eml": null, "Child.eml": "Outer.eml", "Outlook.msg": "Child.eml" })
  const check = openDb(directory)
  try {
    expect(emailRemovalPlan(check, "Outer.eml").attachments.sort()).toEqual(["Child.eml", "Outlook.msg"])
  } finally {
    check.close()
  }
  expect(await reingestLegacyEmails(directory)).toBe(0)
}, 120_000)

// The compound file's own layout (MS-CFB 2.2, 2.6): 512-byte sectors after the
// header, the first FAT sector named at 0x4C, the first directory sector at
// 0x30, and 128-byte directory entries.
function directoryLayout(file: Buffer) {
  const sector = (index: number) => (index + 1) * 512
  const fat = sector(file.readInt32LE(0x4c))
  const start = file.readInt32LE(0x30)
  return { fat, start, directory: sector(start) }
}

// The directory's first sector links to itself in the FAT.
function loopedDirectorySectors(good: Buffer) {
  const file = Buffer.from(good)
  const { fat, start } = directoryLayout(file)
  file.writeInt32LE(start, fat + 4 * start)
  return file
}

// The root's first child names itself as its next sibling.
function loopedDirectoryEntries(good: Buffer) {
  const file = Buffer.from(good)
  const { directory } = directoryLayout(file)
  const child = file.readInt32LE(directory + 0x4c)
  file.writeInt32LE(child, directory + 128 * child + 0x48)
  return file
}
