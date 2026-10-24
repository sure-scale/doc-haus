import { afterAll, expect, test } from "bun:test"
import { mkdtempSync, readdirSync, realpathSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { Document, Packer, Paragraph } from "docx"
import { strToU8, zipSync, type Zippable } from "fflate"
import { liveText } from "../../../dochaus/lib/extract"
import { parseZip } from "../../../dochaus/lib/zip"
import { openDb } from "./db"
import { ingestEmailDocument, MAX_EMAIL_DEPTH, reingestLegacyEmails, removeMatterDocument, withEmailLinks } from "./email-attachments"
import { ingestZipDocument, MAX_ZIP_FILES, ZipTooLargeError } from "./zip-archive"
import { sourceArtifactsMatchRegistry } from "./source-artifacts"

// Every archive, email and document here is invented test content built in the
// test: example.test addresses, no actual person, organization or transaction.

const directories: string[] = []
afterAll(() => directories.forEach((directory) => rmSync(directory, { recursive: true, force: true })))

function matter() {
  const directory = realpathSync(mkdtempSync(path.join(tmpdir(), "dochaus-zip-archive-")))
  directories.push(directory)
  return directory
}

const DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

const docx = async (text: string) =>
  Buffer.from(await Packer.toBuffer(new Document({ sections: [{ children: [new Paragraph(text)] }] })))

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
  return Buffer.from(text)
}

const zip = (files: Zippable) => Buffer.from(zipSync(files))

// Mark one entry password-protected (general purpose flag bit 0), as an
// encrypting archiver does, in its local and central headers.
function encrypt(archive: Buffer, name: string) {
  const out = Buffer.from(archive)
  for (let at = 0; at + 46 < out.length; at++) {
    const signature = out.readUInt32LE(at)
    if (signature === 0x04034b50 && out.toString("utf8", at + 30, at + 30 + out.readUInt16LE(at + 26)) === name)
      out.writeUInt16LE(out.readUInt16LE(at + 6) | 1, at + 6)
    if (signature === 0x02014b50 && out.toString("utf8", at + 46, at + 46 + out.readUInt16LE(at + 28)) === name)
      out.writeUInt16LE(out.readUInt16LE(at + 8) | 1, at + 8)
  }
  return out
}

let boundaries = 0
function email(subject: string, parts: { name: string; type: string; bytes: Buffer }[]) {
  const boundary = `Z${++boundaries}`
  return Buffer.from(
    [
      "From: Sender <sender@example.test>",
      "To: Recipient <recipient@example.test>",
      `Subject: ${subject}`,
      "Date: Mon, 01 Jan 2024 00:00:00 +0000",
      "MIME-Version: 1.0",
      `Content-Type: multipart/mixed; boundary="${boundary}"`,
      "",
      `--${boundary}`,
      "Content-Type: text/plain; charset=utf-8",
      "",
      "Files attached.",
      ...parts.flatMap(({ name, type, bytes }) => [
        `--${boundary}`,
        `Content-Type: ${type}; name="${name}"`,
        `Content-Disposition: attachment; filename="${name}"`,
        "Content-Transfer-Encoding: base64",
        "",
        bytes.toString("base64"),
      ]),
      `--${boundary}--`,
      "",
    ].join("\r\n"),
  )
}

function registry(directory: string) {
  const db = openDb(directory)
  try {
    return db
      .query("SELECT name, parent_name, extraction_quality FROM documents WHERE role = 'source' ORDER BY created_at, id")
      .all() as { name: string; parent_name: string | null; extraction_quality: string | null }[]
  } finally {
    db.close()
  }
}

// Each source with the container it came from, by name.
const links = (directory: string) =>
  registry(directory).map(({ name, parent_name }) => [name, parent_name]).toSorted((a, b) => a[0]!.localeCompare(b[0]!))
const names = (directory: string) => registry(directory).map((row) => row.name).toSorted()

function listed(directory: string) {
  const db = openDb(directory)
  try {
    return withEmailLinks(db, db.query("SELECT name FROM documents ORDER BY created_at, id").all() as { name: string }[])
  } finally {
    db.close()
  }
}

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

test("a zip's .docx and .pdf become documents that name it; other files are listed, folders and macOS metadata skipped", async () => {
  const directory = matter()
  const brief = await docx("Synthetic artwork brief marker")
  const proof = pdf("Synthetic signed proof marker")
  const logo = Buffer.from("synthetic image bytes")
  const archive = zip({
    "Artwork/": new Uint8Array(),
    "Artwork/Brief.docx": brief,
    "Artwork/Proof.pdf": proof,
    "Artwork/logo.png": logo,
    "Artwork/.DS_Store": strToU8("finder"),
    "__MACOSX/Artwork/._Brief.docx": strToU8("resource fork"),
  })

  const result = await ingestZipDocument(directory, "Artwork files.zip", archive)

  expect(result.name).toBe("Artwork files.zip")
  // A container is never reported as unreadable.
  expect(result.quality).toEqual({ status: "ok" })
  expect(result.attachments).toEqual([
    { name: "Artwork/Brief.docx", type: DOCX, size: brief.length, document: "Brief.docx", status: "indexed" },
    { name: "Artwork/Proof.pdf", type: "application/pdf", size: proof.length, document: "Proof.pdf", status: "indexed" },
    { name: "Artwork/logo.png", type: "application/octet-stream", size: logo.length, status: "unsupported" },
  ])
  expect(links(directory)).toEqual([
    ["Artwork files.zip", null],
    ["Brief.docx", "Artwork files.zip"],
    ["Proof.pdf", "Artwork files.zip"],
  ])
  // The zip's text is the listing of its files, derived from its bytes, and it
  // re-extracts to the same text for citations.
  const text = await liveText(path.join(directory, "Artwork files.zip"))
  expect(text).toBe(`Contents:\nArtwork/Brief.docx (${(brief.length / 1024).toFixed(1)} KB)\nArtwork/Proof.pdf (${proof.length} bytes)\nArtwork/logo.png (${logo.length} bytes)\n`)
  await expectCitable(directory, "Artwork files.zip")
  await expectCitable(directory, "Brief.docx")
  await expectCitable(directory, "Proof.pdf")
  expect(sourceArtifactsMatchRegistry(directory)).toBe(true)
  // Only the documents are unpacked into the matter.
  expect(readdirSync(directory).filter((name) => !name.startsWith(".") && !/^(source|matter-|sources)/.test(name)).toSorted())
    .toEqual(["Artwork files.zip", "Brief.docx", "Proof.pdf"])
  expect(listed(directory)).toContainEqual(expect.objectContaining({ name: "Brief.docx", attachment_of: { name: "Artwork files.zip", removed: false } }))

  // Removing the zip removes the documents it brought in.
  const removal = await removeMatterDocument(directory, "Artwork files.zip", false)
  expect(removal.attachments.removed.toSorted()).toEqual(["Brief.docx", "Proof.pdf"])
  expect(registry(directory)).toEqual([])
})

test("a zip attached to an email, and an email inside a zip, each bring in their files", async () => {
  const directory = matter()
  const brief = await docx("Synthetic zipped brief marker")
  const archive = zip({ "Brief.docx": brief, "notes.txt": strToU8("plain notes") })
  const message = email("Synthetic artwork", [{ name: "Artwork files.zip", type: "application/zip", bytes: archive }])

  const result = await ingestEmailDocument(directory, "Thread.eml", message)

  expect(result.attachments).toEqual([
    { name: "Artwork files.zip", type: "application/zip", size: archive.length, document: "Artwork files.zip", status: "indexed" },
  ])
  expect(links(directory)).toEqual([
    ["Artwork files.zip", "Thread.eml"],
    ["Brief.docx", "Artwork files.zip"],
    ["Thread.eml", null],
  ])

  const other = matter()
  const schedule = await docx("Synthetic schedule inside a zipped email")
  const zipped = zip({ "Mail/Forward.eml": email("Synthetic forward", [{ name: "Schedule.docx", type: DOCX, bytes: schedule }]) })
  await ingestZipDocument(other, "Mailbox.zip", zipped)
  expect(links(other)).toEqual([
    ["Forward.eml", "Mailbox.zip"],
    ["Mailbox.zip", null],
    ["Schedule.docx", "Forward.eml"],
  ])
})

test("zips nested past the shared depth limit are listed, not unpacked", async () => {
  const directory = matter()
  // MAX_EMAIL_DEPTH zips below the uploaded one are unpacked; the next is listed.
  let archive = zip({ "Deep.docx": await docx("Synthetic deepest marker") })
  for (let level = MAX_EMAIL_DEPTH + 1; level >= 1; level--) archive = zip({ [`Level ${level}.zip`]: archive })

  await ingestZipDocument(directory, "Outer.zip", archive)

  const levels = Array.from({ length: MAX_EMAIL_DEPTH }, (_, index) => `Level ${index + 1}.zip`)
  expect(names(directory)).toEqual(["Outer.zip", ...levels].toSorted())
  const last = listed(directory).find((row) => row.name === levels.at(-1)) as { attachments?: { name: string; status: string }[] }
  expect(last.attachments).toEqual([expect.objectContaining({ name: `Level ${MAX_EMAIL_DEPTH + 1}.zip`, status: "nested" })])
})

test("a file whose path leaves the archive, or that is password-protected, is listed and not added", async () => {
  const directory = matter()
  const archive = encrypt(
    zip({
      "../../Escape.docx": await docx("Synthetic escape marker"),
      "/abs/Rooted.pdf": pdf("Synthetic rooted marker"),
      "Locked.docx": await docx("Synthetic locked marker"),
      "Open.docx": await docx("Synthetic open marker"),
    }),
    "Locked.docx",
  )

  const result = await ingestZipDocument(directory, "Mixed.zip", archive)

  expect(result.attachments?.map(({ name, status, document }) => ({ name, status, document }))).toEqual([
    { name: "../../Escape.docx", status: "unsafe_path", document: undefined },
    { name: "/abs/Rooted.pdf", status: "unsafe_path", document: undefined },
    { name: "Locked.docx", status: "encrypted", document: undefined },
    { name: "Open.docx", status: "indexed", document: "Open.docx" },
  ])
  expect(names(directory)).toEqual(["Mixed.zip", "Open.docx"])
})

test("a zip over the limits is refused whole; a nested one is listed as too large", async () => {
  const directory = matter()
  const many = zip(Object.fromEntries(Array.from({ length: MAX_ZIP_FILES + 1 }, (_, index) => [`f${index}.txt`, strToU8("x")])))
  await expect(ingestZipDocument(directory, "Many.zip", many)).rejects.toThrow(ZipTooLargeError)
  await expect(ingestZipDocument(directory, "Big.zip", zip({ "a.docx": await docx("Synthetic big marker") }), { budget: { bytes: 100, files: 10 } }))
    .rejects.toThrow("would unpack to more than")
  expect(registry(directory)).toEqual([])
  expect(readdirSync(directory).filter((name) => !name.startsWith("."))).toEqual([])

  // The budget is shared down the tree, so a nested zip cannot multiply it.
  const inner = zip({ "filler.txt": [strToU8("y".repeat(5000)), { level: 0 }] })
  const outer = zip({ "Inner.zip": [inner, { level: 0 }], "Kept.docx": await docx("Synthetic kept marker") })
  const result = await ingestZipDocument(directory, "Outer.zip", outer, { budget: { bytes: parseZip(outer).uncompressed + 100, files: 10 } })
  expect(result.attachments?.map(({ name, status }) => [name, status])).toEqual([
    ["Inner.zip", "too_large"],
    ["Kept.docx", "indexed"],
  ])
})

test("an email whose zip attachment was listed as not supported is re-read once at start, bringing its files in", async () => {
  const directory = matter()
  const archive = zip({ "Brief.docx": await docx("Synthetic earlier brief marker") })
  await ingestEmailDocument(directory, "Earlier.eml", email("Synthetic earlier", [{ name: "Artwork files.zip", type: "application/zip", bytes: archive }]))
  // The state a build that did not read zips leaves: the zip listed as not
  // supported by its extension, nothing extracted from it.
  await removeMatterDocument(directory, "Artwork files.zip", false)
  const db = openDb(directory)
  db.run("UPDATE documents SET attachments = ? WHERE name = 'Earlier.eml'", [
    JSON.stringify({ attachments: [{ name: "Artwork files.zip", type: "application/zip", size: archive.length, outcome: "unsupported", extension: ".zip" }], embedded: 0 }),
  ])
  db.close()

  expect(await reingestLegacyEmails(directory)).toBe(1)
  expect(links(directory)).toEqual([
    ["Artwork files.zip", "Earlier.eml"],
    ["Brief.docx", "Artwork files.zip"],
    ["Earlier.eml", null],
  ])
  expect(await reingestLegacyEmails(directory)).toBe(0)
})
