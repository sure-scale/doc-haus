import { afterAll, expect, test } from "bun:test"
import { existsSync, mkdtempSync, readFileSync, realpathSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { Document, Packer, Paragraph } from "docx"
import { utils, write, type WorkSheet } from "xlsx"
import { openDb } from "./db"
import { ingestEmailDocument, withEmailLinks } from "./email-attachments"
import { extractDocumentText, ingestDocument, sourceTextRevision } from "./ingest"
import { chunkSection, sectionize } from "./structure"

// Every file here is synthetic, built in the test: invented wording,
// example.test addresses, no actual person, organization or transaction.

const directories: string[] = []
afterAll(() => directories.forEach((directory) => rmSync(directory, { recursive: true, force: true })))

function matter() {
  const directory = realpathSync(mkdtempSync(path.join(tmpdir(), "dochaus-extraction-quality-")))
  directories.push(directory)
  return directory
}

const docx = async (paragraphs: string[]) =>
  Buffer.from(await Packer.toBuffer(new Document({ sections: [{ children: paragraphs.map((text) => new Paragraph(text)) }] })))

function xlsx(sheet: WorkSheet) {
  const workbook = utils.book_new()
  utils.book_append_sheet(workbook, sheet, "Sheet1")
  return Buffer.from(write(workbook, { type: "buffer", bookType: "xlsx" }))
}

// A sheet of formulas saved without calculated values, as a program that writes
// a workbook without opening it in a spreadsheet leaves them; `labels` adds a
// column of text cells beside them.
function formulaSheet(labels: boolean) {
  const sheet: WorkSheet = { "!ref": labels ? "A1:B2" : "B1:B2" }
  sheet.B1 = { t: "n", f: "SUM(C1:C9)" }
  sheet.B2 = { t: "n", f: "B1*2" }
  if (labels) {
    sheet.A1 = { t: "s", v: "Synthetic subtotal" }
    sheet.A2 = { t: "s", v: "Synthetic total" }
  }
  return sheet
}

const HEADERS = ["From: Sender <sender@example.test>", "To: Recipient <recipient@example.test>", "Subject: Synthetic terms", "Date: Mon, 01 Jan 2024 00:00:00 +0000", "MIME-Version: 1.0"]
const eml = (lines: string[]) => Buffer.from([...HEADERS, ...lines, ""].join("\r\n"))

const fixture = (name: string) => readFileSync(path.join(import.meta.dir, "fixtures", name))

const PROSE = [
  "1. Scope. The supplier provides the synthetic services described in the schedule.",
  "2. Term. This synthetic agreement runs for twelve months from the start date.",
  "3. Notices. Each notice is given in writing to the address stated above.",
]

function stored(directory: string, name: string) {
  const db = openDb(directory)
  try {
    const document = db.query("SELECT role, rev, injection_report, extraction_quality FROM documents WHERE name = ?").get(name) as {
      role: string
      rev: string
      injection_report: string | null
      extraction_quality: string | null
    }
    const chunks = db.query("SELECT text, char_start, char_end FROM chunks WHERE doc_name = ? ORDER BY char_start").all(name) as {
      text: string
      char_start: number
      char_end: number
    }[]
    return { ...document, quality: document.extraction_quality && JSON.parse(document.extraction_quality), chunks }
  } finally {
    db.close()
  }
}

// What the documents route lists for each source.
function listed(directory: string) {
  const db = openDb(directory)
  try {
    return withEmailLinks(db, db.query("SELECT name FROM documents ORDER BY created_at, id").all() as { name: string }[])
  } finally {
    db.close()
  }
}

test("a blank or garbled file is kept and registered with nothing indexed", async () => {
  const directory = matter()
  const cases = [
    ["Empty.docx", await docx([]), "empty"],
    ["Garbled.docx", await docx(["�".repeat(400)]), "replacement_characters"],
    ["Formulas.xlsx", xlsx(formulaSheet(false)), "formulas_without_values"],
  ] as const
  for (const [name, bytes, code] of cases) {
    const result = await ingestDocument(directory, name, bytes)
    expect(result.quality).toEqual({ status: "no_readable_text", signals: [expect.objectContaining({ code })] })
    expect(result.chunks).toBe(0)
    // A quality result is never an injection finding.
    expect(result.injection).toBeNull()
    const row = stored(directory, name)
    expect(row.role).toBe("source")
    expect(row.chunks).toEqual([])
    expect(row.injection_report).toBeNull()
    expect(row.quality.status).toBe("no_readable_text")
    expect(existsSync(path.join(directory, name))).toBe(true)
  }
})

test("an email with no message body is indexed by its headers and marked degraded", async () => {
  const directory = matter()
  const noBody = eml(["Content-Type: text/plain", ""])
  // A multipart whose boundary never appears: the parser yields headers only.
  const malformed = eml(['Content-Type: multipart/mixed; boundary="MISSING"', "", "--OTHER", "Content-Type: text/plain", "", "Synthetic body text", "--OTHER--"])
  for (const [name, bytes] of [["NoBody.eml", noBody], ["Malformed.eml", malformed]] as const) {
    const result = await ingestEmailDocument(directory, name, bytes)
    expect(result.quality).toEqual({ status: "degraded", signals: [expect.objectContaining({ code: "no_email_body" })] })
    const row = stored(directory, name)
    expect(row.chunks.map((chunk) => chunk.text).join("")).toContain("Subject: Synthetic terms")
    expect(row.quality.status).toBe("degraded")
  }
  // A completely empty file has no headers either.
  expect((await ingestEmailDocument(directory, "Blank.eml", Buffer.alloc(0))).quality.status).toBe("no_readable_text")
})

test("partly garbled text and partly computed workbooks are indexed and marked degraded", async () => {
  const directory = matter()
  const garbled = await ingestDocument(directory, "Partly garbled.docx", await docx([...PROSE, "�".repeat(30)]))
  expect(garbled.quality).toEqual({ status: "degraded", signals: [expect.objectContaining({ code: "replacement_characters" })] })
  expect(garbled.chunks).toBeGreaterThan(0)
  const workbook = await ingestDocument(directory, "Partly computed.xlsx", xlsx(formulaSheet(true)))
  expect(workbook.quality).toEqual({
    status: "degraded",
    signals: [{ code: "formulas_without_values", detail: "2 of 4 filled cells are formulas saved without a calculated value" }],
  })
  expect(workbook.chunks).toBeGreaterThan(0)
})

// Bun.which reads the PATH the process started with, so the ingest runs in a
// child process started with an empty PATH: no OCR, no markitdown.
test("a scanned PDF with no OCR on PATH is not indexed as a normal source", async () => {
  const directory = matter()
  const script = path.join(directory, "..", `${path.basename(directory)}-scan.ts`)
  directories.push(script)
  await Bun.write(script, [
    `import { readFileSync } from "node:fs"`,
    `import { ingestDocument } from ${JSON.stringify(path.join(import.meta.dir, "ingest.ts"))}`,
    `const result = await ingestDocument(${JSON.stringify(directory)}, "Scan.pdf", readFileSync(${JSON.stringify(path.join(import.meta.dir, "fixtures", "synthetic-scanned.pdf"))}))`,
    `console.log(JSON.stringify(result.quality))`,
  ].join("\n"))
  const child = Bun.spawn([process.execPath, "run", script], { env: { ...process.env, PATH: "" }, stdout: "pipe", stderr: "pipe" })
  const output = await new Response(child.stdout).text()
  expect(await child.exited).toBe(0)
  const quality = JSON.parse(output.trim().split("\n").at(-1)!)
  expect(["no_readable_text", "degraded"]).toContain(quality.status)
  expect(stored(directory, "Scan.pdf").quality).toEqual(quality)
}, 120_000)

test("normal DOCX, PDF, XLSX and EML ingest as ok with chunks at the extracted text's offsets", async () => {
  const directory = matter()
  const files = [
    ["Agreement.docx", await docx(PROSE)],
    ["Plan.pdf", fixture("synthetic-text-layer.pdf")],
    ["Register.xlsx", xlsx(utils.aoa_to_sheet([["Holder", "Class", "Shares"], ["Synthetic Holder A", "Ordinary", 1000], ["Synthetic Holder B", "Preferred", 250]]))],
    ["Thread.eml", eml(["Content-Type: text/plain; charset=utf-8", "", ...PROSE])],
  ] as const
  for (const [name, bytes] of files) {
    const result = name.endsWith(".eml") ? await ingestEmailDocument(directory, name, bytes) : await ingestDocument(directory, name, bytes)
    expect(result.quality).toEqual({ status: "ok" })
    // The chunks and offsets are those the extracted text always cut to.
    const text = await extractDocumentText(name, bytes)
    const row = stored(directory, name)
    expect(row.quality).toEqual({ status: "ok" })
    expect(row.chunks.length).toBeGreaterThan(0)
    expect(row.chunks).toEqual(
      sectionize(text).flatMap((section) => chunkSection(section)).map((chunk) => ({ text: chunk.text, char_start: chunk.charStart, char_end: chunk.charEnd })),
    )
    expect(row.rev).toBe(sourceTextRevision(text))
  }
}, 120_000)

test("email attachments pass the same gate, and one with no readable text is listed as such", async () => {
  const directory = matter()
  const DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
  const part = (name: string, bytes: Buffer) =>
    [`Content-Type: ${DOCX}; name="${name}"`, `Content-Disposition: attachment; filename="${name}"`, "Content-Transfer-Encoding: base64", "", bytes.toString("base64")].join("\r\n")
  const message = eml([
    'Content-Type: multipart/mixed; boundary="B"',
    "",
    "--B",
    "Content-Type: text/plain; charset=utf-8",
    "",
    "Please see the synthetic attachments.",
    "--B",
    part("Blank.docx", await docx([])),
    "--B",
    part("Garbled.docx", await docx(["�".repeat(200)])),
    "--B",
    part("Partly garbled.docx", await docx([...PROSE, "�".repeat(30)])),
    "--B--",
  ])
  const result = await ingestEmailDocument(directory, "Thread.eml", message)
  expect(result.quality).toEqual({ status: "ok" })
  expect(result.attachments.map(({ name, status, document, degraded }) => ({ name, status, document, degraded }))).toEqual([
    // Kept as documents, nothing indexed; two unreadable files are not the same document.
    { name: "Blank.docx", status: "no_readable_text", document: "Blank.docx", degraded: undefined },
    { name: "Garbled.docx", status: "no_readable_text", document: "Garbled.docx", degraded: undefined },
    { name: "Partly garbled.docx", status: "indexed", document: "Partly garbled.docx", degraded: true },
  ])
  expect(stored(directory, "Blank.docx").chunks).toEqual([])
  expect(stored(directory, "Garbled.docx").chunks).toEqual([])
  expect(stored(directory, "Partly garbled.docx").chunks.length).toBeGreaterThan(0)
  expect(listed(directory).find((document) => document.name === "Thread.eml")?.attachments?.map((attachment) => attachment.status)).toEqual([
    "no_readable_text",
    "no_readable_text",
    "indexed",
  ])
})

test("an email that only carries attachments is ok", async () => {
  const directory = matter()
  const DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
  const bytes = await docx(PROSE)
  const message = eml([
    'Content-Type: multipart/mixed; boundary="B"',
    "",
    "--B",
    `Content-Type: ${DOCX}; name="Agreement.docx"`,
    'Content-Disposition: attachment; filename="Agreement.docx"',
    "Content-Transfer-Encoding: base64",
    "",
    bytes.toString("base64"),
    "--B--",
  ])
  expect((await ingestEmailDocument(directory, "Cover.eml", message)).quality).toEqual({ status: "ok" })
})
