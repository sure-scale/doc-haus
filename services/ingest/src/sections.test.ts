import { afterAll, beforeAll, describe, expect, spyOn, test } from "bun:test"
import { createHash } from "node:crypto"
import { existsSync, mkdirSync, mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { Document, Packer, Paragraph } from "docx"
import { getMeta, openDb } from "./db"
import { extractDocumentText, ingestDocument } from "./ingest"
import { createMatter } from "./matter"
import { buildSectionPaths, documentSections, migrateSections, reconstructText, sectionize } from "./structure"

// Every fixture below is invented parser test content: no actual person,
// organization, transaction or obligation.

type SectionRow = ReturnType<typeof documentSections>[number]

const docx = async (paragraphs: string[]) =>
  Buffer.from(await Packer.toBuffer(new Document({ sections: [{ children: paragraphs.map((text) => new Paragraph(text)) }] })))

function expectTiling(rows: SectionRow[], length: number) {
  expect(rows.map((row) => row.ordinal)).toEqual(rows.map((_, index) => index + 1))
  expect(rows[0]?.charStart).toBe(0)
  rows.forEach((row, index) => {
    expect(row.charEnd).toBeGreaterThan(row.charStart)
    if (index > 0) expect(row.charStart).toBe(rows[index - 1]!.charEnd)
  })
  expect(rows.at(-1)?.charEnd).toBe(length)
}

test("sections tile the canonical text, starting with the opening text", () => {
  // Canonical text is the extracted text plus the newline chunk offsets count
  // after its last line.
  const canonical =
    ["  ", "Opening words before any heading.", "1. Scope", "Body.", "", "1.1 Detail", "More body."].join("\n") + "\n"
  const rows = documentSections(canonical)
  expectTiling(rows, canonical.length)
  expect(rows[0]).toMatchObject({ ordinal: 1, parentOrdinal: null, depth: 1, label: null, heading: null })
  expect(canonical.slice(rows[1]!.charStart, rows[1]!.charEnd)).toBe("1. Scope\nBody.\n\n")
  expect(rows[2]).toMatchObject({ label: "1.1", heading: "Detail", parentOrdinal: 2, depth: 2 })

  const prose = "Narrative text with no heading line.\nA second line.\n"
  expect(documentSections(prose)).toEqual([
    {
      ordinal: 1,
      parentOrdinal: null,
      depth: 1,
      label: null,
      heading: null,
      charStart: 0,
      charEnd: prose.length,
      structureSource: "sectionize",
    },
  ])
  expect(documentSections("")).toEqual([])
})

test("sections tile text that does not end with a newline, and whitespace-only text has none", () => {
  const text = ["Opening words before any heading.", "1. Scope", "Body."].join("\n")
  const rows = documentSections(text)
  expectTiling(rows, text.length)
  expect(text.slice(rows[1]!.charStart, rows[1]!.charEnd)).toBe("1. Scope\nBody.")
  // Leading blank lines belong to the first section rather than to no section.
  const blankLead = "\n\n1. Scope\nBody."
  const led = documentSections(blankLead)
  expectTiling(led, blankLead.length)
  expect(led.map((row) => row.label)).toEqual(["1"])
  for (const blank of [" ", "\n", "\n\n", "  \n\t\n "]) expect(documentSections(blank)).toEqual([])
})

test("without Markdown headings the recorded hierarchy is the structural path chunks carry", () => {
  const text = [
    "This invented agreement is a parser fixture.",
    "1. Scope",
    "The supplier provides the invented services.",
    "1.1 Services",
    "Folding paper squares.",
    "1.1.1 Materials",
    "Blue paper.",
    "1.2 Exclusions",
    "None.",
    "2. Fees",
    "No fees.",
    "SCHEDULE ONE",
    "1.3 Late item",
    "Stands alone after an unnumbered heading.",
  ].join("\n")
  const paths = buildSectionPaths("Fixture.docx", sectionize(text))
  const rows = documentSections(text + "\n")
  const crumb = (row: SectionRow) => (row.label ? `§${row.label} ${row.heading}` : (row.heading ?? "Preamble"))
  const recorded = (row: SectionRow): string => {
    const parent = rows.find((candidate) => candidate.ordinal === row.parentOrdinal)
    return `${parent ? recorded(parent) : "Doc(Fixture.docx)"} › ${crumb(row)}`
  }
  expect(rows.map(recorded)).toEqual([...paths.values()])
  expect(rows.every((row) => row.structureSource === "sectionize")).toBe(true)
})

test("Markdown heading levels the extractor emitted give the hierarchy numbering cannot", () => {
  const canonical =
    [
      "Subject: Synthetic handover note",
      "From: sender@example.com",
      "",
      "# SYNTHETIC EQUIPMENT HANDOVER",
      "This invented note exercises recorded heading levels.",
      "## 1. Inspection",
      "Read the inventory card before moving the sample box.",
      "1. Count the counters.",
      "2. Record the result.",
      "## 2. Transport sequence",
      "The carrier places the counters into the trays.",
      "### 2.1 Loading",
      "Load the green tray first.",
      "REVIEW NOTE",
      "The reviewer initials the card.",
      "# ANNEX A NOTES",
      "The annex lists no real equipment.",
    ].join("\n") + "\n"
  const rows = documentSections(canonical)
  expectTiling(rows, canonical.length)
  expect(
    rows.map(({ label, heading, parentOrdinal, depth, structureSource }) => [
      label,
      heading,
      parentOrdinal,
      depth,
      structureSource,
    ]),
  ).toEqual([
    [null, null, null, 1, "sectionize"],
    [null, "SYNTHETIC EQUIPMENT HANDOVER", null, 1, "markdown"],
    ["1", "Inspection", 2, 2, "markdown"],
    ["1", "Count the counters.", 3, 3, "sectionize"],
    ["2", "Record the result.", 3, 3, "sectionize"],
    ["2", "Transport sequence", 2, 2, "markdown"],
    ["2.1", "Loading", 6, 3, "markdown"],
    [null, "REVIEW NOTE", 7, 4, "sectionize"],
    [null, "ANNEX A NOTES", null, 1, "markdown"],
  ])
})

describe("ingest records sections and the document revision", () => {
  const filler = (topic: string) =>
    Array.from(
      { length: 30 },
      (_, i) =>
        `The ${topic} paragraph ${i + 1} records an invented step for parser regression testing and describes no actual person, organization or obligation.`,
    )
  let directory = ""
  const sources: Record<string, Buffer> = {}

  beforeAll(async () => {
    directory = realpathSync(mkdtempSync(path.join(tmpdir(), "dochaus-sections-")))
    sources["Synthetic Services.docx"] = await docx([
      "This synthetic services agreement is a parser fixture.",
      "1. Scope",
      "The supplier provides the invented services.",
      "1.1 Services",
      ...filler("services"),
      "1.2 Exclusions",
      "None.",
      "SCHEDULE ONE",
      "1 Fees",
      "Fees are invented.",
    ])
    // The empty first paragraph leaves whitespace before the first heading,
    // which no chunk covers.
    sources["Synthetic Notice.docx"] = await docx([
      "",
      "NOTICE OF HANDOVER",
      "2. Transport",
      "The carrier moves the invented box.",
    ])
    sources["Synthetic Handover.eml"] = Buffer.from(
      [
        "From: Sender <sender@example.com>",
        "To: Recipient <recipient@example.com>",
        "Subject: Synthetic handover note",
        "Date: Mon, 01 Jan 2024 00:00:00 +0000",
        "Content-Type: text/plain; charset=utf-8",
        "",
        "# SYNTHETIC EQUIPMENT HANDOVER",
        "This invented note exercises recorded heading levels.",
        "## 1. Inspection",
        "Read the inventory card before moving the sample box.",
        "## 2. Transport sequence",
        "The carrier places the counters into the trays.",
        "### 2.1 Loading",
        "Load the green tray first.",
        "## 3. Receipt check",
        ...filler("receipt"),
        "",
      ].join("\r\n"),
    )
    for (const [name, bytes] of Object.entries(sources)) await ingestDocument(directory, name, bytes)
  }, 120_000)

  afterAll(() => rmSync(directory, { recursive: true, force: true }))

  const snapshot = () => {
    const db = openDb(directory)
    try {
      return {
        documents: db.query("SELECT doc_path, name, rev FROM documents ORDER BY doc_path").all(),
        sections: db
          .query(
            "SELECT doc_path, doc_name, ordinal, parent_ordinal, depth, label, heading, char_start, char_end, structure_source FROM sections ORDER BY doc_path, ordinal",
          )
          .all(),
        chunks: db
          .query(
            "SELECT id, doc_path, section, structural_path, chunk_index, text, char_start, char_end, embedding, section_ordinal FROM chunks ORDER BY id",
          )
          .all(),
      }
    } finally {
      db.close()
    }
  }

  test("sections tile each document's text and every chunk lies in exactly one section", async () => {
    const db = openDb(directory)
    try {
      for (const [name, bytes] of Object.entries(sources)) {
        const docPath = path.join(directory, name)
        const canonical = reconstructText(db, docPath)
        const chunks = db
          .query("SELECT section, text, char_start, char_end, section_ordinal FROM chunks WHERE doc_path = ?")
          .all(docPath) as Array<{
          section: string
          text: string
          char_start: number
          char_end: number
          section_ordinal: number | null
        }>
        expect(chunks.length).toBeGreaterThan(0)
        expect(canonical.length).toBe(Math.max(...chunks.map((chunk) => chunk.char_end)))
        const rows = db
          .query(
            "SELECT ordinal, parent_ordinal AS parentOrdinal, depth, label, heading, char_start AS charStart, char_end AS charEnd, structure_source AS structureSource FROM sections WHERE doc_path = ? ORDER BY ordinal",
          )
          .all(docPath) as SectionRow[]
        expect(rows.length).toBeGreaterThan(1)
        expectTiling(rows, canonical.length)
        // The spans index the text the document route serves.
        const extracted = await extractDocumentText(name, bytes)
        for (const row of rows.filter((row) => row.heading))
          expect(extracted.slice(row.charStart, row.charEnd)).toContain(row.heading!)
        rows.forEach((row) => {
          const parent = rows.find((candidate) => candidate.ordinal === row.parentOrdinal)
          if (row.parentOrdinal === null) expect(row.depth).toBe(1)
          else expect(parent && parent.ordinal < row.ordinal && row.depth === parent.depth + 1).toBe(true)
        })
        for (const chunk of chunks) {
          const holding = rows.filter((row) => row.charStart <= chunk.char_start && chunk.char_start < row.charEnd)
          expect(holding.map((row) => row.ordinal)).toEqual([chunk.section_ordinal!])
          // The chunk was cut from that section.
          expect(chunk.section).toBe(holding[0]!.label ?? holding[0]!.heading ?? "Preamble")
          // chunkSection splits a long section's text, which already ends in a
          // newline, as if one more line followed, so the section's last chunk
          // ends one newline past the section.
          const overrun = chunk.char_end - holding[0]!.charEnd
          expect(overrun <= 0 || (overrun === 1 && chunk.text.endsWith("\n"))).toBe(true)
        }
      }
      const notice = path.join(directory, "Synthetic Notice.docx")
      const first = db.query("SELECT MIN(char_start) AS start FROM chunks WHERE doc_path = ?").get(notice) as {
        start: number
      }
      expect(first.start).toBeGreaterThan(0)
      expect(db.query("SELECT label, heading FROM sections WHERE doc_path = ? AND ordinal = 1").get(notice)).toEqual({
        label: null,
        heading: "NOTICE OF HANDOVER",
      })
      const services = path.join(directory, "Synthetic Services.docx")
      const long = db.query("SELECT ordinal FROM sections WHERE doc_path = ? AND label = '1.1'").get(services) as {
        ordinal: number
      }
      expect(
        (
          db
            .query("SELECT COUNT(*) AS n FROM chunks WHERE doc_path = ? AND section_ordinal = ?")
            .get(services, long.ordinal) as { n: number }
        ).n,
      ).toBeGreaterThan(1)
      expect(
        db
          .query("SELECT DISTINCT structure_source AS source FROM sections WHERE doc_path = ? ORDER BY source")
          .all(path.join(directory, "Synthetic Handover.eml")),
      ).toEqual([{ source: "markdown" }, { source: "sectionize" }])
    } finally {
      db.close()
    }
  })

  test("the backfill also reaches a document left without a rev and removes sections of no source document", () => {
    const fresh = snapshot()
    const notice = path.join(directory, "Synthetic Notice.docx")
    const db = openDb(directory)
    try {
      expect(getMeta(db, "sections_version")).not.toBeNull()
      expect(migrateSections(db)).toBe(0)
      db.run(
        "INSERT INTO sections (doc_path, doc_name, ordinal, depth, char_start, char_end, structure_source) VALUES (?, 'Removed.docx', 1, 1, 0, 10, 'sectionize')",
        [path.join(directory, "Removed.docx")],
      )
      expect(migrateSections(db)).toBe(0)
      expect(db.query("SELECT COUNT(*) AS n FROM sections WHERE doc_name = 'Removed.docx'").get()).toEqual({ n: 0 })
      db.run("DELETE FROM sections WHERE doc_path = ?", [notice])
      db.run("UPDATE chunks SET section_ordinal = NULL WHERE doc_path = ?", [notice])
      db.run("UPDATE documents SET rev = NULL WHERE doc_path = ?", [notice])
      expect(migrateSections(db)).toBe(1)
      expect(migrateSections(db)).toBe(0)
    } finally {
      db.close()
    }
    expect(snapshot()).toEqual(fresh)
  })

  test("a matter indexed before sections existed gets them at start, equal to a fresh ingest", () => {
    const fresh = snapshot()
    const old = openDb(directory)
    old.run("DROP TABLE sections")
    old.run("ALTER TABLE chunks DROP COLUMN section_ordinal")
    old.run("ALTER TABLE documents DROP COLUMN rev")
    old.run("DELETE FROM meta WHERE key = 'sections_version'")
    old.close()

    const db = openDb(directory)
    try {
      expect(migrateSections(db)).toBe(3)
      expect(migrateSections(db)).toBe(0)
    } finally {
      db.close()
    }
    expect(snapshot()).toEqual(fresh)
  })
})

describe("a document's rev is the hash of the canonical text its spans index", () => {
  let directory = ""
  beforeAll(() => {
    directory = realpathSync(mkdtempSync(path.join(tmpdir(), "dochaus-sections-rev-")))
  })
  afterAll(() => rmSync(directory, { recursive: true, force: true }))

  test("documents that differ only by leading blank lines get different revs", async () => {
    const body = ["NOTICE OF HANDOVER", "2. Transport", "The carrier moves the invented box."]
    const plain = await docx(body)
    const padded = await docx(["", ...body])
    expect(await extractDocumentText("Padded Notice.docx", padded)).toBe(
      "\n\n" + (await extractDocumentText("Plain Notice.docx", plain)),
    )
    await ingestDocument(directory, "Plain Notice.docx", plain)
    await ingestDocument(directory, "Padded Notice.docx", padded)
    const db = openDb(directory)
    try {
      const recorded = ["Plain Notice.docx", "Padded Notice.docx"].map((name) => {
        const docPath = path.join(directory, name)
        const canonical = reconstructText(db, docPath)
        const { rev } = db.query("SELECT rev FROM documents WHERE doc_path = ?").get(docPath) as { rev: string }
        expect(rev).toBe(createHash("sha256").update(canonical).digest("hex").slice(0, 12))
        const rows = db
          .query("SELECT label, heading, char_start, char_end FROM sections WHERE doc_path = ? ORDER BY ordinal")
          .all(docPath) as Array<{ label: string | null; heading: string; char_start: number; char_end: number }>
        expect(rows.map((row) => row.heading)).toEqual(["NOTICE OF HANDOVER", "Transport"])
        // Each span starts at the start of its heading line in the canonical text.
        for (const row of rows) {
          expect(row.char_start === 0 || canonical[row.char_start - 1] === "\n").toBe(true)
          const line = canonical.slice(row.char_start, row.char_end).split("\n")[0]!.trim()
          expect(line.startsWith(row.label ?? row.heading) && line.endsWith(row.heading)).toBe(true)
        }
        return { rev, transport: rows[1]!.char_start }
      })
      // The same pointer would land two characters apart in the two documents.
      expect(recorded[1]!.transport).toBe(recorded[0]!.transport + 2)
      expect(recorded[1]!.rev).not.toBe(recorded[0]!.rev)
    } finally {
      db.close()
    }
  })
})

test("a failed sections backfill at ingest start is logged and the upload still records its own sections", async () => {
  const directory = realpathSync(mkdtempSync(path.join(tmpdir(), "dochaus-sections-backfill-")))
  const errors = spyOn(console, "error").mockImplementation(() => {})
  try {
    await ingestDocument(directory, "First Notice.docx", await docx(["1. Scope", "The invented scope."]))
    const first = path.join(directory, "First Notice.docx")
    const setup = openDb(directory)
    const before = setup.query("SELECT ordinal, char_start, char_end FROM sections WHERE doc_path = ?").all(first)
    expect(before.length).toBeGreaterThan(0)
    // The backfill re-records the first document, which this trigger refuses.
    setup.run(
      "CREATE TRIGGER refuse_first BEFORE INSERT ON sections WHEN new.doc_name = 'First Notice.docx' BEGIN SELECT RAISE(ABORT, 'synthetic sections failure'); END",
    )
    setup.run("DELETE FROM meta WHERE key = 'sections_version'")
    setup.close()

    await ingestDocument(directory, "Second Notice.docx", await docx(["2. Transport", "The invented box moves."]))
    expect(errors.mock.calls.some(([message]) => String(message).includes("sections backfill failed"))).toBe(true)
    const db = openDb(directory)
    try {
      const second = path.join(directory, "Second Notice.docx")
      expect(db.query("SELECT label FROM sections WHERE doc_path = ?").all(second)).toEqual([{ label: "2" }])
      expect(db.query("SELECT rev FROM documents WHERE doc_path = ?").get(second)).toEqual({ rev: expect.stringMatching(/^[0-9a-f]{12}$/) })
      // The failed backfill rolled back whole and runs again at the next start.
      expect(db.query("SELECT ordinal, char_start, char_end FROM sections WHERE doc_path = ?").all(first)).toEqual(before)
      expect(getMeta(db, "sections_version")).toBeNull()
      db.run("DROP TRIGGER refuse_first")
      expect(migrateSections(db)).toBe(2)
    } finally {
      db.close()
    }
  } finally {
    errors.mockRestore()
    rmSync(directory, { recursive: true, force: true })
  }
}, 60_000)

test("legal.db connections wait for a competing writer instead of failing at once", () => {
  const directory = realpathSync(mkdtempSync(path.join(tmpdir(), "dochaus-sections-busy-")))
  try {
    const db = openDb(directory)
    expect(db.query("PRAGMA busy_timeout").get()).toEqual({ timeout: 5000 })
    db.close()
  } finally {
    rmSync(directory, { recursive: true, force: true })
  }
})

test("a sections backfill failure at service start still heals the matter's source artifacts", async () => {
  const temporary = realpathSync(mkdtempSync(path.join(tmpdir(), "dochaus-sections-boot-")))
  const workspace = path.join(temporary, "workspace")
  const config = path.join(temporary, "config")
  mkdirSync(workspace)
  mkdirSync(config)
  writeFileSync(path.join(config, "opencode.json"), "{}\n")
  const previous = process.env.WORKSPACE_ROOT
  process.env.WORKSPACE_ROOT = workspace
  let dir: string
  try {
    dir = createMatter("Synthetic boot matter").dir
  } finally {
    if (previous === undefined) delete process.env.WORKSPACE_ROOT
    else process.env.WORKSPACE_ROOT = previous
  }
  const engine = Bun.serve({ port: 0, fetch: () => new Response("No engine in this test", { status: 503 }) })
  const probe = Bun.serve({ port: 0, fetch: () => new Response() })
  const port = probe.port
  probe.stop(true)
  const serviceLog = path.join(temporary, "ingest.log")
  let service: ReturnType<typeof Bun.spawn> | undefined
  try {
    await ingestDocument(dir, "Synthetic Notice.docx", await docx(["1. Scope", "The invented scope."]))
    const sourceMap = path.join(dir, "source-map.json")
    expect(existsSync(sourceMap)).toBe(true)
    const db = openDb(dir)
    db.run("CREATE TRIGGER refuse_sections BEFORE INSERT ON sections BEGIN SELECT RAISE(ABORT, 'synthetic sections failure'); END")
    db.run("DELETE FROM meta WHERE key = 'sections_version'")
    db.close()
    rmSync(sourceMap)

    service = Bun.spawn([process.execPath, "run", "src/server.ts"], {
      cwd: path.resolve(import.meta.dir, ".."),
      env: {
        ...process.env,
        WORKSPACE_ROOT: workspace,
        DOCHAUS_DIR: config,
        DOCHAUS_RUNTIME_DIR: path.join(temporary, "runtime"),
        INGEST_PORT: String(port),
        INGEST_HOST: "127.0.0.1",
        OPENCODE_URL: `http://127.0.0.1:${engine.port}`,
      },
      stdout: Bun.file(serviceLog),
      stderr: Bun.file(serviceLog),
    })
    const deadline = Date.now() + 30_000
    while (!(await fetch(`http://127.0.0.1:${port}/matters`).then((response) => response.ok).catch(() => false))) {
      if (Date.now() >= deadline) throw new Error(`Ingest startup failed:\n${readFileSync(serviceLog, "utf8")}`)
      await Bun.sleep(50)
    }
    expect(readFileSync(serviceLog, "utf8")).toContain("matter migration failed")
    expect(existsSync(sourceMap)).toBe(true)
  } finally {
    service?.kill()
    await service?.exited
    engine.stop(true)
    rmSync(temporary, { recursive: true, force: true })
  }
}, 60_000)
