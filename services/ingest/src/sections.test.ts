import { afterAll, beforeAll, describe, expect, spyOn, test } from "bun:test"
import { createHash } from "node:crypto"
import { existsSync, mkdirSync, mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { Database } from "bun:sqlite"
import { Document, Packer, Paragraph } from "docx"
import { ftsInSync, getMeta, openDb } from "./db"
import { extractDocumentText, ingestDocument } from "./ingest"
import { createMatter } from "./matter"
import { buildSectionPaths, chunkSection, documentSections, migrateSections, reconstructText, sectionize } from "./structure"

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

test("a split section's chunks are its own text at their offsets", () => {
  const paragraph = (i: number) =>
    `Paragraph ${i} of an invented clause, written only to exceed one chunk and describing no actual obligation.`
  const text = ["1. Long clause", ...Array.from({ length: 60 }, (_, i) => paragraph(i + 1)), "2. Next clause", "Short."].join("\n") + "\n"
  const [long, next] = sectionize(text)
  const chunks = chunkSection(long!)
  expect(chunks.length).toBeGreaterThan(1)
  for (const chunk of chunks) expect(text.slice(chunk.charStart, chunk.charEnd)).toBe(chunk.text)
  expect(chunks.at(-1)!.charEnd).toBe(next!.charStart)
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
          // A chunk is the canonical text at its offsets and ends within its
          // section, including the last chunk of a section split in several.
          expect(chunk.char_end).toBeLessThanOrEqual(holding[0]!.charEnd)
          expect(canonical.slice(chunk.char_start, chunk.char_end)).toBe(chunk.text)
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

  test("a matter recorded by sections version 1 is re-recorded, replacing its revs", () => {
    const fresh = snapshot()
    const db = openDb(directory)
    try {
      // Version 1 hashed the trimmed text and rebuilt it with truncating overwrites.
      db.run("UPDATE meta SET value = '1' WHERE key = 'sections_version'")
      db.run("UPDATE documents SET rev = 'stale0000000' WHERE role = 'source'")
      expect(migrateSections(db)).toBe(3)
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

describe("legal.db write lock", () => {
  let directory = ""
  beforeAll(() => {
    directory = realpathSync(mkdtempSync(path.join(tmpdir(), "dochaus-sections-busy-")))
    openDb(directory).close()
  })
  afterAll(() => rmSync(directory, { recursive: true, force: true }))

  // Another connection holding the write lock, as an ingest commit does across
  // its awaited source-artifact write in this same process.
  const holdWriteLock = () => {
    const holder = new Database(path.join(directory, ".dochaus", "legal.db"))
    holder.run("BEGIN IMMEDIATE")
    holder.run("INSERT INTO meta (key, value) VALUES ('synthetic-lock', '1')")
    return { release: () => (holder.run("ROLLBACK"), holder.close()) }
  }

  test("a request connection opens and reads under another connection's write lock, and its own write fails fast", () => {
    const lock = holdWriteLock()
    try {
      // A long wait here would block the event loop the lock holder needs to finish.
      const db = openDb(directory)
      try {
        expect(db.query("PRAGMA busy_timeout").get()).toEqual({ timeout: 200 })
        expect(db.query("SELECT COUNT(*) AS n FROM chunks").get()).toEqual({ n: 0 })
        const started = performance.now()
        expect(() =>
          db.run("INSERT INTO answer_drafts (scope, answers, revision) VALUES ('synthetic', '{}', 1)"),
        ).toThrow("database is locked")
        expect(performance.now() - started).toBeLessThan(1000)
      } finally {
        db.close()
      }
    } finally {
      lock.release()
    }
  }, 30_000)

  // Another process's last connection checkpoints and removes the WAL as it
  // closes, holding the database briefly; a request open in that moment waits
  // it out instead of failing.
  test("request opens never fail while another process opens, writes and closes in a loop", async () => {
    const file = path.join(directory, ".dochaus", "legal.db")
    const runMs = 2000
    const writer = Bun.spawn(
      [
        process.execPath,
        "-e",
        `const { Database } = require("bun:sqlite")
        const until = Date.now() + ${runMs}
        let n = 0
        while (Date.now() < until) {
          const db = new Database(${JSON.stringify(file)})
          db.run("PRAGMA busy_timeout = 2000")
          db.run("PRAGMA journal_mode = WAL")
          db.run("INSERT INTO meta (key, value) VALUES (?, '1') ON CONFLICT(key) DO UPDATE SET value = excluded.value", ["synthetic-" + (n++ % 50)])
          db.close()
          Bun.sleepSync(2)
        }`,
      ],
      { stderr: "pipe" },
    )
    let opens = 0
    const failures: string[] = []
    try {
      const until = Date.now() + runMs
      while (Date.now() < until) {
        try {
          openDb(directory).close()
          opens++
        } catch (error) {
          failures.push(String((error as { code?: unknown }).code))
        }
        await Bun.sleep(1)
      }
    } finally {
      await writer.exited
    }
    expect(writer.exitCode).toBe(0)
    expect(opens).toBeGreaterThan(100)
    expect(failures).toEqual([])
  }, 30_000)

  test("a held write lock is an error to the lexical index check, never drift", () => {
    const db = new Database(path.join(directory, ".dochaus", "legal.db"))
    const lock = holdWriteLock()
    try {
      expect(() => ftsInSync(db)).toThrow("database is locked")
    } finally {
      lock.release()
      db.close()
    }
  })

  test("the startup connection waits for another process's write lock", async () => {
    const holder = Bun.spawn(
      [
        process.execPath,
        "-e",
        `const { Database } = require("bun:sqlite")
        const db = new Database(${JSON.stringify(path.join(directory, ".dochaus", "legal.db"))})
        db.run("BEGIN IMMEDIATE")
        db.run("INSERT INTO meta (key, value) VALUES ('synthetic-holder', '1')")
        console.log("locked")
        setTimeout(() => { db.run("ROLLBACK"); db.close() }, 300)`,
      ],
      { stdout: "pipe" },
    )
    try {
      const reader = holder.stdout.getReader()
      expect(new TextDecoder().decode((await reader.read()).value)).toContain("locked")
      reader.releaseLock()
      const db = openDb(directory, { startup: true })
      try {
        expect(db.query("PRAGMA busy_timeout").get()).toEqual({ timeout: 5000 })
        expect(ftsInSync(db)).toBe(true)
      } finally {
        db.close()
      }
    } finally {
      holder.kill()
      await holder.exited
    }
  }, 30_000)

  test("the startup connection rebuilds a desynced lexical index, and a lexical index an open creates is built at once", () => {
    const chunk = (db: Database, text: string) =>
      db.run(
        "INSERT INTO chunks (document_id, doc_path, doc_name, section, chunk_index, text, char_start, char_end, embedding) VALUES (1, 'synthetic.docx', 'synthetic.docx', '1', 0, ?, 0, ?, x'')",
        [text, text.length],
      )
    const match = (db: Database, term: string) => db.query("SELECT rowid FROM chunks_fts WHERE chunks_fts MATCH ?").all(term).length
    const raw = new Database(path.join(directory, ".dochaus", "legal.db"))
    raw.run("DROP TRIGGER chunks_fts_ai")
    chunk(raw, "invented gamma clause")
    raw.close()
    const startup = openDb(directory, { startup: true })
    try {
      expect(ftsInSync(startup)).toBe(true)
      expect(match(startup, "gamma")).toBe(1)
    } finally {
      startup.close()
    }
    // A database from before the lexical index existed.
    const older = new Database(path.join(directory, ".dochaus", "legal.db"))
    for (const trigger of ["chunks_fts_ai", "chunks_fts_ad", "chunks_fts_au"]) older.run(`DROP TRIGGER ${trigger}`)
    older.run("DROP TABLE chunks_fts")
    chunk(older, "invented delta clause")
    older.close()
    const db = openDb(directory)
    try {
      expect(ftsInSync(db)).toBe(true)
      expect(match(db, "delta")).toBe(1)
    } finally {
      db.close()
    }
  })
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
    // A lexical index that lost a chunk's entry, healed by the startup check.
    db.run(
      "INSERT INTO chunks_fts(chunks_fts, rowid, text, section, doc_name) SELECT 'delete', id, text, section, doc_name FROM chunks LIMIT 1",
    )
    expect(ftsInSync(db)).toBe(false)
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
    const healed = openDb(dir)
    try {
      expect(ftsInSync(healed)).toBe(true)
    } finally {
      healed.close()
    }
  } finally {
    service?.kill()
    await service?.exited
    engine.stop(true)
    rmSync(temporary, { recursive: true, force: true })
  }
}, 60_000)

describe("chunk index maintenance", () => {
  let directory = ""
  beforeAll(() => {
    directory = realpathSync(mkdtempSync(path.join(tmpdir(), "dochaus-sections-index-")))
  })
  afterAll(() => rmSync(directory, { recursive: true, force: true }))

  test("a document's chunks are read in offset order through an index", () => {
    const db = openDb(directory)
    try {
      const plan = (
        db
          .query("EXPLAIN QUERY PLAN SELECT id, text, char_start FROM chunks WHERE doc_path = ? ORDER BY char_start, id")
          .all("synthetic.docx") as Array<{ detail: string }>
      )
        .map((row) => row.detail)
        .join("\n")
      expect(plan).toMatch(/USING INDEX/)
      expect(plan).not.toMatch(/TEMP B-TREE/)
    } finally {
      db.close()
    }
  })

  test("the lexical index follows chunk text but not a section ordinal", () => {
    // A database from before: its update trigger fired on any column.
    const older = openDb(directory)
    older.run("DROP TRIGGER chunks_fts_au")
    older.run(`CREATE TRIGGER chunks_fts_au AFTER UPDATE ON chunks BEGIN
      INSERT INTO chunks_fts(chunks_fts, rowid, text, section, doc_name) VALUES ('delete', old.id, old.text, old.section, old.doc_name);
      INSERT INTO chunks_fts(rowid, text, section, doc_name) VALUES (new.id, new.text, new.section, new.doc_name);
    END`)
    older.close()
    const db = openDb(directory)
    try {
      const id = Number(
        db.run(
          "INSERT INTO chunks (document_id, doc_path, doc_name, section, chunk_index, text, char_start, char_end, embedding) VALUES (1, 'synthetic.docx', 'synthetic.docx', '1', 0, 'invented alpha clause', 0, 21, x'')",
        ).lastInsertRowid,
      )
      const match = (term: string) =>
        (db.query("SELECT rowid FROM chunks_fts WHERE chunks_fts MATCH ?").all(term) as Array<{ rowid: number }>).map(
          (row) => row.rowid,
        )
      expect(match("alpha")).toEqual([id])
      db.run("UPDATE chunks SET text = 'invented omega clause' WHERE id = ?", [id])
      expect(match("alpha")).toEqual([])
      expect(match("omega")).toEqual([id])
      // A section ordinal write changes the one chunk row and nothing else: no
      // trigger rewrites the row's lexical index entry.
      const changes = () => (db.query("SELECT total_changes() AS n").get() as { n: number }).n
      const before = changes()
      db.run("UPDATE chunks SET section_ordinal = 1 WHERE id = ?", [id])
      expect(changes() - before).toBe(1)
      expect(match("omega")).toEqual([id])
    } finally {
      db.close()
    }
    // Reopening keeps one current trigger rather than adding or reverting it.
    openDb(directory).close()
    const reopened = openDb(directory)
    try {
      expect(
        reopened.query("SELECT COUNT(*) AS n FROM sqlite_master WHERE type = 'trigger' AND tbl_name = 'chunks'").get(),
      ).toEqual({ n: 3 })
    } finally {
      reopened.close()
    }
  })
})

describe("reconstructText", () => {
  let directory = ""
  beforeAll(() => {
    directory = realpathSync(mkdtempSync(path.join(tmpdir(), "dochaus-sections-reconstruct-")))
  })
  afterAll(() => rmSync(directory, { recursive: true, force: true }))

  const store = (docPath: string, chunks: Array<{ text: string; start: number }>) => {
    const db = openDb(directory)
    try {
      db.transaction(() => {
        chunks.forEach((chunk, index) =>
          db.run(
            "INSERT INTO chunks (document_id, doc_path, doc_name, section, chunk_index, text, char_start, char_end, embedding) VALUES (1, ?, ?, '1', ?, ?, ?, ?, x'')",
            [docPath, docPath, index, chunk.text, chunk.start, chunk.start + chunk.text.length],
          ),
        )
      })()
    } finally {
      db.close()
    }
  }
  const rebuilt = (docPath: string) => {
    const db = openDb(directory)
    try {
      return reconstructText(db, docPath)
    } finally {
      db.close()
    }
  }

  test("chunks with the same start give the exact text in either stored order", () => {
    const text = "1. Scope\nThe invented scope runs to here.\n"
    const long = { text, start: 0 }
    const short = { text: text.slice(0, 9), start: 0 }
    store("tied-short-last.docx", [long, short])
    store("tied-short-first.docx", [short, long])
    expect(rebuilt("tied-short-last.docx")).toBe(text)
    expect(rebuilt("tied-short-first.docx")).toBe(text)
  })

  test("the next chunk replaces the extra newline an older split section's last chunk carried", () => {
    // Chunks stored before chunkSection stopped counting one newline past a
    // split section: the first section's last chunk ends with "\n" where the
    // second section's heading starts.
    const text = "1. First\nAlpha line.\nBeta line.\n2. Second\nGamma line.\n"
    const second = text.indexOf("2. Second")
    store("older-split.docx", [
      { text: text.slice(0, 21), start: 0 },
      { text: text.slice(9, second) + "\n", start: 9 },
      { text: text.slice(second), start: second },
    ])
    expect(rebuilt("older-split.docx")).toBe(text)
  })

  test("uncovered gaps are padded with spaces so offsets after them hold", () => {
    store("gap.docx", [
      { text: "NOTICE\n", start: 2 },
      { text: "1. Scope\n", start: 12 },
    ])
    expect(rebuilt("gap.docx")).toBe("  NOTICE\n   1. Scope\n")
  })

  test("a 2000-chunk document reconstructs in about the time it takes to read its chunks", () => {
    // Rewriting the whole string per chunk took about 75 times as long as
    // reading the same rows here; the linear rebuild takes about as long. The
    // bound compares two costs measured together on the same machine rather
    // than a fixed time, and the fastest of several runs discards a pause from
    // garbage collection or a loaded machine.
    const line = (i: number) => `Line ${i} of an invented clause that describes no actual obligation.\n`
    const text = Array.from({ length: 60_000 }, (_, i) => line(i)).join("")
    const count = 2000
    const step = Math.floor(text.length / count)
    store(
      "large.docx",
      Array.from({ length: count }, (_, i) => {
        const start = i * step
        const end = i === count - 1 ? text.length : Math.min(text.length, start + step + 250)
        return { text: text.slice(start, end), start }
      }),
    )
    const fastest = (run: () => unknown) =>
      Math.min(
        ...Array.from({ length: 5 }, () => {
          const started = performance.now()
          run()
          return performance.now() - started
        }),
      )
    const db = openDb(directory)
    try {
      expect(reconstructText(db, "large.docx")).toBe(text)
      const read = fastest(() =>
        (
          db
            .query("SELECT text, char_start FROM chunks WHERE doc_path = ? ORDER BY char_start, id")
            .all("large.docx") as Array<{ text: string }>
        )
          .map((row) => row.text)
          .join(""),
      )
      expect(fastest(() => reconstructText(db, "large.docx"))).toBeLessThan(10 * read)
    } finally {
      db.close()
    }
  })
})
