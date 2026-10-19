import { describe, expect, spyOn, test } from "bun:test"
import { createHash } from "node:crypto"
import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { insertChunk, openDb, upsertDocument } from "./db"
import * as identifierExtraction from "./identifiers"
import { stampIngestLineage } from "./lineage"
import { healSourceArtifacts, sourceArtifactsMatchRegistry, writeSourceArtifacts } from "./source-artifacts"

// The longest line a source chunk file may hold (the read-tool cap in source-artifacts.ts).
const READ_LINE_CAP = 1000

// Undo the soft wrap of a chunk file against its source slice. A line break the
// wrap put in place of a space turns back into that space; a line break it
// inserted between two characters is dropped and reported, so the caller can
// check it sits inside a single run of more than READ_LINE_CAP non-space
// characters. Any other difference is copied through and fails the comparison.
function unwrapChunk(original: string, wrapped: string) {
  let restored = ""
  let at = 0
  const inserted: number[] = []
  for (const char of wrapped) {
    if (char !== "\n" || original.startsWith("\n", at)) {
      restored += char
      at += char.length
    } else if (original[at] === " ") {
      restored += " "
      at += 1
    } else {
      inserted.push(at)
    }
  }
  return { restored, inserted }
}

// The chunk-file wrap of earlier ingests: split on sentence, semicolon and dash
// separators and rejoin with one space. Used only to write a legacy chunk file.
function earlierWrap(text: string) {
  return text.split("\n").flatMap((line) => {
    if (line.length <= READ_LINE_CAP) return [line]
    const out: string[] = []
    let current = ""
    for (const piece of line.split(/(?<=\.)\s+|(?<=;)\s+|\s+—\s+/).filter(Boolean)) {
      if (piece.length > READ_LINE_CAP) {
        if (current) out.push(current)
        current = ""
        for (let i = 0; i < piece.length; i += READ_LINE_CAP) out.push(piece.slice(i, i + READ_LINE_CAP))
      } else if (current && current.length + 1 + piece.length > READ_LINE_CAP) {
        out.push(current)
        current = piece
      } else current = current ? `${current} ${piece}` : piece
    }
    if (current) out.push(current)
    return out
  }).join("\n")
}

function spaceFreeRunAt(text: string, at: number) {
  const start = Math.max(text.lastIndexOf(" ", at - 1), text.lastIndexOf("\n", at - 1)) + 1
  const ends = [text.indexOf(" ", at), text.indexOf("\n", at)].filter((end) => end !== -1)
  return text.slice(start, ends.length ? Math.min(...ends) : text.length)
}

describe("source artifacts", () => {
  test("writes compact source maps and refreshes stale generated files", async () => {
    const dir = mkdtempSync(path.join(tmpdir(), "dochaus-source-artifacts-"))
    try {
      const text = "ARTICLE 1\nAlpha obligations.\n\nARTICLE 2\nBeta obligations."
      const db = openDb(dir)
      const docPath = path.join(dir, "Agreement.docx")
      const documentId = upsertDocument(db, docPath, "Agreement.docx", Date.now(), null)
      insertChunk(db, {
        documentId,
        docPath,
        docName: "Agreement.docx",
        section: "1",
        structuralPath: "Agreement > Article 1",
        party: "",
        chunkIndex: 0,
        text,
        charStart: 0,
        charEnd: text.length,
        embedding: new Float32Array([1, 2, 3]),
        flagged: false,
      })
      db.close()

      const first = await writeSourceArtifacts(dir)
      expect(first.documents).toBe(1)
      expect(first.chunks).toBeGreaterThan(0)
      expect(existsSync(path.join(dir, "source-index.json"))).toBe(true)
      expect(existsSync(path.join(dir, "source-map.md"))).toBe(true)
      const sourceMap = await Bun.file(path.join(dir, "source-map.json")).json() as {
        documents: Array<{ chunks: Array<{ text_path: string }> }>
      }
      expect(existsSync(path.join(dir, sourceMap.documents[0]!.chunks[0]!.text_path))).toBe(true)
      expect((await Bun.file(path.join(dir, "source-map.md")).text()).includes("Agreement.docx")).toBe(true)
      const sourceIndex = await Bun.file(path.join(dir, "source-index.json")).json() as {
        sources: Array<{ document: string }>
      }
      expect(sourceIndex.sources[0]!.document).toBe("Agreement.docx")
      const identifiersJson = await Bun.file(path.join(dir, "matter-identifiers.json")).text()
      const identifiersMarkdown = await Bun.file(path.join(dir, "matter-identifiers.md")).text()

      // #96: ingest stamps PERMANENT lineage rows.
      const ledger = path.join(dir, ".dochaus", "lineage.jsonl")
      expect(existsSync(ledger)).toBe(true)
      const rows = (await Bun.file(ledger).text()).trim().split("\n").map((l) => JSON.parse(l) as { tier: string; artifact: string; op: string; subagent: string })
      expect(rows.length).toBeGreaterThan(0)
      expect(rows.every((r) => r.tier === "PERMANENT" && r.op === "ingest" && r.subagent === "ingest")).toBe(true)
      expect(rows.some((r) => r.artifact === ".dochaus/legal.db")).toBe(true)

      await Bun.write(path.join(dir, "source-chunks", "stale.txt"), "stale")
      await Bun.sleep(5)
      await writeSourceArtifacts(dir)
      expect(existsSync(path.join(dir, "source-chunks", "stale.txt"))).toBe(false)
      expect(await Bun.file(path.join(dir, "matter-identifiers.json")).text()).toBe(identifiersJson)
      expect(await Bun.file(path.join(dir, "matter-identifiers.md")).text()).toBe(identifiersMarkdown)
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  })

  test("refreshes identifier metadata when the derived projection changes", async () => {
    const dir = mkdtempSync(path.join(tmpdir(), "dochaus-source-artifacts-"))
    try {
      const db = openDb(dir)
      const firstPath = path.join(dir, "First.docx")
      const firstText = "Morgan Ellis signed the first notice."
      const firstId = upsertDocument(db, firstPath, "First.docx", Date.now(), null)
      insertChunk(db, {
        documentId: firstId,
        docPath: firstPath,
        docName: "First.docx",
        section: "1",
        structuralPath: "First",
        party: "",
        chunkIndex: 0,
        text: firstText,
        charStart: 0,
        charEnd: firstText.length,
        embedding: new Float32Array([1, 2, 3]),
        flagged: false,
      })
      db.close()

      await writeSourceArtifacts(dir)
      const before = await Bun.file(path.join(dir, "matter-identifiers.json")).json() as { generated_at: string; documents: number }
      await Bun.sleep(5)

      const updated = openDb(dir)
      const secondPath = path.join(dir, "Second.docx")
      const secondText = "Taylor Morgan signed the second notice."
      const secondId = upsertDocument(updated, secondPath, "Second.docx", Date.now(), null)
      insertChunk(updated, {
        documentId: secondId,
        docPath: secondPath,
        docName: "Second.docx",
        section: "1",
        structuralPath: "Second",
        party: "",
        chunkIndex: 0,
        text: secondText,
        charStart: 0,
        charEnd: secondText.length,
        embedding: new Float32Array([1, 2, 3]),
        flagged: false,
      })
      updated.close()

      await writeSourceArtifacts(dir)
      const after = await Bun.file(path.join(dir, "matter-identifiers.json")).json() as { generated_at: string; documents: number }
      expect(after.documents).toBe(2)
      expect(after.generated_at).not.toBe(before.generated_at)
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  })

  test("writes matter-identifiers.json and matter-identifiers.md at ingest", async () => {
    const dir = mkdtempSync(path.join(tmpdir(), "dochaus-source-artifacts-"))
    try {
      // Authored synthetic procurement record; no real matter or personal data.
      const text = [
        "Procurement review meeting on March 17, 2026.",
        "Synthetic Case No. 00-TEST-0000001 was filed by Example Procurement Services.",
        "Coordinator: Taylor R. Morgan. Reviewer: Alex Morgan Bennett.",
        "Authority: 15 U.S.C. § 7001.",
      ].join("\n")
      const db = openDb(dir)
      const docPath = path.join(dir, "Synthetic Procurement.docx")
      const documentId = upsertDocument(db, docPath, "Synthetic Procurement.docx", Date.now(), null)
      insertChunk(db, {
        documentId,
        docPath,
        docName: "Synthetic Procurement.docx",
        section: "1",
        structuralPath: "Procurement",
        party: "",
        chunkIndex: 0,
        text,
        charStart: 0,
        charEnd: text.length,
        embedding: new Float32Array([1, 2, 3]),
        flagged: false,
      })
      db.close()

      await writeSourceArtifacts(dir)

      const identifiersJsonPath = path.join(dir, "matter-identifiers.json")
      const identifiersMdPath = path.join(dir, "matter-identifiers.md")
      expect(existsSync(identifiersJsonPath)).toBe(true)
      expect(existsSync(identifiersMdPath)).toBe(true)

      const parsed = await Bun.file(identifiersJsonPath).json()
      expect(parsed.schema_version).toBe(identifierExtraction.IDENTIFIERS_SCHEMA_VERSION)
      expect(parsed.rows.length).toBeGreaterThan(0)
      const values = parsed.rows.map((row: { value: string }) => row.value)
      expect(values).toEqual(expect.arrayContaining([
        expect.stringMatching(/00-TEST-0000001/),
        expect.stringMatching(/15 U\.S\.C\. § 7001/),
      ]))

      const md = await Bun.file(identifiersMdPath).text()
      expect(md).toContain("# Matter Identifiers")
      expect(md).toContain("case_reference")
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  })

  test("stamps a legal.db PERMANENT lineage row even on a zero-document ingest", async () => {
    const dir = mkdtempSync(path.join(tmpdir(), "dochaus-source-artifacts-"))
    try {
      openDb(dir).close()
      const result = await writeSourceArtifacts(dir)
      expect(result.documents).toBe(0)
      const ledger = path.join(dir, ".dochaus", "lineage.jsonl")
      expect(existsSync(ledger)).toBe(true)
      const rows = (await Bun.file(ledger).text()).trim().split("\n").map((l) => JSON.parse(l) as { tier: string; artifact: string })
      expect(rows.some((r) => r.tier === "PERMANENT" && r.artifact === ".dochaus/legal.db")).toBe(true)
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  })

  test("wraps long chunk lines at spaces without losing or changing a character", async () => {
    const dir = mkdtempSync(path.join(tmpdir(), "dochaus-source-artifacts-"))
    try {
      // Authored synthetic clause text; no real matter data.
      const clause = "the Buyer — subject to Section 9.2 — against Losses"
      const filler = (count: number) =>
        Array.from({ length: count }, (_, i) => ["The", "Seller", "shall", "indemnify;", "notice.", "within", "thirty", "days"][i % 8]).join(" ")
      const prose = `${filler(160)} ${clause} — and  each\tparty — ${filler(170)} ${clause}.`
      const longToken = `See ${"x".repeat(2300)} for details.`
      // Astral characters are two UTF-16 units; a cut between the halves would
      // write a replacement character instead of the original.
      const astralToken = `Marks a${"\u{1D538}".repeat(800)} end.`
      const text = [prose, longToken, astralToken].join("\n\n")
      expect(prose.length).toBeGreaterThan(2 * READ_LINE_CAP)

      const db = openDb(dir)
      const docPath = path.join(dir, "Wrapped.docx")
      const documentId = upsertDocument(db, docPath, "Wrapped.docx", Date.now(), null)
      insertChunk(db, {
        documentId,
        docPath,
        docName: "Wrapped.docx",
        section: "1",
        structuralPath: "Wrapped",
        party: "",
        chunkIndex: 0,
        text,
        charStart: 0,
        charEnd: text.length,
        embedding: new Float32Array([1, 2, 3]),
        flagged: false,
      })
      db.close()

      await writeSourceArtifacts(dir)
      const sourceMap = await Bun.file(path.join(dir, "source-map.json")).json() as {
        documents: Array<{ chunks: Array<{ text_path: string; start_char: number; end_char: number }> }>
      }
      const chunks = sourceMap.documents[0]!.chunks
      const files = await Promise.all(chunks.map((chunk) => Bun.file(path.join(dir, chunk.text_path)).text()))
      const wrappedLines = files.flatMap((file) => file.split("\n"))

      expect(wrappedLines.every((line) => line.length <= READ_LINE_CAP)).toBe(true)
      expect(files.join("\n").replace(/\s+/g, " ").split(clause).length - 1).toBe(2)
      chunks.forEach((chunk, index) => {
        const original = text.slice(chunk.start_char, chunk.end_char)
        const { restored, inserted } = unwrapChunk(original, files[index]!)
        expect(restored).toBe(original)
        for (const at of inserted) expect(spaceFreeRunAt(original, at).length).toBeGreaterThan(READ_LINE_CAP)
      })
      // Prose breaks sit near the cap (the last space that fits), not at every space.
      const proseLines = wrappedLines.filter((line) => line.includes("indemnify;"))
      expect(proseLines.length).toBeGreaterThan(1)
      expect(proseLines.slice(0, -1).every((line) => line.length > READ_LINE_CAP - 12)).toBe(true)
      // The over-long tokens are cut, nothing else: the short words around them keep their own lines intact.
      expect(wrappedLines).toEqual(expect.arrayContaining(["See", "Marks"]))
      expect(files.join("")).not.toContain("\uFFFD")
      expect(sourceArtifactsMatchRegistry(dir)).toBe(true)
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  })

  test("reads identifier values from the unwrapped source and points them at chunk file lines", async () => {
    const dir = mkdtempSync(path.join(tmpdir(), "dochaus-source-artifacts-"))
    try {
      // Authored synthetic text; the names are invented. Each wrapped line ends
      // exactly at the cap so the space after it becomes the line break.
      const pad = (length: number) => "the parties agree to these terms ".repeat(40).slice(0, length)
      const fill = (head: string, tail: string) => `${head}${pad(READ_LINE_CAP - head.length - tail.length)}${tail}`
      const first = fill("", " signed by Morgan")
      const second = fill("Ellis ", " for Example Procurement")
      const prose = `${first} ${second} Services today. ${pad(200)}.`
      const cut = `${"y".repeat(2 * READ_LINE_CAP)} Taylor Morgan signed.`
      const text = `${prose}\n\n${cut}`

      const db = openDb(dir)
      const docPath = path.join(dir, "Names.docx")
      const documentId = upsertDocument(db, docPath, "Names.docx", Date.now(), null)
      insertChunk(db, {
        documentId,
        docPath,
        docName: "Names.docx",
        section: "1",
        structuralPath: "Names",
        party: "",
        chunkIndex: 0,
        text,
        charStart: 0,
        charEnd: text.length,
        embedding: new Float32Array([1, 2, 3]),
        flagged: false,
      })
      db.close()

      await writeSourceArtifacts(dir)
      const sourceMap = await Bun.file(path.join(dir, "source-map.json")).json() as {
        documents: Array<{ chunks: Array<{ id: string; text_path: string }> }>
      }
      const fileLines = new Map(await Promise.all(sourceMap.documents[0]!.chunks.map(async (chunk) =>
        [chunk.id, (await Bun.file(path.join(dir, chunk.text_path)).text()).split("\n")] as const)))
      const { rows } = await Bun.file(path.join(dir, "matter-identifiers.json")).json() as {
        rows: Array<{ type: string; value: string; chunk_id: string; line: number; context: string }>
      }
      // The line a reader opens for a row starts with its context and holds the start of its value.
      const opened = (row: (typeof rows)[number]) => fileLines.get(row.chunk_id)![row.line - 1]
      const people = rows.filter((row) => row.type === "person" && /Ellis|Taylor/.test(row.value))
      expect(people.map((row) => [row.value, opened(row)])).toEqual([
        ["Morgan Ellis", first],
        ["Taylor Morgan", "Taylor Morgan signed."],
      ])
      expect(people.every((row) => opened(row)!.startsWith(row.context.replace(/\u2026$/, "")))).toBe(true)
      const organizations = rows.filter((row) => row.type === "organization")
      expect(organizations.some((row) => row.value.includes("\n"))).toBe(false)
      expect(organizations.filter((row) => row.value.endsWith("Example Procurement Services")).map(opened)).toEqual([second])
      expect(sourceArtifactsMatchRegistry(dir)).toBe(true)
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  })

  test("the startup check reads identifiers stored with an unchanged source map instead of extracting them again", async () => {
    const dir = mkdtempSync(path.join(tmpdir(), "dochaus-source-artifacts-"))
    try {
      // Authored synthetic text; the names are invented.
      const text = "ARTICLE 1\nMorgan Ellis signs for Example Procurement Services.\n\nARTICLE 2\nTaylor Morgan countersigns."
      const db = openDb(dir)
      const docPath = path.join(dir, "Names.docx")
      insertChunk(db, {
        documentId: upsertDocument(db, docPath, "Names.docx", Date.now(), null),
        docPath,
        docName: "Names.docx",
        section: "1",
        structuralPath: "Names",
        party: "",
        chunkIndex: 0,
        text,
        charStart: 0,
        charEnd: text.length,
        embedding: new Float32Array([1, 2, 3]),
        flagged: false,
      })
      db.close()
      await writeSourceArtifacts(dir)
      const markdown = path.join(dir, "matter-identifiers.md")
      const stored = readFileSync(markdown, "utf8")

      const extract = spyOn(identifierExtraction, "extractIdentifiers")
      try {
        expect(sourceArtifactsMatchRegistry(dir)).toBe(true)
        expect(await healSourceArtifacts(dir, false)).toBe("current")
        expect(extract).not.toHaveBeenCalled()
        // Markdown that no longer renders the stored identifiers is stale.
        writeFileSync(markdown, `${stored}Edited.\n`)
        expect(sourceArtifactsMatchRegistry(dir)).toBe(false)
        writeFileSync(markdown, stored)
        // A changed source no longer has the source map's sha256.
        const changed = openDb(dir)
        changed.run("UPDATE chunks SET text = replace(text, 'countersigns', 'also signs')")
        changed.close()
        expect(sourceArtifactsMatchRegistry(dir)).toBe(false)
      } finally {
        extract.mockRestore()
      }
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  })

  test("identifiers an earlier extraction version stored are extracted again at one start, then match", async () => {
    const dir = mkdtempSync(path.join(tmpdir(), "dochaus-source-artifacts-"))
    try {
      // Authored synthetic text; the names are invented.
      const text = "ARTICLE 1\nMorgan Ellis signs for Example Procurement Services."
      const db = openDb(dir)
      const docPath = path.join(dir, "Names.docx")
      insertChunk(db, {
        documentId: upsertDocument(db, docPath, "Names.docx", Date.now(), null),
        docPath,
        docName: "Names.docx",
        section: "1",
        structuralPath: "Names",
        party: "",
        chunkIndex: 0,
        text,
        charStart: 0,
        charEnd: text.length,
        embedding: new Float32Array([1, 2, 3]),
        flagged: false,
      })
      db.close()
      await writeSourceArtifacts(dir)
      // The file as version 1 (values read from the wrapped chunk files) stored it.
      const file = path.join(dir, "matter-identifiers.json")
      writeFileSync(file, `${JSON.stringify({ ...JSON.parse(readFileSync(file, "utf8")), schema_version: 1 }, null, 2)}\n`)

      const extract = spyOn(identifierExtraction, "extractIdentifiers")
      try {
        expect(await healSourceArtifacts(dir, false)).toBe("rebuilt")
        expect(extract).toHaveBeenCalledTimes(1)
        expect(JSON.parse(readFileSync(file, "utf8")).schema_version).toBe(identifierExtraction.IDENTIFIERS_SCHEMA_VERSION)
        expect(await healSourceArtifacts(dir, false)).toBe("current")
        expect(extract).toHaveBeenCalledTimes(1)
      } finally {
        extract.mockRestore()
      }
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  })

  test("the startup heal rebuilds earlier-wrap chunk files and stamps only what changed", async () => {
    const dir = mkdtempSync(path.join(tmpdir(), "dochaus-source-artifacts-"))
    try {
      // Authored synthetic texts; no real matter data.
      const clause = "the Buyer — subject to Section 9.2 — against Losses"
      const words = Array.from({ length: 400 }, (_, i) => ["The", "Seller", "shall", "indemnify;", "notice.", "within"][i % 6]).join(" ")
      const documents = { "Long.docx": `${words} ${clause} ${words}`, "Short.docx": "ARTICLE 1\nShort obligations." }
      const db = openDb(dir)
      for (const [name, text] of Object.entries(documents)) {
        const docPath = path.join(dir, name)
        insertChunk(db, {
          documentId: upsertDocument(db, docPath, name, Date.now(), null),
          docPath,
          docName: name,
          section: "1",
          structuralPath: name,
          party: "",
          chunkIndex: 0,
          text,
          charStart: 0,
          charEnd: text.length,
          embedding: new Float32Array([1, 2, 3]),
          flagged: false,
        })
      }
      db.close()
      await writeSourceArtifacts(dir)
      const sourceMapBytes = readFileSync(path.join(dir, "source-map.json"))
      const sourceMap = JSON.parse(sourceMapBytes.toString("utf8")) as {
        documents: Array<{ document: string; chunks: Array<{ text_path: string; start_char: number; end_char: number }> }>
      }
      const longChunks = sourceMap.documents.find((document) => document.document === "Long.docx")!.chunks
      const currentFormat = new Map(longChunks.map((chunk) => [chunk.text_path, readFileSync(path.join(dir, chunk.text_path), "utf8")]))

      // An earlier ingest left (and stamped) the long document's chunk files in the earlier wrap.
      const earlier = new Map(longChunks.map((chunk) =>
        [chunk.text_path, earlierWrap(documents["Long.docx"].slice(chunk.start_char, chunk.end_char))]))
      for (const [file, text] of earlier) writeFileSync(path.join(dir, file), text)
      const rewrapped = [...earlier].filter(([file, text]) => text !== currentFormat.get(file)).map(([file]) => file)
      expect(rewrapped.length).toBeGreaterThan(0)
      stampIngestLineage(dir, Object.keys(documents), "earlier")
      expect(sourceArtifactsMatchRegistry(dir)).toBe(false)

      const lineage = () => readFileSync(path.join(dir, ".dochaus", "lineage.jsonl"), "utf8").trim().split("\n").map((row) => JSON.parse(row))
      const rebuilt = () => [
        "source-index.json", "source-map.json", "source-map.md", "source-ledger.json", "matter-identifiers.json", "matter-identifiers.md",
        ...[...new Bun.Glob("{source-chunks,sources}/*.txt").scanSync({ cwd: dir })].map((file) => file.replace(/\\/g, "/")),
      ].sort()
      const sha = (file: string) => createHash("sha256").update(readFileSync(path.join(dir, file))).digest("hex")
      const before = new Map(rebuilt().map((file) => [file, sha(file)]))
      const rowsBefore = lineage().length

      // The startup heal (ingest server boot) with no rolled-back mutation.
      expect(await healSourceArtifacts(dir, false)).toBe("rebuilt")

      expect(sourceArtifactsMatchRegistry(dir)).toBe(true)
      for (const [file, text] of currentFormat) expect(readFileSync(path.join(dir, file), "utf8")).toBe(text)
      expect(readFileSync(path.join(dir, "source-map.json")).equals(sourceMapBytes)).toBe(true)
      expect(rebuilt()).toEqual([...before.keys()])
      const changed = rebuilt().filter((file) => sha(file) !== before.get(file))
      expect(changed).toEqual(rewrapped)
      const added = lineage().slice(rowsBefore)
      expect(added.map((row) => row.artifact).sort()).toEqual(changed)
      for (const row of added) {
        expect(row).toMatchObject({ op: "write", content_sha256: sha(row.artifact), provenance: { origin: "source-rebuild" } })
      }
      // A matching projection is left alone.
      expect(await healSourceArtifacts(dir, false)).toBe("current")
      expect(lineage().length).toBe(rowsBefore + added.length)
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  })

  test("the startup heal after a rolled-back mutation stamps every artifact as an ingest, legal.db included", async () => {
    const dir = mkdtempSync(path.join(tmpdir(), "dochaus-source-artifacts-"))
    try {
      // Authored synthetic texts; no real matter data.
      const addDocument = (name: string, text: string) => {
        const db = openDb(dir)
        const docPath = path.join(dir, name)
        insertChunk(db, {
          documentId: upsertDocument(db, docPath, name, Date.now(), null),
          docPath,
          docName: name,
          section: "1",
          structuralPath: name,
          party: "",
          chunkIndex: 0,
          text,
          charStart: 0,
          charEnd: text.length,
          embedding: new Float32Array([1, 2, 3]),
          flagged: false,
        })
        db.close()
      }
      addDocument("Agreement.docx", "ARTICLE 1\nAlpha obligations.")
      await writeSourceArtifacts(dir)
      // Recovery restored a registry (and so a legal.db) that lineage has not recorded.
      addDocument("Schedule.docx", "SCHEDULE 1\nBeta obligations.")

      const lineage = () => readFileSync(path.join(dir, ".dochaus", "lineage.jsonl"), "utf8").trim().split("\n").map((row) => JSON.parse(row))
      const sha = (file: string) => createHash("sha256").update(readFileSync(path.join(dir, file))).digest("hex")
      const latest = (artifact: string) => lineage().findLast((row) => row.artifact === artifact)
      expect(latest(".dochaus/legal.db").content_sha256).not.toBe(sha(".dochaus/legal.db"))
      const rowsBefore = lineage().length

      expect(await healSourceArtifacts(dir, true)).toBe("recovered")

      const added = lineage().slice(rowsBefore)
      expect(added.find((row) => row.artifact === ".dochaus/legal.db")).toMatchObject({
        op: "ingest",
        content_sha256: sha(".dochaus/legal.db"),
        provenance: { origin: "upload" },
      })
      // Every artifact on disk ends with a lineage row that matches it.
      const tracked = [
        ".dochaus/legal.db", "source-index.json", "source-map.json", "source-map.md", "source-ledger.json", "matter-identifiers.json", "matter-identifiers.md",
        ...[...new Bun.Glob("{source-chunks,sources}/*.txt").scanSync({ cwd: dir })].map((file) => file.replace(/\\/g, "/")),
      ]
      expect(tracked.filter((artifact) => latest(artifact)?.content_sha256 !== sha(artifact))).toEqual([])
      expect(sourceArtifactsMatchRegistry(dir)).toBe(true)
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  })
})
