import { expect, test } from "bun:test"
import { createHash } from "node:crypto"
import { mkdtempSync, readFileSync, realpathSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { Document, Packer, Paragraph } from "docx"
import { liveText } from "../../../dochaus/lib/extract"
import { verifyCitation } from "../../../dochaus/plugin/legal/citations"
import { ftsInSync, getMeta, openDb } from "./db"
import { DIM, MODEL } from "./embed"
import type { MatterIdentifiers } from "./identifiers"
import { extractDocumentText, ingestDocument } from "./ingest"
import { sourceArtifactsMatchRegistry } from "./source-artifacts"
import { reconstructText } from "./structure"

type Fixture = { name: string; bytes: Buffer; marker: string; email: string }
type StoredChunk = {
  id: number
  text: string
  char_start: number
  char_end: number
  section_ordinal: number
  embedding: Uint8Array
}
type SourceMap = {
  documents: {
    document: string
    extracted_text: string
    characters: number
    sha256: string
    chunks: { id: string; text_path: string; start_char: number; end_char: number; sha256: string }[]
  }[]
}
const sha256 = (text: string) => createHash("sha256").update(text).digest("hex")

// All people, addresses and records are invented. Different sentinels prove
// each committed document owns its retrieval and evidence coordinates. The
// Markdown line exceeds the source-file soft-wrap limit so identifier pointers
// must address the reader's actual file, rather than the unwrapped index text.
async function fixtures(): Promise<Fixture[]> {
  return [
    {
      name: "Delivery notice.docx",
      bytes: Buffer.from(
        await Packer.toBuffer(
          new Document({
            sections: [
              {
                children: [
                  new Paragraph("1. Collection"),
                  new Paragraph("Rowan Vale approved KESTRELDOCX527 on 14 February 2025."),
                  new Paragraph("Contact rowan.vale@kestrel.example for the original receipt."),
                  new Paragraph("2. Arrival"),
                  new Paragraph("The crates arrived at Moor Quay on 16 February 2025."),
                ],
              },
            ],
          }),
        ),
      ),
      marker: "KESTRELDOCX527",
      email: "rowan.vale@kestrel.example",
    },
    {
      name: "Collection correspondence.eml",
      bytes: Buffer.from(
        [
          'From: "Seren Moss" <seren.moss@oriole.example>',
          "To: records@kestrel.example",
          "Date: Mon, 17 Feb 2025 09:30:00 +0000",
          "Subject: Collection confirmation",
          "Content-Type: text/plain; charset=utf-8",
          "",
          "Seren Moss confirmed ORIOLEMAIL638 at the harbour office.",
          "The original receipt will follow separately.",
          "",
        ].join("\r\n"),
      ),
      marker: "ORIOLEMAIL638",
      email: "seren.moss@oriole.example",
    },
    {
      name: "Inventory record.md",
      bytes: Buffer.from(
        [
          "# Inventory record",
          "",
          "Avery Brook recorded LINNETMARKDOWN816 on 18 February 2025.",
          "",
          `${"The invented inventory was checked against the packing list. ".repeat(24)}Contact avery.brook@linnet.example for the signed record.`,
          "",
          "Receipt was confirmed at the depot.",
          "",
          "Receipt was confirmed at the depot.",
          "",
        ].join("\r\n"),
      ),
      marker: "LINNETMARKDOWN816",
      email: "avery.brook@linnet.example",
    },
    {
      name: "Receipt note.pdf",
      bytes: textPdf([
        "Devon Reed recorded PLOVERPDF749 on 19 February 2025.",
        "Contact devon.reed@plover.example for the receipt.",
      ]),
      marker: "PLOVERPDF749",
      email: "devon.reed@plover.example",
    },
  ]
}

// A normal one-page text PDF, following the existing upload-test convention;
// no renderer, OCR service, remote provider or new dependency is required.
function textPdf(lines: string[]) {
  const stream = `BT /F1 12 Tf 40 700 Td ${lines.map((line, i) => `${i ? "0 -20 Td " : ""}(${line}) Tj`).join(" ")} ET`
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

test("new mixed-format uploads commit one consistent canonical source across retrieval and evidence indexes", async () => {
  const directory = realpathSync.native(mkdtempSync(path.join(tmpdir(), "dochaus-source-index-consistency-")))
  try {
    const documents = await fixtures()
    for (const fixture of documents) {
      // Use the actual extraction/embedding/commit path, as ingest-commit tests
      // do. Embeddings use the local product model; no expected vectors or
      // fake chunks are seeded, and no paid generation connection is used.
      const uploaded = await ingestDocument(directory, fixture.name, fixture.bytes)
      expect(uploaded.quality.status).toBe("ok")
      expect(uploaded.chunks).toBeGreaterThan(0)
      expect(sourceArtifactsMatchRegistry(directory)).toBe(true)
    }
    const map = JSON.parse(readFileSync(path.join(directory, "source-map.json"), "utf8")) as SourceMap
    const identifiers = JSON.parse(
      readFileSync(path.join(directory, "matter-identifiers.json"), "utf8"),
    ) as MatterIdentifiers
    expect(map.documents.map((source) => source.document).toSorted()).toEqual(documents.map((d) => d.name).toSorted())
    expect(identifiers.sources.map((source) => source.document).toSorted()).toEqual(
      documents.map((d) => d.name).toSorted(),
    )
    const chunkFiles = new Map(
      map.documents.flatMap((source) =>
        source.chunks.map(
          (chunk) =>
            [
              chunk.id,
              { source: source.document, text: readFileSync(path.join(directory, chunk.text_path), "utf8") },
            ] as const,
        ),
      ),
    )
    const missingEmails: { source: string; expected: string }[] = []
    const db = openDb(directory)
    try {
      expect(getMeta(db, "embedding_model")).toBe(MODEL)
      expect(ftsInSync(db)).toBe(true)
      for (const fixture of documents) {
        const docPath = path.join(directory, fixture.name)
        const canonical = reconstructText(db, docPath)
        expect(canonical).toBe(`${await extractDocumentText(fixture.name, fixture.bytes)}\n`)
        expect(await liveText(docPath)).toBe(canonical)
        const row = db.query("SELECT id, role, rev FROM documents WHERE name = ?").get(fixture.name) as {
          id: number
          role: string
          rev: string
        }
        expect(row.role).toBe("source")
        expect(row.rev).toBe(sha256(canonical).slice(0, 12))
        const sections = db
          .query("SELECT ordinal, char_start, char_end FROM sections WHERE doc_path = ? ORDER BY ordinal")
          .all(docPath) as { ordinal: number; char_start: number; char_end: number }[]
        expect(sections.length).toBeGreaterThan(0)
        expect(sections[0]!.char_start).toBe(0)
        expect(sections.at(-1)!.char_end).toBe(canonical.length)
        for (let index = 1; index < sections.length; index++)
          expect(sections[index]!.char_start).toBe(sections[index - 1]!.char_end)
        const chunks = db
          .query(
            "SELECT id, text, char_start, char_end, section_ordinal, embedding FROM chunks WHERE document_id = ? ORDER BY char_start, id",
          )
          .all(row.id) as StoredChunk[]
        for (const chunk of chunks) {
          expect(canonical.slice(chunk.char_start, chunk.char_end)).toBe(chunk.text)
          const section = sections.find((item) => item.ordinal === chunk.section_ordinal)
          expect(section).toBeDefined()
          expect(chunk.char_start).toBeGreaterThanOrEqual(section!.char_start)
          expect(chunk.char_start).toBeLessThan(section!.char_end)
          const vector = new Float32Array(Uint8Array.from(chunk.embedding).buffer)
          expect(vector.length).toBe(DIM)
          expect([...vector].every(Number.isFinite)).toBe(true)
          expect([...vector].some((value) => value !== 0)).toBe(true)
        }
        const hits = db
          .query(
            "SELECT DISTINCT chunks.doc_name FROM chunks_fts JOIN chunks ON chunks.id = chunks_fts.rowid WHERE chunks_fts MATCH ?",
          )
          .all(fixture.marker) as { doc_name: string }[]
        expect(hits.map((hit) => hit.doc_name)).toEqual([fixture.name])
        const source = map.documents.find((entry) => entry.document === fixture.name)!
        const trimmed = canonical.trim()
        expect(source.characters).toBe(trimmed.length)
        expect(source.sha256).toBe(sha256(trimmed))
        expect(readFileSync(path.join(directory, source.extracted_text), "utf8")).toBe(`${trimmed}\n`)
        for (const chunk of source.chunks) {
          const excerpt = trimmed.slice(chunk.start_char, chunk.end_char)
          expect(chunk.sha256).toBe(sha256(excerpt))
          expect(chunkFiles.get(chunk.id)!.text.replace(/\s+/g, " ").trim()).toBe(excerpt.replace(/\s+/g, " ").trim())
        }
        // The live verifier must keep this exact occurrence, with no nearest
        // re-anchor masking a canonical-coordinate mismatch.
        const charStart = canonical.indexOf(fixture.marker)
        const citation = await verifyCitation({
          documentName: fixture.name,
          docPath,
          section: "Uploaded record",
          excerpt: fixture.marker,
          charStart,
          charEnd: charStart + fixture.marker.length,
        })
        expect(citation?.charStart).toBe(charStart)
        expect(citation?.reanchored).toBeUndefined()
        if (
          !identifiers.rows.some(
            (identifier) =>
              identifier.source === fixture.name && identifier.type === "email" && identifier.value === fixture.email,
          )
        )
          missingEmails.push({ source: fixture.name, expected: fixture.email })
      }
    } finally {
      db.close()
    }
    for (const identifier of identifiers.rows) {
      const chunk = chunkFiles.get(identifier.chunk_id)
      expect(chunk?.source).toBe(identifier.source)
      const line = chunk!.text.split("\n")[identifier.line - 1]
      expect(line).toBeDefined()
      expect(line!.trimStart().startsWith(identifier.context.replace(/…$/, ""))).toBe(true)
    }
    expect(missingEmails).toEqual([])
  } finally {
    rmSync(directory, { recursive: true, force: true })
  }
}, 120_000)
