import { describe, expect, spyOn, test } from "bun:test"
import * as identifiers from "./identifiers"
import { extractIdentifiers, fileOffset, IDENTIFIERS_SCHEMA_VERSION, renderIdentifiersMarkdown, type IdentifierChunk, type MatterIdentifiers } from "./identifiers"

// Authored synthetic procurement examples. Names, contact details and case IDs
// are invented for extraction tests; no client or court document was used.
const sampleChunk: IdentifierChunk = {
  id: "synthetic-procurement-memo-c007",
  document: "synthetic-procurement-memo",
  text: [
    "C. Procurement Review",
    "Alex Morgan Bennett joined Example Research Institute in 2023.",
    "Example Procurement Services filed Case No. 00-TEST-0000001.",
    "Taylor R. Morgan coordinates a fictional project in Baltimore, Maryland.",
    "Monthly budget is $437.00; contact (410) 555-0113 or coordinator@example.com.",
    "Authority: 15 U.S.C. § 7001 and MD Com Law § 1-101.",
    "Review meeting scheduled for August 3, 2023.",
  ].join("\n"),
}

function rowsOfType(data: MatterIdentifiers, type: string) {
  return data.rows.filter((row) => row.type === type).map((row) => row.value)
}

describe("extractIdentifiers", () => {
  test("schema and metadata are populated", async () => {
    const data = await extractIdentifiers([sampleChunk])
    expect(data.schema_version).toBe(IDENTIFIERS_SCHEMA_VERSION)
    expect(data.documents).toBe(1)
    expect(data.generated_at).toMatch(/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}/)
    expect(data.rows.length).toBeGreaterThan(0)
  })

  test("extracts person names", async () => {
    const data = await extractIdentifiers([sampleChunk])
    expect(rowsOfType(data, "person")).toEqual(expect.arrayContaining([
      "Alex Morgan Bennett",
      "Taylor R. Morgan",
    ]))
  })

  test("extracts organizations", async () => {
    const data = await extractIdentifiers([sampleChunk])
    expect(rowsOfType(data, "organization")).toEqual(
      expect.arrayContaining([
        "Example Procurement Services",
        "Example Research Institute",
      ]),
    )
  })

  test("extracts places", async () => {
    const data = await extractIdentifiers([sampleChunk])
    const places = rowsOfType(data, "place")
    expect(places.some((p) => p.includes("Baltimore"))).toBe(true)
    expect(places.some((p) => p.includes("Maryland"))).toBe(true)
  })

  test("extracts money", async () => {
    const data = await extractIdentifiers([sampleChunk])
    expect(rowsOfType(data, "money")).toContain("$437.00")
  })

  test("extracts dates", async () => {
    const data = await extractIdentifiers([sampleChunk])
    expect(rowsOfType(data, "date")).toEqual(
      expect.arrayContaining([
        expect.stringMatching(/August 3, 2023/),
      ]),
    )
  })

  test("extracts phone numbers", async () => {
    const data = await extractIdentifiers([sampleChunk])
    expect(rowsOfType(data, "phone")).toContain("(410) 555-0113")
  })

  test("extracts emails", async () => {
    const data = await extractIdentifiers([sampleChunk])
    expect(rowsOfType(data, "email")).toContain("coordinator@example.com")
  })

  test("extracts case references via custom regex", async () => {
    const data = await extractIdentifiers([sampleChunk])
    expect(rowsOfType(data, "case_reference")).toEqual(
      expect.arrayContaining(["00-TEST-0000001"]),
    )
  })

  test("extracts statute citations via custom regex", async () => {
    const data = await extractIdentifiers([sampleChunk])
    const statutes = rowsOfType(data, "statute_citation")
    expect(statutes).toEqual(
      expect.arrayContaining([
        expect.stringMatching(/15 U\.S\.C\. § 7001/),
        expect.stringMatching(/MD Com Law § 1-101/),
      ]),
    )
  })

  test("extracts state-law Title/Subtitle statute citations without § symbol", async () => {
    const chunk: IdentifierChunk = {
      id: "intake-c002",
      document: "intake-memo",
      text: [
        "The matter arises under Maryland Family Law Article, Title 5, Subtitle 3.",
        "Counsel also cites California Family Code, Title 4 and Estates and Trusts Article, Title 1, Subtitle 2.",
        "Harry Potter, Chapter 5 is not a statute citation.",
        "Title 5 of the contract governs delivery dates.",
      ].join("\n"),
    }
    const statutes = rowsOfType(await extractIdentifiers([chunk]), "statute_citation")
    expect(statutes).toEqual(
      expect.arrayContaining([
        "Maryland Family Law Article, Title 5, Subtitle 3",
        "California Family Code, Title 4",
        "Estates and Trusts Article, Title 1, Subtitle 2",
      ]),
    )
    // Negative checks: book titles and contract clauses must not match.
    expect(statutes.some((s) => /Harry Potter/.test(s))).toBe(false)
    expect(statutes.some((s) => /contract/.test(s))).toBe(false)
  })

  test("each row carries source, chunk_id, line, and context", async () => {
    const data = await extractIdentifiers([sampleChunk])
    for (const row of data.rows) {
      expect(row.source).toBe("synthetic-procurement-memo")
      expect(row.chunk_id).toBe("synthetic-procurement-memo-c007")
      expect(Number.isInteger(row.line)).toBe(true)
      expect(row.line).toBeGreaterThanOrEqual(1)
      expect(row.context.length).toBeGreaterThan(0)
      expect(row.context.length).toBeLessThanOrEqual(200)
    }
  })

  test("line numbers map to chunk-local line positions", async () => {
    const data = await extractIdentifiers([sampleChunk])
    const caseRow = data.rows.find((row) => row.type === "case_reference")
    expect(caseRow).toBeDefined()
    // "Example Procurement Services filed Case No. 00-TEST-0000001."
    // is line 3 of the chunk (1-indexed).
    expect(caseRow!.line).toBe(3)
  })

  test("context contains the matched value", async () => {
    const data = await extractIdentifiers([sampleChunk])
    const emailRow = data.rows.find((row) => row.type === "email")
    expect(emailRow!.context).toContain("coordinator@example.com")
  })

  test("multiple chunks attribute correctly", async () => {
    const secondChunk: IdentifierChunk = {
      id: "intake-c001",
      document: "intake-memo",
      text: "Synthetic Docket No. 00-TEST-0000002 was filed in 2024.",
    }
    const data = await extractIdentifiers([sampleChunk, secondChunk])
    const intakeRow = data.rows.find((row) => row.chunk_id === "intake-c001")
    expect(intakeRow).toBeDefined()
    expect(intakeRow!.source).toBe("intake-memo")
    expect(intakeRow!.value).toBe("00-TEST-0000002")
    expect(data.documents).toBe(2)
  })

  test("empty chunks produce empty inventory", async () => {
    const data = await extractIdentifiers([])
    expect(data.rows).toEqual([])
    expect(data.documents).toBe(0)
  })

  // Canonical-form normalization: when the same person appears with both an
  // initial-form middle name and a full-form middle name (same first + last
  // word), the inventory normalizes every variant row's Value column to the
  // longest form. The original surface form survives in the Context column
  // (which is the source text, never rewritten); only the Value is normalized
  // so downstream stages reading the Value column for atom pinning always
  // carry the full legal name. Conservative: only collapses when one form's
  // middle is a single initial (R., A.) that matches the first letter of the
  // other form's middle (Renee, Alex). Two different full middle names
  // (Alexis vs Alex) are NOT collapsed — could be different people.
  test("collapses initial-middle variants to the full-middle form", async () => {
    const variantChunk: IdentifierChunk = {
      id: "doc-c001",
      document: "doc",
      text: [
        "Taylor R. Morgan coordinates the fictional project.",
        "Counsel for Taylor Renee Morgan confirmed consent.",
      ].join("\n"),
    }
    const data = await extractIdentifiers([variantChunk])
    const personValues = rowsOfType(data, "person")
    // No row should still carry "Taylor R. Morgan" — all collapsed to "Taylor Renee Morgan"
    expect(personValues).not.toContain("Taylor R. Morgan")
    expect(personValues.filter((v) => v === "Taylor Renee Morgan").length).toBeGreaterThan(0)
  })

  test("does NOT collapse when neither middle is a single initial", async () => {
    // "Jordan Alex Lee" and "Jordan Alexis Lee" — both have multi-letter
    // middles that are NOT initials of each other. Could be different people
    // (Alex vs Alexis). Must remain distinct.
    const distinctChunk: IdentifierChunk = {
      id: "doc-c001",
      document: "doc",
      text: [
        "Jordan Alex Lee signed the agreement.",
        "Witness: Jordan Alexis Lee.",
      ].join("\n"),
    }
    const data = await extractIdentifiers([distinctChunk])
    const personValues = rowsOfType(data, "person")
    expect(personValues).toContain("Jordan Alex Lee")
    expect(personValues).toContain("Jordan Alexis Lee")
  })

  test("does NOT collapse two-word names with no middle", async () => {
    // "Sarah Chen" appears in many sources. There is no variant here, and no
    // middle to compare. Conservative: never collapse a two-word row.
    const twoWordChunk: IdentifierChunk = {
      id: "doc-c001",
      document: "doc",
      text: "Sarah Chen represented the petitioner. Sarah Chen is partner.",
    }
    const data = await extractIdentifiers([twoWordChunk])
    const personValues = rowsOfType(data, "person")
    expect(personValues).toContain("Sarah Chen")
  })

  test("preserves source diversity after collapse", async () => {
    // Same entity from two different sources — collapse to longest form but
    // keep both rows so the inventory still shows the entity in both sources.
    const multiSourceChunkA: IdentifierChunk = {
      id: "docA-c001",
      document: "docA.docx",
      text: "Taylor R. Morgan is the coordinator.",
    }
    const multiSourceChunkB: IdentifierChunk = {
      id: "docB-c001",
      document: "docB.docx",
      text: "I represent Taylor Renee Morgan in this matter.",
    }
    const data = await extractIdentifiers([multiSourceChunkA, multiSourceChunkB])
    const personRows = data.rows.filter((r) => r.type === "person" && r.value === "Taylor Renee Morgan")
    expect(personRows.length).toBeGreaterThanOrEqual(2)
    expect(personRows.some((r) => r.source === "docA.docx")).toBe(true)
    expect(personRows.some((r) => r.source === "docB.docx")).toBe(true)
  })

  test("original surface form survives in the Context column after collapse", async () => {
    // The Value is normalized but the Context column preserves the source
    // text verbatim — so "Taylor R. Morgan" still appears in the row whose
    // Value was normalized, just in the Context. This is the audit trail.
    const variantChunk: IdentifierChunk = {
      id: "doc-c001",
      document: "doc",
      text: "Taylor R. Morgan coordinates the fictional project.",
    }
    const fullChunk: IdentifierChunk = {
      id: "doc-c002",
      document: "doc",
      text: "Counsel for Taylor Renee Morgan confirmed consent.",
    }
    const data = await extractIdentifiers([variantChunk, fullChunk])
    const normalizedRow = data.rows.find(
      (r) => r.type === "person" && r.value === "Taylor Renee Morgan" && r.chunk_id === "doc-c001",
    )
    expect(normalizedRow).toBeDefined()
    expect(normalizedRow!.context).toContain("Taylor R. Morgan")
  })
})

// Authored synthetic documents; the names, case IDs and amounts are invented.
const agreement = ["ARTICLE 1\nCounsel for Taylor Renee Morgan confirmed consent.", "Example Procurement Services pays $437.00 by August 3, 2023."]
const notice = ["Taylor R. Morgan coordinates the fictional project.", "Example Procurement Services filed Case No. 00-TEST-0000001."]
const schedule = ["SCHEDULE 1\nMorgan Ellis of Example Research Institute approves the budget."]

function matterChunks(documents: Record<string, string[]>): IdentifierChunk[] {
  return Object.entries(documents).flatMap(([document, texts]) =>
    texts.map((text, index) => ({ id: `${document}-c${String(index + 1).padStart(3, "0")}`, document, text })))
}

// The file as writeMatterIdentifiers stores and readMatterIdentifiers reads it.
async function stored(documents: Record<string, string[]>) {
  return JSON.parse(JSON.stringify(await extractIdentifiers(matterChunks(documents)))) as MatterIdentifiers
}

function withoutTime(data: MatterIdentifiers) {
  return JSON.stringify({ ...data, generated_at: "" })
}

describe("extraction from the matter's previous identifiers", () => {
  test("gives exactly what a full extraction gives after an upload, an edit and a delete", async () => {
    const previous = await stored({ "Agreement.docx": agreement, "Notice.docx": notice })
    const states: Record<string, string[]>[] = [
      { "Agreement.docx": agreement, "Notice.docx": notice, "Schedule.docx": schedule },
      { "Agreement.docx": agreement, "Notice.docx": [notice[0]!, "Example Procurement Services filed Case No. 00-TEST-0000002."] },
      { "Notice.docx": notice },
      { "Schedule.docx": schedule, "Agreement.docx": agreement },
    ]
    for (const documents of states) {
      const chunks = matterChunks(documents)
      expect(withoutTime(await extractIdentifiers(chunks, previous))).toBe(withoutTime(await extractIdentifiers(chunks)))
    }
  })

  test("parses only the chunks of new or changed documents", async () => {
    const previous = await stored({ "Agreement.docx": agreement, "Notice.docx": notice })
    const scan = spyOn(identifiers, "scanChunk")
    const scanned = () => scan.mock.calls.map(([chunk]) => chunk.id)
    try {
      await extractIdentifiers(matterChunks({ "Agreement.docx": agreement, "Notice.docx": notice, "Schedule.docx": schedule }), previous)
      expect(scanned()).toEqual(["Schedule.docx-c001"])
      scan.mockClear()
      await extractIdentifiers(matterChunks({ "Notice.docx": notice }), previous)
      expect(scanned()).toEqual([])
      // The same text wrapped differently moves the rows' line and context, so
      // it is parsed again.
      const rewrapped = matterChunks({ "Agreement.docx": agreement, "Notice.docx": notice }).map((chunk) =>
        chunk.document === "Notice.docx" ? { ...chunk, file: { text: chunk.text.replace(" ", "\n"), inserted: [] } } : chunk)
      await extractIdentifiers(rewrapped, previous)
      expect(scanned()).toEqual(["Notice.docx-c001", "Notice.docx-c002"])
    } finally {
      scan.mockRestore()
    }
  })

  test("deleting the document with the full middle name returns the others' rows to their surface form", async () => {
    const previous = await stored({ "Agreement.docx": agreement, "Notice.docx": notice })
    const collapsed = previous.rows.find((row) => row.type === "person" && row.source === "Notice.docx")
    expect(collapsed).toMatchObject({ value: "Taylor Renee Morgan", surface: "Taylor R. Morgan" })
    const after = await extractIdentifiers(matterChunks({ "Notice.docx": notice }), previous)
    expect(after.rows.some((row) => row.source === "Agreement.docx")).toBe(false)
    const person = after.rows.find((row) => row.type === "person")
    expect(person?.value).toBe("Taylor R. Morgan")
    expect(person && "surface" in person).toBe(false)
  })

  test("a previous file whose rows are not described by its sources is not reused", async () => {
    const previous = await stored({ "Agreement.docx": agreement, "Notice.docx": notice })
    const chunks = matterChunks({ "Agreement.docx": agreement, "Notice.docx": notice })
    const full = withoutTime(await extractIdentifiers(chunks))
    const first = previous.rows[0]!
    const unusable = [
      { ...previous, sources: undefined },
      { ...previous, sources: [...previous.sources, previous.sources[0]] },
      { ...previous, rows: [...previous.rows, { ...first, source: "Elsewhere.docx" }] },
      { ...previous, rows: [{ ...first, line: String(first.line) }, ...previous.rows.slice(1)] },
      { ...previous, rows: [{ ...first, note: "edited" }, ...previous.rows.slice(1)] },
    ]
    const scan = spyOn(identifiers, "scanChunk")
    try {
      for (const file of unusable) {
        scan.mockClear()
        const data = await extractIdentifiers(chunks, file as unknown as MatterIdentifiers)
        expect(scan).toHaveBeenCalledTimes(chunks.length)
        expect(withoutTime(data)).toBe(full)
      }
    } finally {
      scan.mockRestore()
    }
  })

  test("matches a full extraction through chained uploads, edits, deletes, renames and re-wraps", async () => {
    // Each step extracts from the previous step's stored file, as uploads and
    // deletes do, over a fixed pseudo-random sequence that includes initial and
    // full middle-name variants of one person and ties between full middles.
    const sentences = [...agreement, ...notice, ...schedule,
      "Witness: Taylor Robin Morgan.", "Jordan A. Lee signed the agreement.", "Jordan Alexis Lee approved the budget.", "Jordan Alex Lee filed the notice."]
    const names = ["A.docx", "B.docx", "C.docx", "D.docx"]
    let seed = 7
    const pick = <T,>(items: T[]) => {
      seed = (seed * 1103515245 + 12345) & 0x7fffffff
      return items[seed % items.length]!
    }
    for (let run = 0; run < 6; run++) {
      const documents = new Map<string, { texts: string[]; wrapped: boolean }>()
      let previous: MatterIdentifiers | undefined
      for (let step = 0; step < 10; step++) {
        const operation = documents.size === 0 ? "upload" : pick(["upload", "upload", "edit", "delete", "rename", "rewrap"])
        const name = documents.size === 0 ? pick(names) : pick([...documents.keys()])
        const current = documents.get(name)
        if (operation === "upload") {
          const target = pick(names)
          documents.delete(target)
          documents.set(target, { texts: [pick(sentences), pick(sentences)], wrapped: pick([false, true]) })
        }
        if (operation === "edit" && current) documents.set(name, { ...current, texts: [current.texts[0]!, pick(sentences)] })
        if (operation === "delete") documents.delete(name)
        if (operation === "rename" && current) {
          const target = pick(names)
          documents.delete(name)
          documents.delete(target)
          documents.set(target, current)
        }
        if (operation === "rewrap" && current) documents.set(name, { ...current, wrapped: !current.wrapped })
        const chunks = [...documents].flatMap(([document, { texts, wrapped }]) => matterChunks({ [document]: texts }).map((chunk) =>
          wrapped ? { ...chunk, file: { text: chunk.text.replace(" ", "\n"), inserted: [] } } : chunk))
        const data = await extractIdentifiers(chunks, previous)
        expect(withoutTime(data)).toBe(withoutTime(await extractIdentifiers(chunks)))
        previous = JSON.parse(JSON.stringify(data)) as MatterIdentifiers
      }
    }
  })

  test("lets other work run between the chunks it parses", async () => {
    const chunks = matterChunks({ "Long.docx": Array.from({ length: 16 }, (_, index) => `${index + 1}. ${notice[0]}`) })
    const scan = spyOn(identifiers, "scanChunk")
    try {
      const parsedWhenTimerRan = new Promise<number>((resolve) => setTimeout(() => resolve(scan.mock.calls.length), 0))
      await extractIdentifiers(chunks)
      expect(scan).toHaveBeenCalledTimes(chunks.length)
      expect(await parsedWhenTimerRan).toBeLessThan(chunks.length)
    } finally {
      scan.mockRestore()
    }
  })
})

describe("wrapped chunk files", () => {
  test("fileOffset moves an offset past every line break inserted at or before it", () => {
    expect(fileOffset(5, [])).toBe(5)
    expect(fileOffset(999, [1000, 2000])).toBe(999)
    expect(fileOffset(1000, [1000, 2000])).toBe(1001)
    expect(fileOffset(2500, [1000, 2000])).toBe(2502)
  })

  test("extracts values from the source text and locates them in the wrapped file", async () => {
    // The file hard-cuts the run of x's (a line break inserted before source
    // offset 5) and turns three spaces into line breaks, two of them inside names.
    const text = "xxxxxxxxxx Taylor Morgan signed. The notice was signed by Morgan Ellis for Example Procurement Services today."
    const file = "xxxxx\nxxxxx\nTaylor Morgan signed. The notice was signed by Morgan\nEllis for Example Procurement\nServices today."
    const line3 = "Taylor Morgan signed. The notice was signed by Morgan"
    const chunk: IdentifierChunk = { id: "doc-c001", document: "doc", text, file: { text: file, inserted: [5] } }
    const rows = (await extractIdentifiers([chunk])).rows
    const located = (type: string) => rows.filter((row) => row.type === type).map((row) => [row.value, row.line, row.context])
    expect(located("person")).toEqual([
      ["Taylor Morgan", 3, line3],
      ["Morgan Ellis", 3, line3],
    ])
    expect(located("organization")).toEqual([["Morgan Ellis for Example Procurement Services", 3, line3]])
  })
})

describe("renderIdentifiersMarkdown", () => {
  test("renders grouped table by type with counts", async () => {
    const data = await extractIdentifiers([sampleChunk])
    const md = renderIdentifiersMarkdown(data)
    expect(md).toContain("# Matter Identifiers")
    expect(md).toContain("## person")
    expect(md).toContain("## case_reference")
    expect(md).toContain("| Value | Source | Chunk | Line | Context |")
    expect(md).toContain("synthetic-procurement-memo-c007")
  })

  test("markdown is stable for empty inventory", () => {
    const md = renderIdentifiersMarkdown({ schema_version: IDENTIFIERS_SCHEMA_VERSION, generated_at: "2026-06-28T00:00:00Z", documents: 0, sources: [], rows: [] })
    expect(md).toContain("# Matter Identifiers")
    expect(md).toContain("0 identifiers")
  })
})
