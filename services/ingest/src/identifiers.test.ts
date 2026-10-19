import { describe, expect, test } from "bun:test"
import { extractIdentifiers, fileOffset, IDENTIFIERS_SCHEMA_VERSION, renderIdentifiersMarkdown, type IdentifierChunk } from "./identifiers"

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

function rowsOfType(data: ReturnType<typeof extractIdentifiers>, type: string) {
  return data.rows.filter((row) => row.type === type).map((row) => row.value)
}

describe("extractIdentifiers", () => {
  test("schema and metadata are populated", () => {
    const data = extractIdentifiers([sampleChunk])
    expect(data.schema_version).toBe(IDENTIFIERS_SCHEMA_VERSION)
    expect(data.documents).toBe(1)
    expect(data.generated_at).toMatch(/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}/)
    expect(data.rows.length).toBeGreaterThan(0)
  })

  test("extracts person names", () => {
    const data = extractIdentifiers([sampleChunk])
    expect(rowsOfType(data, "person")).toEqual(expect.arrayContaining([
      "Alex Morgan Bennett",
      "Taylor R. Morgan",
    ]))
  })

  test("extracts organizations", () => {
    const data = extractIdentifiers([sampleChunk])
    expect(rowsOfType(data, "organization")).toEqual(
      expect.arrayContaining([
        "Example Procurement Services",
        "Example Research Institute",
      ]),
    )
  })

  test("extracts places", () => {
    const data = extractIdentifiers([sampleChunk])
    const places = rowsOfType(data, "place")
    expect(places.some((p) => p.includes("Baltimore"))).toBe(true)
    expect(places.some((p) => p.includes("Maryland"))).toBe(true)
  })

  test("extracts money", () => {
    const data = extractIdentifiers([sampleChunk])
    expect(rowsOfType(data, "money")).toContain("$437.00")
  })

  test("extracts dates", () => {
    const data = extractIdentifiers([sampleChunk])
    expect(rowsOfType(data, "date")).toEqual(
      expect.arrayContaining([
        expect.stringMatching(/August 3, 2023/),
      ]),
    )
  })

  test("extracts phone numbers", () => {
    const data = extractIdentifiers([sampleChunk])
    expect(rowsOfType(data, "phone")).toContain("(410) 555-0113")
  })

  test("extracts emails", () => {
    const data = extractIdentifiers([sampleChunk])
    expect(rowsOfType(data, "email")).toContain("coordinator@example.com")
  })

  test("extracts case references via custom regex", () => {
    const data = extractIdentifiers([sampleChunk])
    expect(rowsOfType(data, "case_reference")).toEqual(
      expect.arrayContaining(["00-TEST-0000001"]),
    )
  })

  test("extracts statute citations via custom regex", () => {
    const data = extractIdentifiers([sampleChunk])
    const statutes = rowsOfType(data, "statute_citation")
    expect(statutes).toEqual(
      expect.arrayContaining([
        expect.stringMatching(/15 U\.S\.C\. § 7001/),
        expect.stringMatching(/MD Com Law § 1-101/),
      ]),
    )
  })

  test("extracts state-law Title/Subtitle statute citations without § symbol", () => {
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
    const statutes = rowsOfType(extractIdentifiers([chunk]), "statute_citation")
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

  test("each row carries source, chunk_id, line, and context", () => {
    const data = extractIdentifiers([sampleChunk])
    for (const row of data.rows) {
      expect(row.source).toBe("synthetic-procurement-memo")
      expect(row.chunk_id).toBe("synthetic-procurement-memo-c007")
      expect(Number.isInteger(row.line)).toBe(true)
      expect(row.line).toBeGreaterThanOrEqual(1)
      expect(row.context.length).toBeGreaterThan(0)
      expect(row.context.length).toBeLessThanOrEqual(200)
    }
  })

  test("line numbers map to chunk-local line positions", () => {
    const data = extractIdentifiers([sampleChunk])
    const caseRow = data.rows.find((row) => row.type === "case_reference")
    expect(caseRow).toBeDefined()
    // "Example Procurement Services filed Case No. 00-TEST-0000001."
    // is line 3 of the chunk (1-indexed).
    expect(caseRow!.line).toBe(3)
  })

  test("context contains the matched value", () => {
    const data = extractIdentifiers([sampleChunk])
    const emailRow = data.rows.find((row) => row.type === "email")
    expect(emailRow!.context).toContain("coordinator@example.com")
  })

  test("multiple chunks attribute correctly", () => {
    const secondChunk: IdentifierChunk = {
      id: "intake-c001",
      document: "intake-memo",
      text: "Synthetic Docket No. 00-TEST-0000002 was filed in 2024.",
    }
    const data = extractIdentifiers([sampleChunk, secondChunk])
    const intakeRow = data.rows.find((row) => row.chunk_id === "intake-c001")
    expect(intakeRow).toBeDefined()
    expect(intakeRow!.source).toBe("intake-memo")
    expect(intakeRow!.value).toBe("00-TEST-0000002")
    expect(data.documents).toBe(2)
  })

  test("empty chunks produce empty inventory", () => {
    const data = extractIdentifiers([])
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
  test("collapses initial-middle variants to the full-middle form", () => {
    const variantChunk: IdentifierChunk = {
      id: "doc-c001",
      document: "doc",
      text: [
        "Taylor R. Morgan coordinates the fictional project.",
        "Counsel for Taylor Renee Morgan confirmed consent.",
      ].join("\n"),
    }
    const data = extractIdentifiers([variantChunk])
    const personValues = rowsOfType(data, "person")
    // No row should still carry "Taylor R. Morgan" — all collapsed to "Taylor Renee Morgan"
    expect(personValues).not.toContain("Taylor R. Morgan")
    expect(personValues.filter((v) => v === "Taylor Renee Morgan").length).toBeGreaterThan(0)
  })

  test("does NOT collapse when neither middle is a single initial", () => {
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
    const data = extractIdentifiers([distinctChunk])
    const personValues = rowsOfType(data, "person")
    expect(personValues).toContain("Jordan Alex Lee")
    expect(personValues).toContain("Jordan Alexis Lee")
  })

  test("does NOT collapse two-word names with no middle", () => {
    // "Sarah Chen" appears in many sources. There is no variant here, and no
    // middle to compare. Conservative: never collapse a two-word row.
    const twoWordChunk: IdentifierChunk = {
      id: "doc-c001",
      document: "doc",
      text: "Sarah Chen represented the petitioner. Sarah Chen is partner.",
    }
    const data = extractIdentifiers([twoWordChunk])
    const personValues = rowsOfType(data, "person")
    expect(personValues).toContain("Sarah Chen")
  })

  test("preserves source diversity after collapse", () => {
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
    const data = extractIdentifiers([multiSourceChunkA, multiSourceChunkB])
    const personRows = data.rows.filter((r) => r.type === "person" && r.value === "Taylor Renee Morgan")
    expect(personRows.length).toBeGreaterThanOrEqual(2)
    expect(personRows.some((r) => r.source === "docA.docx")).toBe(true)
    expect(personRows.some((r) => r.source === "docB.docx")).toBe(true)
  })

  test("original surface form survives in the Context column after collapse", () => {
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
    const data = extractIdentifiers([variantChunk, fullChunk])
    const normalizedRow = data.rows.find(
      (r) => r.type === "person" && r.value === "Taylor Renee Morgan" && r.chunk_id === "doc-c001",
    )
    expect(normalizedRow).toBeDefined()
    expect(normalizedRow!.context).toContain("Taylor R. Morgan")
  })
})

describe("wrapped chunk files", () => {
  test("fileOffset moves an offset past every line break inserted at or before it", () => {
    expect(fileOffset(5, [])).toBe(5)
    expect(fileOffset(999, [1000, 2000])).toBe(999)
    expect(fileOffset(1000, [1000, 2000])).toBe(1001)
    expect(fileOffset(2500, [1000, 2000])).toBe(2502)
  })

  test("extracts values from the source text and locates them in the wrapped file", () => {
    // The file hard-cuts the run of x's (a line break inserted before source
    // offset 5) and turns three spaces into line breaks, two of them inside names.
    const text = "xxxxxxxxxx Taylor Morgan signed. The notice was signed by Morgan Ellis for Example Procurement Services today."
    const file = "xxxxx\nxxxxx\nTaylor Morgan signed. The notice was signed by Morgan\nEllis for Example Procurement\nServices today."
    const line3 = "Taylor Morgan signed. The notice was signed by Morgan"
    const chunk: IdentifierChunk = { id: "doc-c001", document: "doc", text, file: { text: file, inserted: [5] } }
    const rows = extractIdentifiers([chunk]).rows
    const located = (type: string) => rows.filter((row) => row.type === type).map((row) => [row.value, row.line, row.context])
    expect(located("person")).toEqual([
      ["Taylor Morgan", 3, line3],
      ["Morgan Ellis", 3, line3],
    ])
    expect(located("organization")).toEqual([["Morgan Ellis for Example Procurement Services", 3, line3]])
  })
})

describe("renderIdentifiersMarkdown", () => {
  test("renders grouped table by type with counts", () => {
    const data = extractIdentifiers([sampleChunk])
    const md = renderIdentifiersMarkdown(data)
    expect(md).toContain("# Matter Identifiers")
    expect(md).toContain("## person")
    expect(md).toContain("## case_reference")
    expect(md).toContain("| Value | Source | Chunk | Line | Context |")
    expect(md).toContain("synthetic-procurement-memo-c007")
  })

  test("markdown is stable for empty inventory", () => {
    const md = renderIdentifiersMarkdown({ schema_version: IDENTIFIERS_SCHEMA_VERSION, generated_at: "2026-06-28T00:00:00Z", documents: 0, rows: [] })
    expect(md).toContain("# Matter Identifiers")
    expect(md).toContain("0 identifiers")
  })
})
