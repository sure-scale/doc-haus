import { expect, test } from "bun:test"
import { mkdtempSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { insertChunk, openDb, setMeta, upsertDocument } from "../../services/ingest/src/db"
import { buildSectionPaths, extractStructure } from "../../services/ingest/src/structure"
import type { DocumentCitation } from "../lib/citations"

type SearchArgs = { query: string; k?: number; document?: string; structural_path?: string }
type SearchResult = {
  output: string
  metadata: {
    citations: DocumentCitation[]
    pending: Array<{ new_text: string }>
    pending_context?: { status: string; reason: string }
  }
}

// Only embedding vectors are synthetic. A fresh subprocess owns the mock so it
// cannot change model loading in another suite; all SQL, FTS, fusion and output
// formatting run through the actual search tool.
async function search(directory: string, args: SearchArgs[]) {
  const child = Bun.spawn([process.execPath, "--no-env-file", "-e", `
    import { mock } from "bun:test"
    import { z } from "zod"
    mock.module("@huggingface/transformers", () => ({
      env: {},
      pipeline: async () => async () => {
        const data = new Float32Array(384)
        data[0] = 1
        return { data }
      },
    }))
    const searchDocument = (await import("../tool/search-document")).default
    const input = JSON.parse(await Bun.stdin.text())
    const results = []
    for (const args of input.args) {
      const valid = z.object(searchDocument.args).safeParse(args).success
      results.push({ valid, result: valid ? await searchDocument.execute(args, {
        directory: input.directory,
        sessionID: crypto.randomUUID(),
      }) : null })
    }
    console.log(JSON.stringify(results))
  `], {
    cwd: import.meta.dir,
    env: { PATH: process.env.PATH ?? "" },
    stdin: "pipe",
    stdout: "pipe",
    stderr: "pipe",
  })
  child.stdin.write(JSON.stringify({ directory, args }))
  child.stdin.end()
  const [stdout, stderr, code] = await Promise.all([
    new Response(child.stdout).text(),
    new Response(child.stderr).text(),
    child.exited,
  ])
  expect({ code, stderr }).toEqual({ code: 0, stderr: "" })
  return JSON.parse(stdout) as Array<{ valid: boolean; result: SearchResult | null }>
}

function matter() {
  const directory = mkdtempSync(path.join(tmpdir(), "search-filters-"))
  const db = openDb(directory)
  setMeta(db, "embedding_model", "onnx-community/granite-embedding-small-english-r2-ONNX")
  const addDocument = (name: string) => {
    const document = path.join(directory, name)
    writeFileSync(document, "synthetic source")
    return { id: upsertDocument(db, document, name, 1, null), path: document, name }
  }
  const addChunk = (
    document: ReturnType<typeof addDocument>,
    structuralPath: string,
    text: string,
    similarity = 0.2,
    start = 0,
    section = "1",
    flagged = false,
  ) => {
    const embedding = new Float32Array(384)
    embedding[0] = similarity
    embedding[1] = Math.sqrt(1 - similarity * similarity)
    insertChunk(db, {
      documentId: document.id,
      docPath: document.path,
      docName: document.name,
      section,
      structuralPath,
      party: "participant context",
      chunkIndex: start,
      text,
      charStart: start,
      charEnd: start + text.length,
      embedding,
      flagged,
    })
  }
  return { directory, db, addDocument, addChunk }
}

test.each([
  ["vector", "unmatched semantic query", "excluded passage", "permitted passage"],
  ["lexical", "needle", "needle", `needle ${"padding ".repeat(80)}`],
  ["phrase", "literal sequence", "literal sequence", `literal sequence ${"padding ".repeat(80)}`],
])("%s candidates are structurally filtered before the channel limit and ranking", async (_channel, query, excluded, permitted) => {
  const fixture = matter()
  const document = fixture.addDocument("source.docx")
  const scope = "Doc(source.docx) › §3 Included"
  // More excluded, stronger candidates than the channel cap. A post-filter
  // would miss the eligible row entirely; a missed channel leaks exclusions.
  for (let i = 0; i < 21; i++)
    fixture.addChunk(document, "Doc(source.docx) › §9 Excluded", `${excluded} ${i}`, 1, i * 1000, "9")
  fixture.addChunk(document, scope, permitted, 0.2, 30000, "3")
  fixture.db.close()
  try {
    const [unfiltered, filtered] = await search(fixture.directory, [
      { query, k: 20 },
      { query, k: 20, structural_path: scope },
    ])
    expect(unfiltered.result!.metadata.citations).toHaveLength(20)
    expect(unfiltered.result!.metadata.citations.every((citation) => citation.section === "9")).toBe(true)
    expect(filtered.result!.metadata.citations).toHaveLength(1)
    expect(filtered.result!.metadata.citations[0]).toMatchObject({
      structuralPath: scope,
      excerpt: permitted,
      charStart: 30000,
      charEnd: 30000 + permitted.length,
      docPath: document.path,
    })
  } finally {
    rmSync(fixture.directory, { recursive: true, force: true })
  }
})

test("structural scope is a literal, case-sensitive full segment and its descendants", async () => {
  const fixture = matter()
  const document = fixture.addDocument("Notes%_\"'.docx")
  const parent = { label: "3", number: "3", title: "Heading_%'\" \\literal", text: "", charStart: 0, charEnd: 0 }
  const child = { ...parent, label: "3.1", number: "3.1", title: "Child" }
  const grandchild = { ...parent, label: "3.1.1", number: "3.1.1", title: "Grandchild" }
  const paths = buildSectionPaths(document.name, [parent, child, grandchild])
  const scope = paths.get(parent)!
  const eligible = [scope, paths.get(child)!, paths.get(grandchild)!]
  const excluded = [
    `${scope} suffix`,
    scope.replace("Heading", "heading"),
    scope.replace("%_", "xx"),
    scope.replace("§3 ", "§30 "),
    `${scope}›Not a stored delimiter`,
    "",
  ]
  ;[...eligible, ...excluded].forEach((structuralPath, i) =>
    fixture.addChunk(document, structuralPath, `needle passage ${i}`, i < eligible.length ? 0.2 : 1, i * 100),
  )
  fixture.db.close()
  try {
    const [filtered, partial, absent] = await search(fixture.directory, [
      { query: "needle", k: 20, structural_path: scope },
      { query: "needle", k: 20, structural_path: scope.slice(0, scope.indexOf(" Heading")) },
      { query: "needle", k: 20, structural_path: "Doc(not indexed.docx)" },
    ])
    expect(filtered.result!.metadata.citations.map((citation) => citation.structuralPath).sort()).toEqual(eligible.sort())
    expect(partial.result!.metadata.citations).toHaveLength(0)
    expect(absent.result!.metadata.citations).toHaveLength(0)
    expect(absent.result!.output).toContain("No relevant passages found.")
  } finally {
    rmSync(fixture.directory, { recursive: true, force: true })
  }
})

test("document and structural scopes intersect, including an unknown document name", async () => {
  const fixture = matter()
  const first = fixture.addDocument("First.docx")
  const second = fixture.addDocument("Second.docx")
  const scope = "Doc(First.docx) › §3 Included"
  fixture.addChunk(first, scope, "needle first included", 0.2)
  fixture.addChunk(first, "Doc(First.docx) › §9 Other", "needle first other", 1, 100)
  fixture.addChunk(second, "Doc(Second.docx) › §3 Included", "needle second included", 1)
  fixture.db.close()
  try {
    const [both, mismatch, unknown, documentOnly] = await search(fixture.directory, [
      { query: "needle", k: 20, document: first.name, structural_path: scope },
      { query: "needle", k: 20, document: second.name, structural_path: scope },
      { query: "needle", k: 20, document: "unknown.docx", structural_path: scope },
      { query: "needle", k: 20, document: first.name },
    ])
    expect(both.result!.metadata.citations.map((citation) => citation.excerpt)).toEqual(["needle first included"])
    expect(mismatch.result!.metadata.citations).toHaveLength(0)
    // A document with no indexed passages is refused before any source search runs.
    expect(String(unknown.result)).toContain("Source view unavailable")
    expect(documentOnly.result!.metadata.citations.map((citation) => citation.excerpt).sort()).toEqual([
      "needle first included", "needle first other",
    ])
    expect(documentOnly.result!.metadata).not.toHaveProperty("pending_context")
  } finally {
    rmSync(fixture.directory, { recursive: true, force: true })
  }
})

test("missing structural metadata stays searchable unfiltered but cannot satisfy an explicit path", async () => {
  const fixture = matter()
  const document = fixture.addDocument("legacy.docx")
  const text = "alpha legacy source"
  fixture.addChunk(document, "", text, 1, 37, "A", true)
  fixture.db.close()
  try {
    const [unfiltered, scoped, invalid] = await search(fixture.directory, [
      { query: "alpha" },
      { query: "alpha", structural_path: "Doc(legacy.docx)" },
      { query: "alpha", structural_path: "" },
    ])
    expect(unfiltered.result!.metadata.citations).toHaveLength(1)
    expect(unfiltered.result!.metadata.citations[0]).toMatchObject({
      documentName: document.name,
      docPath: document.path,
      section: "A",
      excerpt: text,
      charStart: 37,
      charEnd: 37 + text.length,
      flagged: true,
      party: "participant context",
    })
    expect(unfiltered.result!.metadata.citations[0].score).toBeCloseTo(2 / 3, 10)
    expect(unfiltered.result!.metadata.citations[0]).not.toHaveProperty("structuralPath")
    expect(unfiltered.result!.output).toContain("legacy.docx § A")
    expect(unfiltered.result!.output).toContain("instruction")
    expect(scoped.result!.metadata.citations).toHaveLength(0)
    expect(invalid.valid).toBe(false)
    expect(invalid.result).toBeNull()
  } finally {
    rmSync(fixture.directory, { recursive: true, force: true })
  }
})

test("definitions, cross-reference hints and pending context respect an explicit strict scope", async () => {
  const fixture = matter()
  const document = fixture.addDocument("source.docx")
  const scope = "Doc(source.docx) › §3 Included"
  const text = "needle InsideTerm OutsideTerm AcrossTerm; Section 3.1 and Section 9. This clause uses those terms."
  fixture.addChunk(document, scope, text, 1, 100, "3")
  fixture.addChunk(document, `${scope} › §3.1 Child`, "needle permitted child", 0.2, 300, "3.1")
  fixture.addChunk(document, "Doc(source.docx) › §9 Outside", "outside definition source", 0.1, 500, "9")
  for (const [term, definition, start, end] of [
    ["OutsideTerm", `outside-definition-${crypto.randomUUID()}`, 500, 510],
    ["AcrossTerm", `unproven-span-${crypto.randomUUID()}`, 190, 510],
  ] as const)
    fixture.db.run("INSERT INTO defined_terms (doc_path, doc_name, term, definition, char_start, char_end) VALUES (?, ?, ?, ?, ?, ?)",
      [document.path, document.name, term, definition, start, end])
  const childDefinition = '"InsideTerm" means the permitted child definition.'
  fixture.addChunk(document, `${scope} › §3.1 Child`, childDefinition, 0.1, 350, "3.1")
  fixture.db.run("INSERT INTO defined_terms (doc_path, doc_name, term, definition, char_start, char_end) VALUES (?, ?, ?, ?, ?, ?)",
    [document.path, document.name, "InsideTerm", childDefinition, 350, 350 + childDefinition.length])
  const pendingText = `unassessed-proposal-${crypto.randomUUID()}`
  fixture.db.run("INSERT INTO redlines (doc_path, doc_name, scope, find_text, old_text, new_text, author, created_at) VALUES (?, ?, 'phrase', 'outside definition', 'outside definition', ?, 'reviewer', 1)",
    [document.path, document.name, pendingText])
  fixture.db.close()
  try {
    const [unfiltered, filtered, documentOnly] = await search(fixture.directory, [
      { query: "needle", k: 1 },
      { query: "needle", k: 1, structural_path: scope },
      { query: "needle", k: 1, document: document.name },
    ])
    expect(documentOnly.result).toEqual(unfiltered.result)
    expect(unfiltered.result!.output).toContain(childDefinition)
    expect(unfiltered.result!.output).toContain("outside-definition-")
    expect(unfiltered.result!.output).toContain("unproven-span-")
    expect(unfiltered.result!.output).toContain(`- source.docx has a section labelled "9": §9 Outside`)
    expect(unfiltered.result!.output).toContain(pendingText)
    expect(unfiltered.result!.metadata.pending).toHaveLength(1)
    expect(unfiltered.result!.metadata).not.toHaveProperty("pending_context")
    expect(filtered.result!.output).not.toContain("outside-definition-")
    expect(filtered.result!.output).not.toContain("unproven-span-")
    expect(filtered.result!.output).toContain(childDefinition)
    expect(filtered.result!.output).not.toContain(`labelled "9"`)
    expect(filtered.result!.output).toContain(`- source.docx has a section labelled "3.1": §3.1 Child`)
    expect(filtered.result!.output).not.toContain(pendingText)
    expect(filtered.result!.metadata.pending).toHaveLength(0)
    expect(filtered.result!.metadata.pending_context).toEqual({ status: "omitted", reason: "unassessed_for_structural_scope" })
    expect(filtered.result!.output).toContain("Pending redline context omitted for this structural scope; pending proposals have not been assessed.")
  } finally {
    rmSync(fixture.directory, { recursive: true, force: true })
  }
})

const HINT_HEADER =
  "Sections whose labels match references in these passages (a matching label does not show which document a reference means; to read one, pass both its document and section label to get-section):"

// The section-label note is one blank-line-separated block of the output.
function hintOf(output: string) {
  return output.split("\n\n").find((block) => block.startsWith("Sections whose labels")) ?? ""
}

test("section-label hints match the reference kind and state a label, not a referent", async () => {
  const fixture = matter()
  const document = fixture.addDocument("Guarantee.docx")
  const root = "Doc(Guarantee.docx)"
  // Labels in the shapes ingest records: a numbered line ("8.3 Notices") keeps
  // only its number; an unnumbered heading keeps its own text.
  const text = "needle The Guarantor complies with Section 8.3 of the Credit Agreement, Article 5 of the Regulation, Clause 4.1, Schedule 3, Annex B and Exhibit C."
  fixture.addChunk(document, `${root} › §2 Guarantee`, text, 1, 0, "2")
  fixture.addChunk(document, `${root} › §3 Representations`, "representations", 0.1, 1000, "3")
  fixture.addChunk(document, `${root} › §4.1 Costs`, "costs", 0.1, 2000, "4.1")
  fixture.addChunk(document, `${root} › §5 Payments`, "payments", 0.1, 3000, "5")
  fixture.addChunk(document, `${root} › §8.3 Notices`, "notices", 0.1, 4000, "8.3")
  fixture.addChunk(document, `${root} › EXHIBIT B - FORM OF NOTICE`, "form of notice", 0.1, 5000, "EXHIBIT B - FORM OF NOTICE")
  fixture.addChunk(document, `${root} › Exhibit C`, "exhibit c", 0.1, 6000, "Exhibit C")
  fixture.db.close()
  try {
    const [result] = await search(fixture.directory, [{ query: "needle", k: 1 }])
    const hint = hintOf(result.result!.output)
    // A numbered label records no kind: Schedule 3 and Article 5 are other
    // numbering series, and Annex B is not the EXHIBIT B heading.
    expect(hint).not.toContain(`"3"`)
    expect(hint).not.toContain(`"5"`)
    expect(hint).not.toContain("EXHIBIT B")
    // The note states which labels the document has; the reading agent decides
    // what each reference means (Section 8.3 here names another agreement).
    expect(hint).not.toContain("→")
    expect(hint).not.toMatch(/resolve/i)
    expect(hint).toBe(
      [
        HINT_HEADER,
        `- Guarantee.docx has a section labelled "8.3": §8.3 Notices`,
        `- Guarantee.docx has a section labelled "4.1": §4.1 Costs`,
        `- Guarantee.docx has a section labelled "Exhibit C"`,
      ].join("\n"),
    )
  } finally {
    rmSync(fixture.directory, { recursive: true, force: true })
  }
})

test("section-label hints keep numbering boundaries, kinds, label separators and letter case", async () => {
  const fixture = matter()
  const add = (name: string, citing: string, sections: Array<[label: string, crumb: string]>) => {
    const document = fixture.addDocument(name)
    fixture.addChunk(document, `Doc(${name}) › §1 Citing`, `needle ${citing}`, 1, 0, "1")
    sections.forEach(([label, crumb], i) =>
      fixture.addChunk(document, `Doc(${name}) › ${crumb}`, `body ${i}`, 0.1, 1000 * (i + 1), label))
  }
  // 8.3 is neither 8.30 nor 8.31; without its own label, 9.2 reports its child.
  add("Numbers.docx", "Section 8.3 and Section 9.2.", [
    ["8.30", "§8.30 Thirty"],
    ["8.31", "§8.31 Thirty-one"],
    ["9.2.1", "§9.2.1 Child"],
  ])
  // A roman-numeral Article matches its kind-named heading, never clause 5.
  add("Roman.docx", "Article V applies.", [
    ["5", "§5 Five"],
    ["ARTICLE V", "ARTICLE V"],
  ])
  // The label may be followed by punctuation; a longer label (A-1) is another
  // section; the label compares without letter case.
  add("Punctuation.docx", "Article IV, Schedule 2, Exhibit A and Exhibit D.", [
    ["Article IV—Definitions", "Article IV—Definitions"],
    ["Schedule 2: Pricing", "Schedule 2: Pricing"],
    ["EXHIBIT A-1", "EXHIBIT A-1"],
    ["EXHIBIT d", "EXHIBIT d"],
  ])
  fixture.db.close()
  try {
    const [numbers, roman, punctuation] = await search(fixture.directory, [
      { query: "needle", k: 1, document: "Numbers.docx" },
      { query: "needle", k: 1, document: "Roman.docx" },
      { query: "needle", k: 1, document: "Punctuation.docx" },
    ])
    expect(hintOf(numbers.result!.output)).toBe(
      [HINT_HEADER, `- Numbers.docx has a section labelled "9.2.1": §9.2.1 Child`].join("\n"),
    )
    expect(hintOf(roman.result!.output)).toBe([HINT_HEADER, `- Roman.docx has a section labelled "ARTICLE V"`].join("\n"))
    expect(hintOf(punctuation.result!.output)).toBe(
      [
        HINT_HEADER,
        `- Punctuation.docx has a section labelled "Article IV—Definitions"`,
        `- Punctuation.docx has a section labelled "Schedule 2: Pricing"`,
        `- Punctuation.docx has a section labelled "EXHIBIT d"`,
      ].join("\n"),
    )
  } finally {
    rmSync(fixture.directory, { recursive: true, force: true })
  }
})

test("section-label hints name each document, respect strict scope and clip long headings", async () => {
  const fixture = matter()
  const alpha = fixture.addDocument("Alpha.docx")
  const beta = fixture.addDocument("Beta.docx")
  fixture.addChunk(alpha, "Doc(Alpha.docx) › §1 Citing", "pairquery Section 4 and Section 7.", 1, 0, "1")
  fixture.addChunk(alpha, "Doc(Alpha.docx) › §4 Alpha four", "alpha four", 0.1, 1000, "4")
  fixture.addChunk(beta, "Doc(Beta.docx) › §1 Citing", "pairquery Section 4.", 1, 0, "1")
  fixture.addChunk(beta, "Doc(Beta.docx) › §4 Beta four", "beta four", 0.1, 1000, "4")
  // Beta's 7 is not a label of Alpha, where the reference appears.
  fixture.addChunk(beta, "Doc(Beta.docx) › §7 Beta seven", "beta seven", 0.1, 2000, "7")

  const scoped = fixture.addDocument("Scoped.docx")
  const scope = "Doc(Scoped.docx) › §3 Included"
  fixture.addChunk(scoped, scope, "scopequery Exhibit B and Section 3.1.", 1, 0, "3")
  fixture.addChunk(scoped, `${scope} › §3.1 Child`, "child", 0.1, 1000, "3.1")
  fixture.addChunk(scoped, "Doc(Scoped.docx) › EXHIBIT B", "exhibit b", 0.1, 2000, "EXHIBIT B")

  // 76 x's put the emoji (two UTF-16 units) at the 80th code point.
  const long = fixture.addDocument("Long.docx")
  const heading = `§3 ${"x".repeat(76)}😀${"y".repeat(20)}`
  fixture.addChunk(long, "Doc(Long.docx) › §1 Citing", "longquery Clause 3.", 1, 0, "1")
  fixture.addChunk(long, `Doc(Long.docx) › ${heading}`, "long", 0.1, 1000, "3")
  fixture.db.close()
  try {
    const [pair, unscoped, strict, clipped] = await search(fixture.directory, [
      { query: "pairquery", k: 2 },
      { query: "scopequery", k: 1, document: "Scoped.docx" },
      { query: "scopequery", k: 1, structural_path: scope },
      { query: "longquery", k: 1, document: "Long.docx" },
    ])
    const pairHint = hintOf(pair.result!.output).split("\n")
    expect(pairHint[0]).toBe(HINT_HEADER)
    expect(pairHint.slice(1).toSorted()).toEqual([
      `- Alpha.docx has a section labelled "4": §4 Alpha four`,
      `- Beta.docx has a section labelled "4": §4 Beta four`,
    ])
    expect(hintOf(unscoped.result!.output)).toContain(`- Scoped.docx has a section labelled "EXHIBIT B"`)
    expect(hintOf(strict.result!.output)).toBe(
      [HINT_HEADER, `- Scoped.docx has a section labelled "3.1": §3.1 Child`].join("\n"),
    )
    expect(hintOf(clipped.result!.output)).toBe(
      [HINT_HEADER, `- Long.docx has a section labelled "3": §3 ${"x".repeat(76)}😀…`].join("\n"),
    )
  } finally {
    rmSync(fixture.directory, { recursive: true, force: true })
  }
})

test("structural passage selection preserves matter-wide supersession metadata and its existing penalty", async () => {
  const fixture = matter()
  const source = fixture.addDocument("Source.docx")
  const amender = fixture.addDocument("Update.docx")
  const scope = "Doc(Source.docx) › §3 Included"
  fixture.addChunk(source, scope, "needle source", 1, 0, "3")
  const amendmentText = `Amendment to Source.docx\nSection 3 is hereby amended and restated.\nneedle ${"amendment body ".repeat(40)}`
  fixture.addChunk(amender, "Doc(Update.docx) › §3 Amendment", amendmentText, 0.2, 0, "3")
  // Use the normal deterministic producer for relation and operation evidence.
  extractStructure(fixture.db, amender.path, amender.name, amendmentText)
  const operation = fixture.db.query("SELECT char_start, char_end FROM section_ops WHERE doc_path = ?").get(amender.path) as {
    char_start: number
    char_end: number
  }
  fixture.db.close()
  try {
    const [unfiltered, scoped] = await search(fixture.directory, [
      { query: "needle", k: 20 },
      { query: "needle", k: 20, structural_path: scope },
    ])
    const original = unfiltered.result!.metadata.citations.find((citation) => citation.docPath === source.path)!
    expect(unfiltered.result!.metadata.citations[0].docPath).toBe(amender.path)
    // The code-point position runs from the operative statement to the end of
    // the indexed amendment text (no later statement bounds it).
    const read = { document: amender.name, offset: operation.char_start, limit: amendmentText.length - operation.char_start }
    expect(original.superseded_by).toEqual({
      amenderName: amender.name,
      operation: "restate",
      charStart: operation.char_start,
      charEnd: operation.char_end,
      read,
    })
    expect(original.score).toBeCloseTo(1 / 3, 10)
    expect(scoped.result!.metadata.citations).toEqual([original])
    expect(scoped.result!.output).toContain("[⚠ SUPERSEDED]")
    expect(scoped.result!.output).toContain(`superseded by: ${amender.name}`)
    expect(scoped.result!.output).toContain(
      `read the amendment's operative wording in ${amender.name} at code-point offset ${read.offset} (${read.limit} characters).`,
    )
    expect(scoped.result!.output).not.toContain("read-document")
    expect(scoped.result!.output).not.toContain("current-section")
  } finally {
    rmSync(fixture.directory, { recursive: true, force: true })
  }
})
