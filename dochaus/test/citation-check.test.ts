import { describe, expect, test } from "bun:test"
import { mkdirSync, mkdtempSync, readFileSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { mergeAuthorityLedgerEntries } from "../lib/authority-ledger"
import { buildCitationCheckReport, formatCitationCheckSummary, writeCitationCheckReport } from "../lib/citation-check"
import { citationRowsForArtifact } from "../lib/citation-extraction"

function matter() {
  return mkdtempSync(path.join(tmpdir(), "citation-check-"))
}

function writeSourceText(dir: string, file: string, text: string) {
  mkdirSync(path.dirname(path.join(dir, file)), { recursive: true })
  writeFileSync(path.join(dir, file), text)
  writeFileSync(
    path.join(dir, "source-index.json"),
    JSON.stringify(
      {
        source_ledger: "source-ledger.json",
        source_map: "source-map.json",
        source_map_markdown: "source-map.md",
        source_chunk_directory: "source-chunks",
        sources: [
          {
            document: "Generic Source",
            doc_path: file,
            extracted_text: file,
            characters: text.length,
          },
        ],
      },
      null,
      2,
    ),
  )
}

function researchContent(rows: unknown[]) {
  return `\`\`\`research\n${JSON.stringify(rows, null, 2)}\n\`\`\`\n`
}

describe("buildCitationCheckReport", () => {
  test("groups repeated legal citations and links them to source artifacts", () => {
    const dir = matter()
    writeSourceText(dir, "sources/generic-source.txt", "The Generic Authority Act § 12 requires written notice.\n")
    writeFileSync(
      path.join(dir, "agreement-draft.md"),
      "# Draft\n\nThe notice position follows Generic Authority Act § 12.\n",
    )
    writeFileSync(
      path.join(dir, "drafting-notes-memo.md"),
      "# Memo\n\nThe same point is supported by Generic Authority Act   § 12.\n",
    )

    const report = buildCitationCheckReport({ directory: dir })
    const rows = report.rows.filter((row) => row.normalized_key === "generic authority act § 12")

    expect(rows).toHaveLength(2)
    expect(rows.every((row) => row.status === "verified")).toBe(true)
    expect(rows.every((row) => row.linked_source_ids.some((id) => id.includes("sources/generic-source.txt")))).toBe(true)
    expect(report.groups.find((group) => group.normalized_key === "generic authority act § 12")?.count).toBe(2)
  })

  test("marks citations unresolved when only an open research row carries the authority", () => {
    const dir = matter()
    writeFileSync(
      path.join(dir, "research.md"),
      researchContent([
        {
          id: "R-open",
          question: "Confirm whether Generic Rule § 4 applies.",
          tier: "open",
          sources: [],
          note: "Generic Rule § 4 still needs durable source retrieval.",
        },
      ]),
    )
    writeFileSync(path.join(dir, "issue-ledger.md"), "# Issues\n\n- Generic Rule § 4 remains open.\n")

    const report = buildCitationCheckReport({ directory: dir })
    const row = report.rows.find((item) => item.artifact.file === "issue-ledger.md")

    expect(row?.citation_text).toBe("Generic Rule § 4")
    expect(row?.status).toBe("unresolved")
    expect(row?.linked_research_ids).toContain("R-open")
  })

  test("marks generated citations unsupported when no durable support exists", () => {
    const dir = matter()
    writeFileSync(
      path.join(dir, "review-packet.md"),
      "# Review\n\nThe packet cites Sample Claimant v. Example Respondent without support.\n",
    )

    const report = buildCitationCheckReport({ directory: dir })
    const row = report.rows[0]

    expect(row.citation_text).toBe("Sample Claimant v. Example Respondent")
    expect(row.status).toBe("unsupported")
    expect(row.artifact).toMatchObject({ file: "review-packet.md", line: 3 })
  })

  test("links generated citations to verified authority-ledger entries", () => {
    const dir = matter()
    mergeAuthorityLedgerEntries(dir, [
      {
        raw: "Generic Code § 9",
        kind: "citation",
        status: "verified",
        source: { artifact: "research.md", tool: "write-matter-file", session_id: "research-session" },
        jurisdiction_hint: null,
      },
    ])
    writeFileSync(path.join(dir, "advice-memo.md"), "# Advice\n\nGeneric Code § 9 is supported.\n")

    const report = buildCitationCheckReport({ directory: dir })
    const row = report.rows[0]

    expect(row.status).toBe("verified")
    expect(row.support[0]).toMatchObject({ kind: "authority-ledger", artifact: "research.md" })
    expect(row.linked_source_ids[0]?.startsWith("auth_")).toBe(true)
  })
})

describe("writeCitationCheckReport", () => {
  test("writes a machine-readable JSON report", () => {
    const dir = matter()
    writeFileSync(path.join(dir, "advice-memo.md"), "# Advice\n\nGeneric Code § 9 is mentioned.\n")

    const report = writeCitationCheckReport({ directory: dir })
    const persisted = JSON.parse(readFileSync(path.join(dir, "citation-check-report.json"), "utf8"))

    expect(report.report_file).toBe("citation-check-report.json")
    expect(persisted.rows[0].normalized_key).toBe("generic code § 9")
    expect(persisted.rows[0].status).toBe("unsupported")
  })
})

describe("pinpoint detection", () => {
  test("flags unresolved pinpoint when parent authority exists but subsection does not", () => {
    const dir = matter()
    writeSourceText(dir, "sources/generic-source.txt", "The Generic Authority Act § 12 requires written notice.\n")
    writeFileSync(
      path.join(dir, "agreement-draft.md"),
      "# Draft\n\nThe draft relies on Generic Authority Act § 12(a)(3) for this position.\n",
    )

    const report = buildCitationCheckReport({ directory: dir })
    const row = report.rows[0]

    expect(row.status).toBe("unresolved")
    expect(row.pinpoint).toBeDefined()
    expect(row.pinpoint?.base_key).toBe("generic authority act § 12")
    expect(row.pinpoint?.suffix).toBe("(a)(3)")
    expect(row.pinpoint?.reason).toContain("pinpoint")
  })

  test("marks citation unsupported when neither pinpoint nor parent authority exists", () => {
    const dir = matter()
    writeFileSync(
      path.join(dir, "review-packet.md"),
      "# Review\n\nRelies on Unknown Statute § 99(b)(iv).\n",
    )

    const report = buildCitationCheckReport({ directory: dir })
    const row = report.rows[0]

    expect(row.status).toBe("unsupported")
    expect(row.pinpoint).toBeUndefined()
  })
})

describe("formatCitationCheckSummary", () => {
  test("renders a concise user-facing summary with mixed statuses", () => {
    const dir = matter()
    writeSourceText(dir, "sources/generic-source.txt", "The Generic Authority Act § 12 requires written notice.\n")
    mergeAuthorityLedgerEntries(dir, [
      {
        raw: "Generic Code § 9",
        kind: "citation",
        status: "verified",
        source: { artifact: "research.md", tool: "write-matter-file", session_id: "s1" },
        jurisdiction_hint: null,
      },
    ])
    writeFileSync(
      path.join(dir, "agreement-draft.md"),
      "# Draft\n\n" +
        "Generic Authority Act § 12 is supported.\n" +
        "Generic Code § 9 is also cited.\n" +
        "The draft cites Sample Claimant v. Example Respondent without support.\n" +
        "Also relies on Generic Authority Act § 12(a)(3) for this point.\n",
    )

    const report = buildCitationCheckReport({ directory: dir })
    const summary = formatCitationCheckSummary(report)

    expect(summary).toContain("verified")
    expect(summary).toContain("unresolved")
    expect(summary).toContain("unsupported")
    expect(summary).toContain("citation-check-report.json")
    expect(summary).toContain("Sample Claimant v. Example Respondent")
    expect(summary).toContain("pinpoint")
  })

  test("renders all-clear when every citation is verified", () => {
    const dir = matter()
    writeSourceText(dir, "sources/generic-source.txt", "The Generic Authority Act § 12 requires written notice.\n")
    writeFileSync(path.join(dir, "agreement-draft.md"), "# Draft\n\nGeneric Authority Act § 12 is cited.\n")

    const report = buildCitationCheckReport({ directory: dir })
    const summary = formatCitationCheckSummary(report)

    expect(summary).toContain("All citations have durable source support.")
  })

  test("renders no-citations message when artifacts have none", () => {
    const dir = matter()
    writeFileSync(path.join(dir, "agreement-draft.md"), "# Draft\n\nNo citations here.\n")

    const report = buildCitationCheckReport({ directory: dir })
    const summary = formatCitationCheckSummary(report)

    expect(summary).toContain("No citation-like references found")
  })
})

describe("case-name support matching", () => {
  // A case name, cited bare or after a signal, matches only where it begins a
  // word of the support text: at its start, or after a space, punctuation or
  // an opening parenthesis, never inside a longer name.
  const rows = (cited: string, supportText: string) => {
    const dir = matter()
    writeFileSync(path.join(dir, "draft.md"), `The court decided ${cited}, on notice.\n`)
    const index = [{ normalized_text: supportText, support: { kind: "source-artifact" as const, id: "source-artifact:s1", status: "verified" as const } }]
    return citationRowsForArtifact(dir, "draft.md", index).map(row => [row.citation_text, row.status])
  }

  test("a case name inside a longer name is unsupported", () => {
    expect(rows("Down v. Merrow", "ashdown v. merrow, 12 synth. 34 (2020)")).toEqual([["Down v. Merrow", "unsupported"]])
    expect(rows("Merrow, 12 Synth. 34", "ashmerrow, 12 synth. 34 (2020)")).toEqual([["Merrow, 12 Synth. 34", "unsupported"]])
  })

  test("a case name at the start of the support text or after punctuation is supported", () => {
    for (const supportText of ["down v. merrow, 12 synth. 34", "the leading case (down v. merrow) applies", "see notice; down v. merrow", "held:down v. merrow", "as held in down v. merrow"])
      expect(rows("Down v. Merrow", supportText)).toEqual([["Down v. Merrow", "verified"]])
    expect(rows("Merrow, 12 Synth. 34", "ashdown v. merrow, 12 synth. 34 (2020)")).toEqual([["Merrow, 12 Synth. 34", "verified"]])
  })
})
