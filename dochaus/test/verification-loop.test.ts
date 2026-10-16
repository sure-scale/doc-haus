import { describe, expect, test } from "bun:test"
import { Database } from "bun:sqlite"
import { existsSync, mkdirSync, mkdtempSync, readFileSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { computeOverall, discoverDeliverable, runVerificationGates, type GateResult } from "../lib/verification-loop"

function matter() {
  return mkdtempSync(path.join(tmpdir(), "verify-loop-"))
}

// Minimal legal.db fixture: one indexed document with the given chunk. Columns
// mirror services/ingest/src/db.ts; embedding is an empty blob since the gates
// never run vector retrieval. Used to exercise citation-anchors against a real
// SQLite index rather than a reimplemented query.
function indexChunk(dir: string, docName: string, section: string, text: string) {
  const dbDir = path.join(dir, ".dochaus")
  mkdirSync(dbDir, { recursive: true })
  const db = new Database(path.join(dbDir, "legal.db"))
  db.run(`
    CREATE TABLE chunks (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      document_id INTEGER NOT NULL,
      doc_path TEXT NOT NULL,
      doc_name TEXT NOT NULL,
      section TEXT NOT NULL,
      structural_path TEXT NOT NULL DEFAULT '',
      party TEXT NOT NULL DEFAULT '',
      chunk_index INTEGER NOT NULL,
      text TEXT NOT NULL,
      char_start INTEGER NOT NULL,
      char_end INTEGER NOT NULL,
      embedding BLOB NOT NULL,
      flagged INTEGER NOT NULL DEFAULT 0
    )
  `)
  db.run(
    "INSERT INTO chunks (document_id, doc_path, doc_name, section, chunk_index, text, char_start, char_end, embedding) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
    [1, `${docName}.docx`, docName, section, 0, text, 0, text.length, new Uint8Array()],
  )
  db.close()
}

// Empty legal.db fixture: the chunks table exists with zero rows. Mirrors the
// transient ingest state where the DB is created but no document has been
// indexed yet, so verifyQuote returns reason "no-docs".
function emptyLegalDb(dir: string) {
  const dbDir = path.join(dir, ".dochaus")
  mkdirSync(dbDir, { recursive: true })
  const db = new Database(path.join(dbDir, "legal.db"))
  db.run(`
    CREATE TABLE chunks (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      document_id INTEGER NOT NULL,
      doc_path TEXT NOT NULL,
      doc_name TEXT NOT NULL,
      section TEXT NOT NULL,
      structural_path TEXT NOT NULL DEFAULT '',
      party TEXT NOT NULL DEFAULT '',
      chunk_index INTEGER NOT NULL,
      text TEXT NOT NULL,
      char_start INTEGER NOT NULL,
      char_end INTEGER NOT NULL,
      embedding BLOB NOT NULL,
      flagged INTEGER NOT NULL DEFAULT 0
    )
  `)
  db.close()
}

function gateResult(id: string, status: GateResult["status"], summary = ""): GateResult {
  return { id: id as GateResult["id"], status, summary, findings: [] }
}

const GOOD_DELIVERABLE = [
  "# Draft Agreement",
  "",
  "## 1. Payment",
  "",
  "The fee is [insert fee amount] payable on [insert effective date].",
  "Source: payment terms per term sheet.",
  "",
].join("\n")

describe("runVerificationGates", () => {
  test("returns PASS with not-applicable gates when no deliverable is present", async () => {
    const dir = matter()
    const report = await runVerificationGates({ directory: dir })
    expect(report.status).toBe("PASS")
    expect(report.gates["citation-anchors"].status).toBe("not-applicable")
  })

  test("value-grounding FAILs on an ungrounded dollar amount", async () => {
    const dir = matter()
    writeFileSync(
      path.join(dir, "agreement.md"),
      ["# Agreement", "", "## 1. Fee", "", "The fee is $42,500.00, payable immediately.", ""].join("\n"),
    )
    const report = await runVerificationGates({ directory: dir })
    expect(report.gates["value-grounding"].status).toBe("fail")
    expect(report.gates["value-grounding"].findings.some((f) => f.detail?.includes("$42,500.00"))).toBe(true)
  })

  test("value-grounding passes when the amount is in binding-decisions.md", async () => {
    const dir = matter()
    writeFileSync(
      path.join(dir, "binding-decisions.md"),
      ["# Binding", "", "| Term | Binding value |", "|---|---|", "| Fee | $42,500.00 |", ""].join("\n"),
    )
    writeFileSync(
      path.join(dir, "agreement.md"),
      ["# Agreement", "", "## 1. Fee", "", "The fee is $42,500.00.", ""].join("\n"),
    )
    const report = await runVerificationGates({ directory: dir })
    expect(report.gates["value-grounding"].status).toBe("pass")
  })

  test("value-grounding passes when the amount is inside an [insert ...] placeholder", async () => {
    const dir = matter()
    writeFileSync(
      path.join(dir, "agreement.md"),
      ["# Agreement", "", "The fee is [insert fee amount in dollars, e.g. $50,000].", ""].join("\n"),
    )
    const report = await runVerificationGates({ directory: dir })
    expect(report.gates["value-grounding"].status).toBe("pass")
  })

  test("render-readiness FAILs on bare [___] blank", async () => {
    const dir = matter()
    writeFileSync(
      path.join(dir, "agreement.md"),
      ["# Agreement", "", "Counterparty: [___]", ""].join("\n"),
    )
    const report = await runVerificationGates({ directory: dir })
    expect(report.gates["render-readiness"].status).toBe("fail")
  })

  test("render-readiness passes on intentional [insert ...] placeholder", async () => {
    const dir = matter()
    writeFileSync(
      path.join(dir, "agreement.md"),
      ["# Agreement", "", "Counterparty: [insert counterparty name]", ""].join("\n"),
    )
    const report = await runVerificationGates({ directory: dir })
    expect(report.gates["render-readiness"].status).toBe("pass")
  })

  test("research-sources is not-applicable when no research.md exists", async () => {
    const dir = matter()
    writeFileSync(path.join(dir, "agreement.md"), GOOD_DELIVERABLE)
    const report = await runVerificationGates({ directory: dir })
    expect(report.gates["research-sources"].status).toBe("not-applicable")
  })

  test("redline-anchors is not-applicable when no pending-redlines file exists", async () => {
    const dir = matter()
    writeFileSync(path.join(dir, "agreement.md"), GOOD_DELIVERABLE)
    const report = await runVerificationGates({ directory: dir })
    expect(report.gates["redline-anchors"].status).toBe("not-applicable")
  })

  test("privilege-export is not-applicable when no redaction-log exists", async () => {
    const dir = matter()
    writeFileSync(path.join(dir, "agreement.md"), GOOD_DELIVERABLE)
    const report = await runVerificationGates({ directory: dir })
    expect(report.gates["privilege-export"].status).toBe("not-applicable")
  })

  test("gates arg restricts which gates run", async () => {
    const dir = matter()
    writeFileSync(path.join(dir, "agreement.md"), GOOD_DELIVERABLE)
    const report = await runVerificationGates({
      directory: dir,
      gates: ["render-readiness"],
    })
    expect(Object.keys(report.gates)).toEqual(["render-readiness"])
  })

  test("writes verification-report.md into the matter directory", async () => {
    const dir = matter()
    writeFileSync(path.join(dir, "agreement.md"), GOOD_DELIVERABLE)
    await runVerificationGates({ directory: dir })
    const report = await Bun.file(path.join(dir, "verification-report.md")).text()
    expect(report).toContain("# Verification report")
    expect(report).toContain("Overall:")
  })

  test("privilege-export is not-applicable when log has no export-requested entries", async () => {
    const dir = matter()
    writeFileSync(
      path.join(dir, "redaction-log.md"),
      ["# Redaction log", "", "- identifier redacted (benign)", "- term redacted (benign)", ""].join("\n"),
    )
    const report = await runVerificationGates({ directory: dir })
    expect(report.gates["privilege-export"].status).toBe("not-applicable")
  })

  test("citation-anchors passes when a bracketed quote matches the indexed DB", async () => {
    const dir = matter()
    const quoted = "the parties hereby agree to keep this arrangement confidential"
    indexChunk(dir, "source-agreement", "Section 2", `Recital: ${quoted}.`)
    writeFileSync(
      path.join(dir, "draft.md"),
      [`# Draft`, "", `Article: [${quoted}]`, ""].join("\n"),
    )
    const report = await runVerificationGates({ directory: dir, gates: ["citation-anchors"] })
    expect(report.gates["citation-anchors"].status).toBe("pass")
  })

  test("citation-anchors fails when a bracketed quote is absent from the indexed DB", async () => {
    const dir = matter()
    const indexed = "the parties hereby agree to keep this arrangement confidential"
    indexChunk(dir, "source-agreement", "Section 2", `Recital: ${indexed}.`)
    const missingQuote = "this exact wording never appears anywhere in the indexed text"
    writeFileSync(
      path.join(dir, "draft.md"),
      [`# Draft`, "", `Article: [${missingQuote}]`, ""].join("\n"),
    )
    const report = await runVerificationGates({ directory: dir, gates: ["citation-anchors"] })
    expect(report.gates["citation-anchors"].status).toBe("fail")
  })

  test("citation-anchors writes the broad citation support report", async () => {
    const dir = matter()
    writeFileSync(path.join(dir, "agreement.md"), "# Agreement\n\nGeneric Code § 9 appears without support.\n")

    const report = await runVerificationGates({ directory: dir, gates: ["citation-anchors"] })
    const citationReportFile = path.join(dir, "citation-check-report.json")
    const citationReport = JSON.parse(await Bun.file(citationReportFile).text())

    expect(report.gates["citation-anchors"].status).toBe("fail")
    expect(existsSync(citationReportFile)).toBe(true)
    expect(citationReport.rows[0]).toMatchObject({
      citation_text: "Generic Code § 9",
      normalized_key: "generic code § 9",
      status: "unsupported",
    })
  })

  test("citation-anchors is not-applicable when the DB exists but has no indexed documents", async () => {
    const dir = matter()
    emptyLegalDb(dir)
    const quoted = "this bracketed span is intended as a verbatim quotation"
    writeFileSync(
      path.join(dir, "draft.md"),
      [`# Draft`, "", `Article: [${quoted}]`, ""].join("\n"),
    )
    const report = await runVerificationGates({ directory: dir, gates: ["citation-anchors"] })
    expect(report.gates["citation-anchors"].status).toBe("not-applicable")
    expect(report.gates["citation-anchors"].summary).toContain("no documents indexed")
  })

  test("redline-anchors fails when a pending redline anchor is orphaned", async () => {
    const dir = matter()
    const targetDoc = "agreement.md"
    writeFileSync(path.join(dir, targetDoc), ["# Agreement", "", "Body text with no matching anchor here.", ""].join("\n"))
    writeFileSync(
      path.join(dir, "pending-redlines.json"),
      JSON.stringify([{ anchor: "this anchor string does not exist in the target", document: targetDoc }]),
    )
    const report = await runVerificationGates({ directory: dir, gates: ["redline-anchors"] })
    expect(report.gates["redline-anchors"].status).toBe("fail")
  })

  test("privilege-export fails on an entry flagged both privileged and export-requested", async () => {
    const dir = matter()
    writeFileSync(
      path.join(dir, "redaction-log.md"),
      [
        "# Redaction log",
        "",
        "- memo excerpt redacted (privileged, export-requested)",
        "",
      ].join("\n"),
    )
    const report = await runVerificationGates({ directory: dir, gates: ["privilege-export"] })
    expect(report.gates["privilege-export"].status).toBe("fail")
  })

  test("computeOverall returns BLOCKED when any gate threw", () => {
    const results = {
      "citation-anchors": gateResult("citation-anchors", "fail", "gate threw"),
    } as Record<GateResult["id"], GateResult>
    expect(computeOverall(results, true)).toBe("BLOCKED")
  })

  test("computeOverall returns PASS when no gate failed or threw", () => {
    const results = {
      "citation-anchors": gateResult("citation-anchors", "pass"),
      "render-readiness": gateResult("render-readiness", "not-applicable"),
    } as Record<GateResult["id"], GateResult>
    expect(computeOverall(results, false)).toBe("PASS")
  })

  test("computeOverall returns PARTIAL when a gate failed but none threw", () => {
    const results = {
      "citation-anchors": gateResult("citation-anchors", "pass"),
      "render-readiness": gateResult("render-readiness", "fail"),
    } as Record<GateResult["id"], GateResult>
    expect(computeOverall(results, false)).toBe("PARTIAL")
  })
})

describe("discoverDeliverable", () => {
  test("picks a real deliverable over a memo and meta-files when multiple candidates exist", () => {
    const dir = matter()
    writeFileSync(path.join(dir, "draft.md"), "# Draft")
    writeFileSync(path.join(dir, "agreement.md"), "# Agreement")
    writeFileSync(path.join(dir, "memo.md"), "# Memo")
    writeFileSync(path.join(dir, "coverage-report.md"), "# Coverage")
    writeFileSync(path.join(dir, "verification-report.md"), "# Verification")
    const discovered = discoverDeliverable(dir)
    expect(discovered).not.toBeNull()
    // Must be one of the deliverable shapes, never the memo or a meta-file.
    expect(["draft.md", "agreement.md"]).toContain(discovered)
  })

  test("returns null when only meta-files are present", () => {
    const dir = matter()
    writeFileSync(path.join(dir, "coverage-report.md"), "# Coverage")
    writeFileSync(path.join(dir, "verification-report.md"), "# Verification")
    writeFileSync(path.join(dir, "citation-check-report.md"), "# Citations")
    expect(discoverDeliverable(dir)).toBeNull()
  })

  test("does not pick a meta-file over a real deliverable", () => {
    const dir = matter()
    writeFileSync(path.join(dir, "agreement-draft.md"), "# Agreement Draft")
    writeFileSync(path.join(dir, "coverage-report.md"), "# Coverage")
    writeFileSync(path.join(dir, "verification-report.md"), "# Verification Report")
    expect(discoverDeliverable(dir)).toBe("agreement-draft.md")
  })
})
