import { describe, expect, test } from "bun:test"
import { mkdtempSync, readFileSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import checkCitationsTool from "../tool/check-citations"

// Tool wrapper test: descriptor shape, execute return shape, and report file
// output. Behavior of the underlying citation-check library is covered by
// test/citation-check.test.ts; here we confirm the tool exposes the right args,
// returns the agreed shape, and writes the JSON report.

describe("check-citations tool descriptor", () => {
  test("exposes description, artifacts, and output args", () => {
    expect(checkCitationsTool.description).toMatch(/citation checker/i)
    expect(checkCitationsTool.description).toMatch(/verified.*unresolved.*unsupported/is)
    expect(checkCitationsTool.args.artifacts).toBeDefined()
    expect(checkCitationsTool.args.output).toBeDefined()
  })

  test("description contains no fixture-specific legal values", () => {
    expect(checkCitationsTool.description).not.toMatch(/\$\d/)
    expect(checkCitationsTool.description).not.toMatch(/\b20\d{2}\b/)
  })
})

describe("check-citations tool execute", () => {
  test("returns summary output and writes JSON report on an empty matter", async () => {
    const dir = mkdtempSync(path.join(tmpdir(), "check-citations-tool-"))
    const result = await checkCitationsTool.execute({}, { directory: dir } as any)
    expect(typeof result).toBe("object")
    expect(result).not.toBeNull()
    const r = result as { title: string; output: string; metadata: Record<string, unknown> }
    expect(r.title).toMatch(/Citation check:/)
    expect(r.output).toMatch(/No citation-like references found/)
    expect(r.metadata.total).toBe(0)
    expect(r.metadata.reportFile).toBe("citation-check-report.json")
    const persisted = JSON.parse(readFileSync(path.join(dir, "citation-check-report.json"), "utf8"))
    expect(persisted.schema_version).toBe(1)
  })

  test("detects verified and unsupported citations from matter artifacts", async () => {
    const dir = mkdtempSync(path.join(tmpdir(), "check-citations-tool-"))
    writeFileSync(
      path.join(dir, "source-map.md"),
      "# Source Map\n\nThe Generic Authority Act § 12 requires written notice.\n",
    )
    writeFileSync(
      path.join(dir, "agreement-draft.md"),
      "# Draft\n\n" +
        "The draft follows Generic Authority Act § 12.\n" +
        "Also cites Example Claimant v. Sample Respondent without support.\n",
    )

    const result = await checkCitationsTool.execute({}, { directory: dir } as any)
    const r = result as { title: string; output: string; metadata: Record<string, unknown> }
    expect(r.metadata.total).toBe(2)
    expect(r.metadata.verified).toBe(1)
    expect(r.metadata.unsupported).toBe(1)
    expect(r.output).toContain("verified")
    expect(r.output).toContain("unsupported")
  })
})
