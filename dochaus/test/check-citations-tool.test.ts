import { afterEach, describe, expect, test } from "bun:test"
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import checkCitationsTool from "../tool/check-citations"

// Tool wrapper test: execute return shape and JSON report output.
// Behavior of the underlying citation-check library is covered by
// lib/citation-check.test.ts.

const dirs: string[] = []
afterEach(() => {
  for (const dir of dirs.splice(0)) rmSync(dir, { recursive: true, force: true })
})
const tempMatter = () => {
  const dir = mkdtempSync(path.join(tmpdir(), "check-citations-tool-"))
  dirs.push(dir)
  return dir
}

describe("check-citations tool execute", () => {
  test("returns summary output and writes JSON report on an empty matter", async () => {
    const dir = tempMatter()
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
    const dir = tempMatter()
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
