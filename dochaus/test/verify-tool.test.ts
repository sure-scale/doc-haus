import { afterEach, describe, expect, test } from "bun:test"
import { mkdtempSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import verifyTool from "../tool/verify"

// Verify-tool wrapper execution and return-shape test. Behavior of the
// underlying gates is covered by test/verification-loop.test.ts.

const dirs: string[] = []
afterEach(() => {
  for (const dir of dirs.splice(0)) rmSync(dir, { recursive: true, force: true })
})
const tempMatter = () => {
  const dir = mkdtempSync(path.join(tmpdir(), "verify-tool-"))
  dirs.push(dir)
  return dir
}

describe("verify tool execute", () => {
  test("returns PASS title/output/metadata on an empty matter and writes the report", async () => {
    const dir = tempMatter()
    const result = await verifyTool.execute({}, { directory: dir } as any)
    expect(typeof result).toBe("object")
    expect(result).not.toBeNull()
    const r = result as { title: string; output: string; metadata: Record<string, unknown> }
    expect(r.title).toMatch(/^Verification: (PASS|PARTIAL|BLOCKED)$/)
    expect(r.output).toMatch(/^(PASS|PARTIAL|BLOCKED):/)
    expect(r.metadata.status).toBe(r.title.split(": ")[1])
    expect(typeof r.metadata.reportFile).toBe("string")
    expect(r.metadata.gates).toBeTypeOf("object")
    const report = await Bun.file(path.join(dir, "verification-report.md")).text()
    expect(report).toContain("# Verification report")
  })

  test("gates arg restricts the metadata.gates keys", async () => {
    const dir = tempMatter()
    writeFileSync(
      path.join(dir, "agreement.md"),
      ["# Agreement", "", "Counterparty: [insert counterparty name]", ""].join("\n"),
    )
    const result = await verifyTool.execute(
      { gates: ["render-readiness"] },
      { directory: dir } as any,
    )
    const r = result as { metadata: { gates: Record<string, string> } }
    expect(Object.keys(r.metadata.gates)).toEqual(["render-readiness"])
    expect(r.metadata.gates["render-readiness"]).toBe("pass")
  })
})
