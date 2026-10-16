import { describe, expect, test } from "bun:test"
import { mkdtempSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import verifyTool from "../tool/verify"

// Task 7: verify-tool descriptor + return-shape test. Behavior of the underlying
// gates is covered by test/verification-loop.test.ts (Task 6); here we only
// confirm the wrapper exposes the right args and returns the agreed shape.

describe("verify tool descriptor", () => {
  test("exposes description, deliverable, gates, and output args", () => {
    expect(verifyTool.description).toMatch(/verification gates/i)
    expect(verifyTool.description).toMatch(/PASS\s*\/\s*PARTIAL\s*\/\s*BLOCKED/)
    expect(verifyTool.args.deliverable).toBeDefined()
    expect(verifyTool.args.gates).toBeDefined()
    expect(verifyTool.args.output).toBeDefined()
  })

  test("description contains no fixture-specific legal values", () => {
    // Transfer check: the description must not bake in benchmark artifacts.
    expect(verifyTool.description).not.toMatch(/\$\d/)
    expect(verifyTool.description).not.toMatch(/\b20\d{2}\b/)
  })
})

describe("verify tool execute", () => {
  test("returns PASS title/output/metadata on an empty matter and writes the report", async () => {
    const dir = mkdtempSync(path.join(tmpdir(), "verify-tool-"))
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
    const dir = mkdtempSync(path.join(tmpdir(), "verify-tool-"))
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
