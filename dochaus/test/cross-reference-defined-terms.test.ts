import { describe, expect, test } from "bun:test"
import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import crossReferenceDefinedTerms from "../tool/cross-reference-defined-terms"

describe("cross-reference-defined-terms tool", () => {
  function matter() {
    const dir = mkdtempSync(path.join(tmpdir(), "cross-reference-defined-terms-"))
    mkdirSync(path.join(dir, "source-chunks"))
    writeFileSync(path.join(dir, "source-chunks", "source-a-c001.txt"), "The Northern Corridor applies to this matter.\n")
    writeFileSync(path.join(dir, "source-chunks", "source-b-c001.txt"), "This source does not repeat that named reference.\n")
    return dir
  }

  test("normalizes absolute output paths to a matter-local report filename", async () => {
    const dir = matter()
    const output = path.join(dir, "defined-terms-cross-reference.md")

    const result = await crossReferenceDefinedTerms.execute({ output }, { directory: dir } as never)

    expect(result).toContain("defined-terms-cross-reference.md")
    expect(existsSync(path.join(dir, "defined-terms-cross-reference.md"))).toBe(true)
    expect(existsSync(path.join(dir, output))).toBe(false)
    rmSync(dir, { recursive: true, force: true })
  })

  test("uses the durable deliverable record without document-family filename guesses", async () => {
    const dir = matter()
    writeFileSync(path.join(dir, "formal-response.md"), "# Formal Response\n\nThe Eastern Review Panel should continue.\n")
    writeFileSync(path.join(dir, ".dochaus-deliverable.json"), JSON.stringify({ assembledBody: "formal-response.md" }))

    await crossReferenceDefinedTerms.execute({}, { directory: dir } as never)

    expect(readFileSync(path.join(dir, "defined-terms-cross-reference.md"), "utf8")).toContain("Eastern Review Panel")
    rmSync(dir, { recursive: true, force: true })
  })

  test("cross-references quoted labels without a document-family suffix", async () => {
    const dir = matter()
    writeFileSync(path.join(dir, "source-chunks", "source-a-c002.txt"), 'The process uses the “Blue Route”.\n')

    await crossReferenceDefinedTerms.execute({}, { directory: dir } as never)

    expect(readFileSync(path.join(dir, "defined-terms-cross-reference.md"), "utf8")).toContain("Blue Route")
    rmSync(dir, { recursive: true, force: true })
  })

  test("uses an explicit body filename instead of a larger unrelated markdown file", async () => {
    const dir = matter()
    writeFileSync(path.join(dir, "instrument-body.md"), "# Instrument\n\nThe Alpha Instrument governs.\n")
    writeFileSync(path.join(dir, "unrelated-large.md"), "# Unrelated Mirage\n\n" + "Unrelated Mirage ".repeat(300))

    await crossReferenceDefinedTerms.execute({ body: "instrument-body.md" }, { directory: dir } as never)

    const report = readFileSync(path.join(dir, "defined-terms-cross-reference.md"), "utf8")
    expect(report).toContain("Alpha Instrument")
    expect(report).not.toContain("Unrelated Mirage")
    rmSync(dir, { recursive: true, force: true })
  })
})
