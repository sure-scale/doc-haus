import { describe, expect, test } from "bun:test"
import { existsSync, mkdirSync, mkdtempSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import writeMatterFile from "../tool/write-matter-file"

function matter() {
  const dir = mkdtempSync(path.join(tmpdir(), "matter-markdown-name-"))
  mkdirSync(dir, { recursive: true })
  return dir
}

describe("matter markdown artifact names", () => {
  test("writer rejects a tool name as the target artifact", async () => {
    const dir = matter()
    const result = await writeMatterFile.execute({ name: "write-matter-file", content: "body" }, { directory: dir } as never)
    expect(String(result)).toContain("Refusing tool-name file")
    expect(existsSync(path.join(dir, "write-matter-file.md"))).toBe(false)
  })

  test("writer still accepts ordinary markdown artifact basenames", async () => {
    const dir = matter()
    await writeMatterFile.execute({ name: "drafting-strategy.md", content: "body" }, { directory: dir } as never)
    expect(existsSync(path.join(dir, "drafting-strategy.md"))).toBe(true)
  })

  test("writer rejects a JSON basename", async () => {
    const dir = matter()
    const result = await writeMatterFile.execute(
      { name: "arbitrary.json", content: "{}" },
      { directory: dir } as never,
    )
    expect(String(result)).toContain("is not agent-writable")
    expect(existsSync(path.join(dir, "arbitrary.json"))).toBe(false)
    expect(existsSync(path.join(dir, "arbitrary.json.md"))).toBe(false)
  })
})
