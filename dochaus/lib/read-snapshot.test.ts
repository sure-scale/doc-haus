import { describe, expect, test } from "bun:test"
import { mkdtempSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { readDurableJson, withDurableJsonReadSnapshot } from "./durable-json"
import { sha256File, withFileHashReadSnapshot } from "./lineage"

describe("observational read snapshots", () => {
  test("reuse file hashes and durable JSON only inside one explicit read scope", () => {
    const directory = mkdtempSync(path.join(tmpdir(), "read-snapshot-"))
    try {
      const text = path.join(directory, "artifact.md")
      const json = path.join(directory, "state.json")
      writeFileSync(text, "stable artifact")
      writeFileSync(json, JSON.stringify({ state: "stable" }))

      const hashes = withFileHashReadSnapshot(() => [sha256File(text), sha256File(text)])
      const values = withDurableJsonReadSnapshot(() => [readDurableJson(json), readDurableJson(json)])

      expect(hashes[0]).toBe(hashes[1])
      expect(values[0]).toBe(values[1])
      expect(sha256File(text)).not.toBe(hashes[0])
      expect(readDurableJson(json)).not.toBe(values[0])
    } finally {
      rmSync(directory, { recursive: true, force: true })
    }
  })
})
