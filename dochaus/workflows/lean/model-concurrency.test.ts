import { expect, test } from "bun:test"
import { mkdtempSync, mkdirSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { parallel } from "./checklist-flow"

test("workflow pool follows the saved installation limit", async () => {
  const root = mkdtempSync(path.join(tmpdir(), "dochaus-concurrency-"))
  const previousLegal = process.env.DOCHAUS_DIR
  const previousRuntime = process.env.DOCHAUS_RUNTIME_DIR
  try {
    const legal = path.join(root, "legal")
    mkdirSync(legal)
    process.env.DOCHAUS_DIR = legal
    process.env.DOCHAUS_RUNTIME_DIR = path.join(root, "runtime")
    for (const [saved, limit] of [[2, 3], [7, 5]]) {
      writeFileSync(path.join(legal, "opencode.json"), JSON.stringify({ model_concurrency: saved }))
      let active = 0
      let peak = 0
      const finished: number[] = []
      await parallel(Array.from({ length: 9 }, (_, index) => index), async job => {
        active++
        peak = Math.max(peak, active)
        await Bun.sleep(10)
        finished.push(job)
        active--
      })
      expect(peak).toBe(limit)
      expect(finished.sort((a, b) => a - b)).toEqual([0, 1, 2, 3, 4, 5, 6, 7, 8])
    }
  } finally {
    if (previousLegal === undefined) delete process.env.DOCHAUS_DIR
    else process.env.DOCHAUS_DIR = previousLegal
    if (previousRuntime === undefined) delete process.env.DOCHAUS_RUNTIME_DIR
    else process.env.DOCHAUS_RUNTIME_DIR = previousRuntime
    rmSync(root, { recursive: true, force: true })
  }
})
