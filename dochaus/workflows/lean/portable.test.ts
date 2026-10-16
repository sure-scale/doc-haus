import { afterEach, expect, test } from "bun:test"
import { mkdtempSync, readdirSync, readFileSync, realpathSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { readLeanDocuments } from "./deliverables"
import { enqueueLeanRequest } from "./requests"
import { restoreLeanStage } from "./restore-stage"
import { runLean } from "./runner"
import { LEAN_STAGE_AGENTS } from "./stages"
import { leanDir, matterRelative, readLeanState } from "./state"
import { happyLeanReplies, leanFixtureMatter, scriptedLeanModel } from "./test-support"

const directories: string[] = []
afterEach(() => directories.splice(0).forEach(directory => rmSync(directory, { recursive: true, force: true })))

test("a lean matter copied to another folder restores and resumes there and records no machine path", async () => {
  const original = leanFixtureMatter("lean-portable-")
  directories.push(original)
  // An engine error naming a matter file is runtime text the run records.
  const failing = scriptedLeanModel({
    ...happyLeanReplies(),
    [LEAN_STAGE_AGENTS.review]: () => {
      throw new Error(`engine refused ${path.join(realpathSync(original), ".dochaus", "lean", "review.md")}`)
    },
  })
  const stopped = await runLean(original, { sessions: failing.port })
  expect(stopped.status).toBe("failed")
  expect(stopped.stop_reason).toBe("d1.review: stopped by a runtime error: engine refused .dochaus/lean/review.md")

  enqueueLeanRequest(original, { kind: "continue" })
  expect((await runLean(original, { sessions: scriptedLeanModel().port })).status).toBe("completed")

  const parent = mkdtempSync(path.join(tmpdir(), "lean-portable-copy-"))
  directories.push(parent)
  const copy = path.join(parent, "moved matter")
  expect(Bun.spawnSync(["cp", "-Rp", original, copy]).exitCode).toBe(0)

  restoreLeanStage(copy, "d1.changes")
  const resume = scriptedLeanModel({ [LEAN_STAGE_AGENTS.finalize]: happyLeanReplies()[LEAN_STAGE_AGENTS.finalize] })
  const resumed = await runLean(copy, { sessions: resume.port })
  expect(resumed.status).toBe("completed")
  expect(resume.agents()).toEqual(["lean-checklist-extract", LEAN_STAGE_AGENTS["change-review"], LEAN_STAGE_AGENTS.finalize, LEAN_STAGE_AGENTS["package-review"]])
  expect(readLeanState(copy)?.stages.gate?.status).toBe("done")
  expect(readLeanDocuments(copy)?.status).toBe("completed")

  const roots = [...new Set([original, copy].flatMap(root => [root, realpathSync(root)]))]
  const leaks = files(leanDir(copy)).filter(file => {
    const text = readFileSync(file, "latin1")
    return roots.some(root => text.includes(root))
  })
  expect(leaks.map(file => path.relative(copy, file))).toEqual([])
}, 120_000)

test("runtime error text names matter files relative to the matter folder, through either path form", () => {
  const directory = leanFixtureMatter("lean-portable-")
  directories.push(directory)
  const text = `open '${path.join(realpathSync(directory), ".dochaus", "lean", "gate.json")}'; scandir '${directory}'`
  expect(matterRelative(directory, text)).toBe("open '.dochaus/lean/gate.json'; scandir '.'")
})

function files(directory: string): string[] {
  return readdirSync(directory, { withFileTypes: true }).flatMap(entry =>
    entry.isDirectory() ? files(path.join(directory, entry.name)) : [path.join(directory, entry.name)])
}
