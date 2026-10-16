import { afterEach, expect, test } from "bun:test"
import { mkdtempSync, mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { pendingRedlinesForDoc } from "../lib/redlines"
import { renderMarkdownToDocx } from "../lib/render"
import trackedChanges from "../tool/tracked-changes"
import redline from "../tool/redline"

const previous = process.env.WORKSPACE_ROOT
const roots: string[] = []
afterEach(() => {
  if (previous === undefined) delete process.env.WORKSPACE_ROOT
  else process.env.WORKSPACE_ROOT = previous
  roots.splice(0).forEach(root => rmSync(root, { recursive: true, force: true }))
})

test("editing tools default new proposals to saved identity and preserve an explicit author", async () => {
  const root = mkdtempSync(path.join(tmpdir(), "dh-edit-attribution-"))
  roots.push(root)
  process.env.WORKSPACE_ROOT = root
  mkdirSync(path.join(root, ".preferences"))
  const preferences = path.join(root, ".preferences/preferences.json")
  writeFileSync(preferences, JSON.stringify({ attorney: "Saved Lawyer", firm: "Saved Firm" }))
  const source = (await renderMarkdownToDocx("Current review period is five days.")).bytes
  for (const [name, target] of [
    ["phrase", { find: "five days", replace: "seven days" }],
    ["clause", { clause: "Current review period", replacement: "Current review period is seven days." }],
  ] as const) {
    const document = path.join(root, `${name}.docx`)
    writeFileSync(document, source)
    const context = { sessionID: "identity-test", messageID: "identity-test", agent: "redliner", directory: root, worktree: root, abort: new AbortController().signal, metadata() {}, async ask() {} }
    const args = { document: path.basename(document), comment: "Reflects the instructed review period.", ...target }
    // Each tool keeps its own argument contract; the fixtures exercise both.
    const propose = (author?: string) => name === "phrase"
      ? trackedChanges.execute({ ...args, find: "five days", replace: "seven days", author }, context)
      : redline.execute({ ...args, clause: "Current review period", replacement: "Current review period is seven days.", author }, context)
    await propose()
    expect(pendingRedlinesForDoc(root, document)[0]?.author).toBe("Saved Lawyer")
    await propose("Explicit Reviewer")
    expect(pendingRedlinesForDoc(root, document)[0]?.author).toBe("Explicit Reviewer")
    writeFileSync(preferences, JSON.stringify({ attorney: "", firm: "Saved Firm" }))
    await propose()
    expect(pendingRedlinesForDoc(root, document)[0]?.author).toBe("Saved Firm")
    writeFileSync(preferences, "{}")
    await propose()
    expect(pendingRedlinesForDoc(root, document)[0]?.author).toBe("doc.haus")
    expect(readFileSync(document)).toEqual(Buffer.from(source))
    writeFileSync(preferences, JSON.stringify({ attorney: "Saved Lawyer", firm: "Saved Firm" }))
  }
})
