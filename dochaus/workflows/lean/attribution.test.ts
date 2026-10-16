import { afterEach, expect, test } from "bun:test"
import { mkdtempSync, mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { DOMParser } from "@xmldom/xmldom"
import { strFromU8, unzipSync } from "fflate"
import { runLean } from "./runner"
import { leanFile, readLeanJson } from "./state"
import { cleanupMatters, happyPath, matter, scriptedModel } from "./test-support"

const previous = process.env.WORKSPACE_ROOT
const roots: string[] = []
afterEach(() => {
  if (previous === undefined) delete process.env.WORKSPACE_ROOT
  else process.env.WORKSPACE_ROOT = previous
  roots.splice(0).forEach(root => rmSync(root, { recursive: true, force: true }))
  cleanupMatters()
})

test("lean body, notes, native revisions and review comments use saved identity", async () => {
  const workspace = mkdtempSync(path.join(tmpdir(), "dh-lean-attribution-"))
  roots.push(workspace)
  process.env.WORKSPACE_ROOT = workspace
  mkdirSync(path.join(workspace, ".preferences"))
  writeFileSync(path.join(workspace, ".preferences/preferences.json"), JSON.stringify({ attorney: "Reviewing Lawyer", firm: "Reviewing Firm" }))
  const directory = matter()
  const state = await runLean(directory, { sessions: scriptedModel(happyPath).port })
  expect(state.status).toBe("completed")
  const parse = (bytes: Uint8Array) => new DOMParser().parseFromString(strFromU8(bytes), "application/xml")
  for (const name of ["d1/body.docx", "d1/body-applied.docx", "d1/redline.docx", "drafting-notes.docx"]) {
    const parts = unzipSync(readFileSync(leanFile(directory, name)))
    expect(parse(parts["docProps/core.xml"]).getElementsByTagNameNS("http://purl.org/dc/elements/1.1/", "creator")[0]?.textContent).toBe("Reviewing Lawyer")
    expect(parse(parts["docProps/app.xml"]).getElementsByTagNameNS("http://schemas.openxmlformats.org/officeDocument/2006/extended-properties", "Company")[0]?.textContent).toBe("Reviewing Firm")
  }
  const redline = unzipSync(readFileSync(leanFile(directory, "d1/redline.docx")))
  const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  const revisions = ["ins", "del"].flatMap(name => Array.from(parse(redline["word/document.xml"]).getElementsByTagNameNS(W, name)))
  expect(revisions.length).toBeGreaterThan(0)
  expect(new Set(revisions.map(node => node.getAttributeNS(W, "author")))).toEqual(new Set(["Reviewing Lawyer"]))
  const comments = Array.from(parse(redline["word/comments.xml"]).getElementsByTagNameNS(W, "comment"))
  expect(comments.length).toBeGreaterThan(0)
  expect(new Set(comments.map(node => node.getAttributeNS(W, "author")))).toEqual(new Set(["Reviewing Lawyer"]))
  expect(readLeanJson(directory, "d1/apply.json")).toMatchObject({ comments: { author: "Reviewing Lawyer" } })
}, 120_000)
