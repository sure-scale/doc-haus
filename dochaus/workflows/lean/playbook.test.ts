import { afterEach, expect, test } from "bun:test"
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { boundPlaybook, boundPlaybookContext } from "../../lib/playbook-source"
import { verifyQuote } from "../../lib/quote-verify"
import readDocument from "../../tool/read-document"
import { evidencePrompt, checkEvidence } from "./evidence"
import { stagePrompt } from "./prompts"
import { packageReviewPrompt } from "./package-review"
import { researchPrompt } from "./research"
import { matterSourceIDs } from "./quotes"
import { stageInputs } from "./stages"
import { hashStageInputs, initLeanState } from "./state"

const roots: string[] = []
const workspace = process.env.WORKSPACE_ROOT
afterEach(() => {
  roots.splice(0).forEach(root => rmSync(root, { recursive: true, force: true }))
  if (workspace === undefined) delete process.env.WORKSPACE_ROOT
  else process.env.WORKSPACE_ROOT = workspace
})

function fixture() {
  const root = mkdtempSync(path.join(tmpdir(), "playbook-drafting-"))
  roots.push(root)
  process.env.WORKSPACE_ROOT = root
  const directory = path.join(root, "matter")
  const library = path.join(root, ".playbooks/playbook-workshop")
  mkdirSync(directory, { recursive: true })
  mkdirSync(library, { recursive: true })
  const text = "---\nname: playbook-workshop\ndescription: Workshop positions\n---\n\n## Cancellation\nPreferred: Cancellation notice is eight working days.\nFallback: Three working days only when a substitute is confirmed.\n"
  const file = path.join(library, "SKILL.md")
  writeFileSync(file, text)
  writeFileSync(path.join(directory, "matter.json"), JSON.stringify({ id: "matter", playbook: "playbook-workshop" }))
  initLeanState(directory, { request: "Propose changes for the organiser using the assigned guidance. Retain the agreed catering arrangement.", ownerSessionID: "ses_owner", model: { providerID: "test", modelID: "scripted" } })
  return { directory, text, file }
}

test("the selected library source is retrievable and quoted exactly without an indexed copy", async () => {
  const { directory, text } = fixture()
  expect(boundPlaybook(directory)?.text).toBe(text)
  expect(matterSourceIDs(directory)).toContain("playbook:playbook-workshop")
  const result = await readDocument.execute({ document: "playbook:playbook-workshop" }, { directory } as Parameters<typeof readDocument.execute>[1])
  expect(typeof result === "object" && result.output).toContain(text)
  expect(verifyQuote({ directory, document: "playbook:playbook-workshop", quote: "Three working days only when a substitute is confirmed." }).found).toBe(true)
  expect(verifyQuote({ directory, quote: "Three working days only when a substitute is confirmed." }).hits[0]?.doc).toBe("playbook:playbook-workshop")
  expect(verifyQuote({ directory, document: "playbook:playbook-other", quote: "Cancellation notice is eight working days." }).found).toBe(false)
  expect(verifyQuote({ directory, document: "playbook:playbook-workshop", quote: "Three working days are always acceptable." }).found).toBe(false)
  expect(checkEvidence(directory, "spine", { entries: [{ kind: "recommendation", detail: "Conditional notice fallback", evidence: [{ source_id: "playbook:playbook-workshop", quote: "Three working days only when a substitute is confirmed." }], limitations: [] }], limitations: [] })).toEqual([])
})

test("binding reaches early stages, document review and final package; changes invalidate dependent inputs", () => {
  const { directory, file } = fixture()
  for (const prompt of [evidencePrompt(directory, "discrepancy"), evidencePrompt(directory, "spine"), researchPrompt(directory, "intake"), stagePrompt(directory, "d1.draft"), stagePrompt(directory, "d1.review-independent"), stagePrompt(directory, "d1.changes"), stagePrompt(directory, "finalize"), packageReviewPrompt(directory)]) expect(prompt).toContain("playbook:playbook-workshop")
  const before = hashStageInputs("d1.review-independent", "test", stageInputs(directory, "d1.review-independent"))
  writeFileSync(file, "The cancellation notice is now ten working days.")
  expect(hashStageInputs("d1.review-independent", "test", stageInputs(directory, "d1.review-independent"))).not.toBe(before)
  expect(verifyQuote({ directory, document: "playbook:playbook-workshop", quote: "Cancellation notice is eight working days." }).found).toBe(false)
  writeFileSync(path.join(directory, "matter.json"), JSON.stringify({ id: "matter" }))
  expect(boundPlaybookContext(directory)).toBe("")
  expect(matterSourceIDs(directory)).not.toContain("playbook:playbook-workshop")
})

test("missing and invalid bindings are disclosed without inferring policy or importing another library entry", () => {
  const { directory } = fixture()
  writeFileSync(path.join(directory, "matter.json"), JSON.stringify({ playbook: "playbook-missing" }))
  expect(boundPlaybookContext(directory)).toContain("unavailable")
  expect(matterSourceIDs(directory)).toEqual([])
  writeFileSync(path.join(directory, "matter.json"), JSON.stringify({ playbook: "../playbook-workshop" }))
  expect(boundPlaybook(directory)?.text).toBeUndefined()
})
