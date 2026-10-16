import { afterEach, expect, test } from "bun:test"
import { readFileSync } from "node:fs"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { docxodus } from "../../lib/docxodus"
import { renderMarkdownToDocx } from "../../lib/render"
import { readChangeList, type ApplyResult } from "./checks"
import { finalGate } from "./gate"
import { packageReviewPrompt } from "./package-review"
import { appliedChangeEvidence, exportAppliedChangeComments } from "./review-evidence"
import { runLean } from "./runner"
import { runApply } from "./apply-stage"
import { stagePrompt } from "./prompts"
import { parsePlan, writeLeanPlan } from "./plan"
import { acceptStageReply, paragraphRows } from "./stages"
import { leanFile, readLeanJson, readLeanState, replaceStageOpenItems, updateLeanStage, writeLeanBytes, writeLeanText } from "./state"
import { AGENT, APPLIED, DRAFT, DRAFTED, SOURCE, SOURCE_TEXT, SUPPORT, cleanupMatters, drafter, happyPath, matter, pidOf, scriptedModel } from "./test-support"

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
afterEach(cleanupMatters)

// Existing saved change lists can contain overlapping paragraph proposals.
// Exercise their legacy apply/notes boundary directly: new runLean batches must
// first consolidate replacements, so they cannot reach these merge outcomes.
async function applyLegacyList(directory: string, model: ReturnType<typeof scriptedModel>, draft = DRAFT) {
  const plan = parsePlan(`Acting for: not stated\n${draft}`)
  expect(plan.issues).toEqual([])
  writeLeanPlan(directory, plan.value!.documents)
  const rendered = await acceptStageReply(directory, "d1.draft", draft)
  expect(rendered.issues).toEqual([])
  const changes = await acceptStageReply(directory, "d1.changes", await model.port.prompt({
    directory, sessionID: "ses_legacy_changes", agent: AGENT.changes, text: stagePrompt(directory, "d1.changes"),
  }))
  expect(changes.issues).toEqual([])
  replaceStageOpenItems(directory, "changes", "d1", changes.openItems)
  const applied = await runApply(directory, readLeanState(directory)!, { sessions: model.port }, "d1")
  expect(applied.status).toBe("done")
  if (applied.status !== "done") throw new Error(applied.reason)
  replaceStageOpenItems(directory, "apply", "d1", applied.openItems)
  updateLeanStage(directory, readLeanState(directory)!, "d1.apply", { status: "done", outputs: applied.outputs })
  const notes = await acceptStageReply(directory, "finalize", await model.port.prompt({
    directory, sessionID: "ses_legacy_notes", agent: AGENT.finalize, text: stagePrompt(directory, "finalize"),
  }))
  expect(notes.issues).toEqual([])
  const gate = finalGate(directory)
  writeLeanText(directory, "gate.json", `${JSON.stringify(gate, null, 2)}\n`)
  expect(gate.status).toBe("pass")
  return applied
}

test("additive merges disclose proposed explanations for review instead of calling them explanations of final wording", async () => {
  const directory = matter(`${SOURCE_TEXT} The supplier shall promptly deliver the goods.`)
  const merged = "The supplier shall promptly deliver the goods within thirty days of the order date."
  const model = scriptedModel({
    ...happyPath,
    [AGENT.changes]: prompt => JSON.stringify({ amendments: [
      { pid: pidOf(prompt, DRAFTED), updated: APPLIED, reason: "State the source delivery period.", sources: SUPPORT },
      { pid: pidOf(prompt, DRAFTED), updated: "The supplier shall promptly deliver the goods.", reason: "State the source delivery promptness.", sources: [{ source_id: SOURCE, quote: "The supplier shall promptly deliver the goods." }] },
    ], questions: [] }),
    [AGENT.merge]: () => `<paragraph>${merged}</paragraph>`,
  })
  expect((await applyLegacyList(directory, model)).status).toBe("done")
  const result = readLeanJson(directory, "d1/apply.json") as ApplyResult
  expect(readChangeList(directory, "d1")!.withheld ?? []).toEqual([])
  expect(result.unapplied).toEqual([])
  expect(result.applied.map(row => row.mode)).toEqual(["code", "merged"])
  expect(result.comments?.attached).toEqual([])
  expect(result.comments?.unresolved.map(row => row.id)).toEqual(["a1", "a2"])
  expect(result.comments?.unresolved.every(row => row.reason.includes("explanation needs review"))).toBe(true)
  const parts = unzipSync(readFileSync(leanFile(directory, "d1/redline.docx")))
  const doc = new DOMParser().parseFromString(strFromU8(parts["word/document.xml"]!), "application/xml")
  expect(parts["word/comments.xml"]).toBeUndefined()
  expect(doc.getElementsByTagNameNS(W, "commentRangeStart").length).toBe(0)
  expect(doc.getElementsByTagNameNS(W, "ins").length).toBeGreaterThan(0)
  const clean = unzipSync(readFileSync(leanFile(directory, "d1/body-applied.docx")))
  expect(clean["word/comments.xml"]).toBeUndefined()
  expect(strFromU8(clean["word/document.xml"]!)).not.toContain("State the source delivery period")
  const evidence = appliedChangeEvidence(directory, "d1")
  expect(evidence[1]!.native_comment).toMatchObject({ status: "unresolved", text: result.comments!.unresolved[1]!.text })
  for (const prompt of [model.prompts.find(row => row.agent === AGENT.finalize)!.text, packageReviewPrompt(directory)]) {
    expect(prompt).toContain('"status": "unresolved"')
    expect(prompt).toContain(result.comments!.unresolved[1]!.text.replaceAll("\n", "\\n"))
  }
  expect(finalGate(directory).disclosed.filter(row => row.kind === "omission")).toHaveLength(2)
}, 120_000)

test("a final redline anchor mismatch creates a disclosed omission, never an unapplied amendment or a nearby comment", async () => {
  const directory = matter()
  expect((await runLean(directory, { sessions: scriptedModel(happyPath).port })).status).toBe("completed")
  const amendments = readChangeList(directory, "d1")!.amendments
  const result = readLeanJson(directory, "d1/apply.json") as ApplyResult
  const original = readFileSync(leanFile(directory, "d1/body.docx"))
  const valid = exportAppliedChangeComments(readFileSync(leanFile(directory, "d1/redline.docx")), readFileSync(leanFile(directory, "d1/body-applied.docx")), "d1", amendments, result.applied)
  expect(valid.comments.unresolved).toEqual([])
  expect(valid.comments.attached).toHaveLength(1)
  const stale = (await renderMarkdownToDocx("# Supply note\n\n## Delivery\n\nDifferent current wording.")).bytes
  const mismatched = await (await docxodus()).compareDocuments(original, stale)
  const missing = exportAppliedChangeComments(mismatched, readFileSync(leanFile(directory, "d1/body-applied.docx")), "d1", amendments, result.applied)
  expect(missing.comments.attached).toEqual([])
  expect(missing.comments.unresolved).toHaveLength(1)
  expect(missing.openItems[0]).toMatchObject({ kind: "omission", ref: { document: "d1", amendment_id: "a1" } })
  expect(missing.openItems[0]!.detail).toContain("body change remains applied")
  expect(unzipSync(missing.bytes)["word/comments.xml"]).toBeUndefined()
  replaceStageOpenItems(directory, "apply", "d1", missing.openItems)
  expect(finalGate(directory).blocking.map(row => row.id)).toContain(missing.openItems[0]!.id)
  writeLeanText(directory, "drafting-notes.md", `# Drafting notes\n\n${missing.openItems[0]!.id}: ${missing.openItems[0]!.detail}`)
  expect(finalGate(directory).disclosed.map(row => row.id)).toContain(missing.openItems[0]!.id)
}, 120_000)


test("native final paragraph mapping anchors redline comments across tabs, line breaks and empty paragraphs", async () => {
  const directory = matter()
  const baseline = (await renderMarkdownToDocx("# Supply note\n\n## Delivery\n\nBefore.\n\nAfter.")).bytes
  const parts = unzipSync(baseline)
  const xml = new DOMParser().parseFromString(strFromU8(parts["word/document.xml"]!), "application/xml")
  const target = Array.from(xml.getElementsByTagNameNS(W, "t")).find(node => node.textContent === "Before.")!
  const run = target.parentNode!
  target.textContent = "Before"
  for (const [name, text] of [["tab", ""], ["t", "tab."], ["br", ""], ["t", "Next line."]] as const) {
    const node = xml.createElementNS(W, `w:${name}`)
    if (text) node.textContent = text
    run.appendChild(node)
  }
  const paragraph = run.parentNode!
  paragraph.parentNode!.insertBefore(xml.createElementNS(W, "w:p"), paragraph)
  const final = zipSync({ ...parts, "word/document.xml": strToU8(new XMLSerializer().serializeToString(xml)) })
  writeLeanBytes(directory, "d1/body-applied.docx", final)
  const paragraphs = await paragraphRows(leanFile(directory, "d1/body-applied.docx"))
  const located = paragraphs.find(row => row.text.startsWith("Before"))!
  // Docxodus matches omit tabs and breaks; native Word anchors retain them.
  expect(located.text).toBe("Beforetab.Next line.")
  const redline = await (await docxodus()).compareDocuments(baseline, readFileSync(leanFile(directory, "d1/body-applied.docx")))
  const exported = exportAppliedChangeComments(redline, readFileSync(leanFile(directory, "d1/body-applied.docx")), "d1", [{
    id: "a1", pid: located.pid, existing_sha256: located.sha256, updated: "Before tab. Next line.",
    reason: "Explain the current wording.", sources: [],
  }], [{ id: "a1", pid: located.pid, mode: "merged" }])
  expect(exported.comments.unresolved).toEqual([])
  expect(exported.comments.attached).toEqual([expect.objectContaining({
    paragraph_index: paragraphs.indexOf(located), paragraph_text: "Before\ttab.\nNext line.",
  })])
  expect(strFromU8(unzipSync(exported.bytes)["word/comments.xml"]!)).toContain("Explain the current wording.")
}, 120_000)

test("an unapplied amendment stays in open items and receives no applied Word comment", async () => {
  const directory = matter(`${SOURCE_TEXT} The supplier shall promptly deliver the goods.`)
  const model = scriptedModel({
    ...happyPath,
    [AGENT.changes]: prompt => JSON.stringify({ amendments: [
      { pid: pidOf(prompt, DRAFTED), updated: APPLIED, reason: "State the delivery period.", sources: SUPPORT },
      { pid: pidOf(prompt, DRAFTED), updated: "The supplier shall promptly deliver the goods.", reason: "Require prompt delivery.", sources: [{ source_id: SOURCE, quote: "The supplier shall promptly deliver the goods." }] },
    ], questions: [] }),
    [AGENT.merge]: () => `<paragraph>${APPLIED}</paragraph>`,
  })
  expect((await applyLegacyList(directory, model)).status).toBe("done")
  const result = readLeanJson(directory, "d1/apply.json") as ApplyResult
  expect(result.applied.map(row => row.id)).toEqual(["a1"])
  expect(result.unapplied.map(row => row.id)).toEqual(["a2"])
  expect(result.comments?.attached.map(row => row.id)).toEqual(["a1"])
  const comments = strFromU8(unzipSync(readFileSync(leanFile(directory, "d1/redline.docx")))["word/comments.xml"]!)
  expect(comments).toContain("State the delivery period.")
  expect(comments).not.toContain("Require prompt delivery.")
  expect(finalGate(directory).disclosed).toContainEqual(expect.objectContaining({ kind: "unapplied_amendment", ref: expect.objectContaining({ amendment_id: "a2" }) }))
}, 120_000)


test("overlapping replacement attaches only the explanation matching the final wording", async () => {
  const initial = "The supplier shall deliver the goods within ten days of the order date."
  const earlier = "The supplier shall deliver the goods within twenty days of the order date."
  const directory = matter(`${SOURCE_TEXT} ${earlier}`)
  const model = scriptedModel({
    ...happyPath,
    [AGENT.draft]: drafter(DRAFT.replace(DRAFTED, initial)),
    [AGENT.review]: () => "The delivery period needs review.",
    [AGENT.changes]: prompt => JSON.stringify({ amendments: [
      { pid: pidOf(prompt, initial), updated: earlier, reason: "Allows a twenty day delivery period.", sources: [{ source_id: SOURCE, quote: "within twenty days of the order date" }] },
      { pid: pidOf(prompt, initial), updated: APPLIED, reason: "Uses the thirty day delivery period.", sources: SUPPORT },
    ], questions: [] }),
    [AGENT.merge]: () => `<paragraph>${APPLIED}</paragraph>`,
  })
  expect((await applyLegacyList(directory, model, DRAFT.replace(DRAFTED, initial))).status).toBe("done")
  const result = readLeanJson(directory, "d1/apply.json") as ApplyResult
  expect(result.applied.map(row => row.id)).toEqual(["a1", "a2"])
  expect(result.unapplied).toEqual([])
  expect(result.comments?.attached.map(row => row.id)).toEqual(["a2"])
  expect(result.comments?.unresolved).toEqual([expect.objectContaining({
    id: "a1", reason: "Final wording was reconciled after this proposal; its explanation needs review.",
  })])
  const parts = unzipSync(readFileSync(leanFile(directory, "d1/redline.docx")))
  const comments = new DOMParser().parseFromString(strFromU8(parts["word/comments.xml"]!), "application/xml")
  expect(comments.getElementsByTagNameNS(W, "comment").length).toBe(1)
  expect(comments.documentElement.textContent).toContain("Uses the thirty day delivery period.")
  expect(comments.documentElement.textContent).not.toContain("Allows a twenty day delivery period.")
  expect(finalGate(directory).disclosed).toContainEqual(expect.objectContaining({ kind: "omission", ref: expect.objectContaining({ amendment_id: "a1" }) }))
}, 120_000)
