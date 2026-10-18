import { afterEach, expect, test } from "bun:test"
import { existsSync, readFileSync } from "node:fs"
import { DOMParser } from "@xmldom/xmldom"
import { strFromU8, unzipSync } from "fflate"
import { amendmentSchema, ensureParagraphPointers, paragraphSha256 } from "../../lib/drafting-amendment"
import { applyAmendments } from "../../lib/drafting-amendment-apply"
import { readDocxCommentParagraphs, readDocxRevisionText } from "../../lib/docx-comments"
import { renderMarkdownToDocx } from "../../lib/render"
import { docxodus } from "../../lib/docxodus"
import { checkChanges } from "./change-checks"
import { projectedChangeParagraphs } from "./change-pipeline"
import { readChangeList, readParagraphs, type ApplyResult } from "./checks"
import { finalGate } from "./gate"
import { packageReviewPrompt } from "./package-review"
import { writeLeanPlan } from "./plan"
import { appliedChangeEvidence, exportAppliedChangeComments } from "./review-evidence"
import { readRepairRound } from "./repairs"
import { runLean, stageInputSha } from "./runner"
import { paragraphRows, stageInputs } from "./stages"
import { leanFile, readLeanJson, readLeanState, readLeanText, replaceStageOpenItems, writeLeanBytes, writeLeanText, writeQuestions } from "./state"
import { AGENT, cleanupMatters, drafter, happyPath, matter, pidOf, scriptedModel, type Reply } from "./test-support"

afterEach(cleanupMatters)
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const BEFORE = "The custodian keeps the inspection register."
const DELETE = "Copies of the retired inspection form must be retained."
const COMBINED = "The custodian records each inspection and the supervisor signs the register."
const SPLIT = ["The custodian records each inspection.", "The supervisor signs the register."]
const ADJACENT = "The register is stored in the reception cabinet."
const REPLACED = "The register is stored in the records cabinet."
const AFTER = "The supervisor arranges access to the register."
const DELETE_REASON = "Removes the retired form retention instruction expressly withdrawn by the supplied instructions."
const SPLIT_REASON = "Separates the custodian and supervisor duties into distinct paragraphs while preserving both."
const REPLACE_REASON = "Uses the expressly designated storage location."
const REQUEST = "Prepare an inspection note. Remove the instruction to retain copies of the retired inspection form. Put the custodian's recording duty and the supervisor's signing duty in separate paragraphs. Store the register in the records cabinet. Preserve the remaining instructions."
const SUPPORT = {
  delete: [{ source_id: "request", quote: "Remove the instruction to retain copies of the retired inspection form." }],
  split: [{ source_id: "request", quote: "Put the custodian's recording duty and the supervisor's signing duty in separate paragraphs." }],
  replace: [{ source_id: "request", quote: "Store the register in the records cabinet." }],
}
const DRAFT = `# Inspection note\n\n## Instructions\n\n${[BEFORE, DELETE, COMBINED, ADJACENT, AFTER].join("\n\n")}\n`

function fixture() {
  const directory = matter([BEFORE, DELETE, COMBINED, ADJACENT, AFTER].join(" "))
  writeLeanText(directory, "request.md", REQUEST)
  return directory
}

const amendments = (prompt: string) => [
  { pid: pidOf(prompt, DELETE), operation: "delete", updated: "", reason: DELETE_REASON, sources: SUPPORT.delete },
  { pid: pidOf(prompt, COMBINED), operation: "split", updated: SPLIT, reason: SPLIT_REASON, sources: SUPPORT.split },
  { pid: pidOf(prompt, ADJACENT), updated: REPLACED, reason: REPLACE_REASON, sources: SUPPORT.replace },
]

function replies() {
  const changes: Reply = prompt => JSON.stringify({ amendments: amendments(prompt), questions: [] })
  return {
    ...happyPath,
    [AGENT.draft]: drafter(DRAFT),
    [AGENT.review]: () => "Apply the expressly supplied withdrawal, paragraph separation and storage instruction.",
    [AGENT["change-plan"]]: (prompt: string) => JSON.stringify({
      shared_decisions: ["Preserve all duties except the withdrawn retired form instruction."],
      groups: [{ id: "inspection-instructions", pids: amendments(prompt).map(row => row.pid),
        goal: "Delete the withdrawn instruction, separate the duties, and use the designated cabinet without changing adjacent instructions.",
        sources: Object.values(SUPPORT).flat(), limitations: [] }], dismissed: [], findings: [], questions: [],
    }),
    [AGENT["change-prepare"]]: changes,
    [AGENT.changes]: changes,
    [AGENT.finalize]: (prompt: string) => {
      const blocks = [...prompt.matchAll(/<open_items>\n([\s\S]*?)\n<\/open_items>/g)].map(match => match[1]!)
      const ids = blocks.flatMap(block => [...block.matchAll(/^- (\S+):/gm)].map(match => match[1]))
      return ["# Drafting notes", "", DELETE_REASON, SPLIT_REASON, REPLACE_REASON,
        ...ids.map(id => `- The lawyer must review the disclosed limitation before use. [ref: ${id}]`)].join("\n\n")
    },
  }
}

test("source-grounded structural changes survive preparation, projected review, native export and unchanged resume", async () => {
  const directory = fixture(), model = scriptedModel(replies())
  const repairs = { round: 2, requested_at: "2026-09-01T10:00:00.000Z", items: [], substantive_attempts: 2 }
  writeLeanText(directory, "repairs.json", JSON.stringify(repairs))
  const expected = ["Inspection note", "Instructions", BEFORE, ...SPLIT, REPLACED, AFTER]
  const port = { ...model.port, async prompt(input: Parameters<typeof model.port.prompt>[0]) {
    if (input.agent === AGENT.changes) {
      expect(input.text).toContain('"operation": "delete"')
      expect(input.text).toContain('"operation": "split"')
      expect(input.text).toContain(SPLIT[1])
    }
    if (input.agent === AGENT["change-review"]) {
      expect(existsSync(leanFile(directory, "d1/body-applied.docx"))).toBe(false)
      expect(projectedChangeParagraphs(directory, "d1").map(row => row.text)).toEqual(expected)
    }
    return model.port.prompt(input)
  } }
  const result = await runLean(directory, { sessions: port })
  expect(result.status).toBe("completed")
  expect(model.count(AGENT["change-prepare"])).toBe(1)
  expect(model.count(AGENT.changes)).toBe(1)
  expect(model.count(AGENT.merge)).toBe(0)
  expect(readRepairRound(directory)).toEqual(repairs)
  const original = readParagraphs(directory, "d1")
  const applied = readParagraphs(directory, "d1", "paragraphs-applied.json")
  expect(applied.map(row => row.text)).toEqual(expected)
  expect(new Set(applied.map(row => row.pid)).size).toBe(applied.length)
  expect(applied.find(row => row.text === SPLIT[0])?.pid).toBe(original.find(row => row.text === COMBINED)?.pid)
  for (const text of [BEFORE, AFTER]) expect(applied.find(row => row.text === text)?.pid).toBe(original.find(row => row.text === text)?.pid)
  const list = readChangeList(directory, "d1")!
  expect(list.amendments.map(row => row.operation ?? "replace")).toEqual(["delete", "split", "replace"])
  const apply = readLeanJson(directory, "d1/apply.json") as ApplyResult
  expect(apply.applied.map(row => row.id)).toEqual(["a1", "a2", "a3"])
  expect(apply.unapplied).toEqual([])
  const clean = readFileSync(leanFile(directory, "d1/body-applied.docx"))
  expect(readDocxCommentParagraphs(clean).map(row => row.paragraph_text)).toEqual(expected)
  expect(unzipSync(clean)["word/comments.xml"]).toBeUndefined()
  const redline = readFileSync(leanFile(directory, "d1/redline.docx"))
  expect(readDocxCommentParagraphs(redline).map(row => row.paragraph_text)).toEqual(expected)
  expect(readDocxRevisionText(redline).deleted_text.join(" ")).toContain(DELETE)
  const splitComments = apply.comments?.attached.filter(row => row.text.includes(SPLIT_REASON)) ?? []
  expect(splitComments).toHaveLength(1)
  expect(splitComments[0]?.paragraphs?.map(row => row.paragraph_text)).toEqual(SPLIT)
  expect(apply.applied[1]?.result_pids).toEqual(applied.filter(row => SPLIT.includes(row.text)).map(row => row.pid))
  expect(apply.applied[0]?.result_pids).toEqual([])
  const nativeParts = unzipSync(redline)
  const nativeComments = new DOMParser().parseFromString(strFromU8(nativeParts["word/comments.xml"]!), "application/xml")
  expect(nativeComments.getElementsByTagNameNS(W, "comment").length).toBeGreaterThan(0)
  expect(nativeComments.documentElement.textContent).toContain(SPLIT_REASON)
  expect(nativeComments.documentElement.textContent).toContain(SUPPORT.split[0].quote)
  const deletion = apply.comments?.attached.find(row => row.id === "a1") ?? apply.comments?.unresolved.find(row => row.id === "a1")
  expect(deletion?.text).toContain(DELETE_REASON)
  const evidence = appliedChangeEvidence(directory, "d1")
  expect(evidence).toHaveLength(3)
  expect(evidence.map(row => row.sources)).toEqual(Object.values(SUPPORT))
  expect(evidence.map(row => row.operation)).toEqual(["delete", "split", "replace"])
  expect(evidence[0]?.deleted_text).toBe(DELETE)
  expect(evidence[0]?.result_paragraphs).toEqual([])
  expect(evidence[1]?.result_paragraphs).toEqual(applied.filter(row => SPLIT.includes(row.text)).map(row => ({ pid: row.pid, text: row.text })))
  for (const prompt of [model.prompts.find(row => row.agent === AGENT.finalize)!.text, packageReviewPrompt(directory)]) {
    expect(prompt).toContain(DELETE_REASON)
    expect(prompt).toContain(SPLIT_REASON)
    expect(prompt).toContain(SUPPORT.delete[0].quote)
    expect(prompt).toContain(SUPPORT.split[0].quote)
  }
  expect(readLeanText(directory, "drafting-notes.md")).toContain(DELETE_REASON)
  expect(finalGate(directory).status).toBe("pass")
  const stableKeys = ["intake", "research", "research-review", "discrepancy", "spine", "d1.draft", "d1.review", "d1.change-plan", "d1.change-prepare", "d1.changes", "d1.change-review", "d1.apply", "finalize", "package-review"] as const
  const keys = stableKeys.map(key => stageInputSha(directory, key))
  for (const key of stableKeys) {
    const structural = stageInputs(directory, key).find(input => input.name === "structural-amendment-contract")
    expect(Boolean(structural)).toBe(["d1.change-plan", "d1.change-prepare", "d1.changes", "d1.change-review", "d1.apply", "finalize", "package-review"].includes(key))
  }
  const state = readLeanState(directory)!
  const pointers = await ensureParagraphPointers(leanFile(directory, "d1/body-applied.docx"))
  expect(readFileSync(leanFile(directory, "d1/body-applied.docx"))).toEqual(clean)
  const resumed = scriptedModel(replies())
  expect((await runLean(directory, { sessions: resumed.port })).status).toBe("completed")
  expect(resumed.prompts).toEqual([])
  expect(await ensureParagraphPointers(leanFile(directory, "d1/body-applied.docx"))).toEqual(pointers)
  expect(readFileSync(leanFile(directory, "d1/redline.docx"))).toEqual(redline)
  expect(readLeanState(directory)?.stages).toEqual(state.stages)
  expect(readRepairRound(directory)).toEqual(repairs)
  stableKeys.forEach((key, index) => expect(stageInputSha(directory, key)).toBe(keys[index]))
  const savedList = readLeanText(directory, "d1/changes.json")!
  const changedList = structuredClone(list)
  changedList.amendments[1] = { ...changedList.amendments[1]!, operation: "replace", updated: SPLIT.join(" ") }
  writeLeanText(directory, "d1/changes.json", JSON.stringify(changedList))
  for (const key of ["intake", "research", "research-review", "discrepancy", "spine", "d1.draft", "d1.review", "d1.change-plan", "d1.change-prepare", "d1.changes"] as const)
    expect(stageInputSha(directory, key)).toBe(keys[stableKeys.indexOf(key)])
  for (const key of ["d1.change-review", "d1.apply", "finalize", "package-review"] as const)
    expect(stageInputSha(directory, key)).not.toBe(keys[stableKeys.indexOf(key)])
  expect(readRepairRound(directory)).toEqual(repairs)
  expect(readLeanState(directory)?.stages).toEqual(state.stages)
  writeLeanText(directory, "d1/changes.json", savedList)
  expect((await runLean(directory, { sessions: resumed.port })).status).toBe("completed")
  expect(resumed.prompts).toEqual([])
}, 120_000)

test("structural admission rejects only missing support, misquotes and disappearing uncertainty", () => {
  const directory = fixture()
  const paragraphs = [DELETE, COMBINED, ADJACENT].map((text, index) => ({ pid: `p00000000000${index + 1}`, text, sha256: paragraphSha256(text) }))
  const prompt = paragraphs.map(row => `[${row.pid}] ${row.text}`).join("\n")
  const good = { amendments: amendments(prompt), questions: [] }
  expect(checkChanges(directory, JSON.stringify(good), paragraphs).issues).toEqual([])
  for (const index of [0, 1]) {
    const missing = structuredClone(good)
    missing.amendments[index]!.sources = []
    expect(checkChanges(directory, JSON.stringify(missing), paragraphs).issues.length).toBeGreaterThan(0)
    const withheld = checkChanges(directory, JSON.stringify(missing), paragraphs, { withholdUnverified: true })
    expect(withheld.issues).toEqual([])
    expect(withheld.changes?.amendments).toHaveLength(2)
    expect(withheld.changes?.withheld?.[0]?.amendment.pid).toBe(paragraphs[index]!.pid)
    expect(withheld.openItems).toContainEqual(expect.objectContaining({ kind: "unapplied_amendment" }))
    const misquote = structuredClone(good)
    misquote.amendments[index]!.sources[0]!.quote = "Remove all inspection register duties."
    expect(checkChanges(directory, JSON.stringify(misquote), paragraphs).issues.join(" ")).toContain("not in")
  }
  expect(checkChanges(directory, JSON.stringify(good), paragraphs).issues).toEqual([])
  const question = { ...good, questions: [{ pid: paragraphs[0]!.pid, question: "Is the retention instruction withdrawn?" }] }
  expect(checkChanges(directory, JSON.stringify(question), paragraphs).issues.join(" ")).toContain("open item")
})

test("paragraph deletion and splitting require explicit operations while replacement safeguards remain strict", () => {
  const base = { id: "fictional", pid: "p000000000001", existing_sha256: paragraphSha256(DELETE), reason: DELETE_REASON, sources: SUPPORT.delete }
  expect(amendmentSchema.safeParse({ ...base, operation: "delete", updated: "" }).success).toBe(true)
  expect(amendmentSchema.safeParse({ ...base, operation: "split", updated: SPLIT }).success).toBe(true)
  for (const value of [
    { ...base, updated: "" },
    { ...base, updated: SPLIT },
    { ...base, operation: "delete", updated: DELETE },
    { ...base, operation: "split", updated: [SPLIT[0]] },
    { ...base, operation: "split", updated: [SPLIT[0], ""] },
    { ...base, operation: "split", updated: [SPLIT[0], "First line.\nSecond line."] },
    { ...base, operation: "replace", updated: "First line.\nSecond line." },
  ]) expect(amendmentSchema.safeParse(value).success).toBe(false)
})

test("structural table edits retain the cell boundary, unaffected cells and unique stable pointers", async () => {
  const directory = fixture()
  const bytes = (await renderMarkdownToDocx(`# Inspection table\n\n| Instructions | Custody |\n| --- | --- |\n| ${DELETE} | ${BEFORE} |\n| ${COMBINED} | ${AFTER} |\n`)).bytes
  writeLeanBytes(directory, "table-before.docx", bytes)
  const docxPath = leanFile(directory, "table-before.docx"), outPath = leanFile(directory, "table-after.docx")
  const pointers = await ensureParagraphPointers(docxPath)
  const pid = (text: string) => Object.entries(pointers).find(([, value]) => value === text)![0]
  const baseline = new DOMParser().parseFromString(strFromU8(unzipSync(readFileSync(docxPath))["word/document.xml"]!), "application/xml")
  const baselineCells = Array.from(baseline.getElementsByTagNameNS(W, "tc"))
  expect(baselineCells).toHaveLength(6)
  expect(baselineCells.every(cell => Array.from(cell.childNodes).some(node => node.nodeType === 1 && (node as Element).localName === "p"))).toBe(true)
  const targets = [
    { id: "delete-cell", pid: pid(DELETE), existing_sha256: paragraphSha256(DELETE), operation: "delete", updated: "", reason: DELETE_REASON, sources: SUPPORT.delete },
    { id: "split-cell", pid: pid(COMBINED), existing_sha256: paragraphSha256(COMBINED), operation: "split", updated: SPLIT, reason: SPLIT_REASON, sources: SUPPORT.split },
  ] satisfies Parameters<typeof applyAmendments>[0]["amendments"]
  const result = await applyAmendments({ docxPath, outPath, merge: async () => "unused", amendments: targets })
  expect(result.unapplied).toEqual([])
  expect(result.applied.map(row => row.id)).toEqual(["delete-cell", "split-cell"])
  const written = readFileSync(outPath)
  const actual = await ensureParagraphPointers(outPath)
  expect(readFileSync(outPath)).toEqual(written)
  expect(actual[pid(DELETE)]).toBeUndefined()
  expect(actual[pid(BEFORE)]).toBe(BEFORE)
  expect(actual[pid(AFTER)]).toBe(AFTER)
  expect(actual[pid(COMBINED)]).toBe(SPLIT[0])
  expect(Object.values(actual).filter(text => SPLIT.includes(text))).toEqual(SPLIT)
  expect(result.applied[0]?.result_pids).toEqual([])
  expect(result.applied[1]?.result_pids).toEqual(Object.entries(actual).filter(([, text]) => SPLIT.includes(text)).map(([pid]) => pid))
  const body = new DOMParser().parseFromString(strFromU8(unzipSync(written)["word/document.xml"]!), "application/xml")
  const cells = Array.from(body.getElementsByTagNameNS(W, "tc"))
  expect(cells).toHaveLength(6)
  expect(cells.every(cell => Array.from(cell.childNodes).some(node => node.nodeType === 1 && (node as Element).localName === "p"))).toBe(true)
  const splitCell = cells.find(cell => cell.textContent?.includes(SPLIT[0]))!
  expect(Array.from(splitCell.getElementsByTagNameNS(W, "p")).map(paragraph => paragraph.textContent)).toEqual(SPLIT)
  const redline = await (await docxodus()).compareDocuments(new Uint8Array(readFileSync(docxPath)), new Uint8Array(written))
  const comments = exportAppliedChangeComments(redline, written, "d1", targets, result.applied)
  expect(comments.comments.attached.find(row => row.id === "split-cell")?.paragraphs?.map(row => row.paragraph_text)).toEqual(SPLIT)
  expect(comments.comments.unresolved.filter(row => row.id === "split-cell")).toEqual([])
}, 120_000)

test("stale structural targets preserve their text and missing deletion targets remain unapplied", async () => {
  const directory = fixture()
  writeLeanBytes(directory, "stale-before.docx", (await renderMarkdownToDocx(DRAFT)).bytes)
  const docxPath = leanFile(directory, "stale-before.docx"), outPath = leanFile(directory, "stale-after.docx")
  const pointers = await ensureParagraphPointers(docxPath)
  const pid = (text: string) => Object.entries(pointers).find(([, value]) => value === text)![0]
  const targets = [
    { id: "delete", pid: pid(DELETE), existing_sha256: paragraphSha256(DELETE), operation: "delete" as const, updated: "", reason: DELETE_REASON, sources: SUPPORT.delete },
    { id: "split", pid: pid(COMBINED), existing_sha256: paragraphSha256(COMBINED), operation: "split" as const, updated: SPLIT, reason: SPLIT_REASON, sources: SUPPORT.split },
  ]
  const merge = async () => { throw new Error("Unchanged structural paragraphs must never require a merge.") }
  expect((await applyAmendments({ docxPath, outPath, amendments: targets, merge })).unapplied).toEqual([])
  const mutated = targets.map(row => ({ ...row, existing_sha256: paragraphSha256("A different earlier inspection instruction.") }))
  const stale = await applyAmendments({ docxPath, outPath, amendments: mutated, merge })
  expect(stale.applied).toEqual([])
  expect(stale.unapplied.map(row => row.reason)).toEqual(["stale", "stale"])
  expect(readDocxCommentParagraphs(readFileSync(outPath)).map(row => row.paragraph_text)).toContain(DELETE)
  expect(readDocxCommentParagraphs(readFileSync(outPath)).map(row => row.paragraph_text)).toContain(COMBINED)
  await applyAmendments({ docxPath, outPath, amendments: [targets[0]!], merge })
  const missing = await applyAmendments({ docxPath: outPath, outPath: leanFile(directory, "reapplied.docx"), amendments: [targets[0]!], merge })
  expect(missing.applied).toEqual([])
  expect(missing.unapplied.map(row => row.reason)).toEqual(["target_deleted"])
}, 120_000)

test("a question remains tied to uncertainty in the second split paragraph and cannot disappear through deletion", async () => {
  const directory = fixture()
  writeLeanPlan(directory, [{ title: "Inspection note", purpose: "Record supplied instructions", parts: [{ heading: "Instructions", scope: "Inspection instructions" }] }])
  writeLeanBytes(directory, "d1/body.docx", (await renderMarkdownToDocx(DRAFT)).bytes)
  const original = await paragraphRows(leanFile(directory, "d1/body.docx"))
  writeLeanText(directory, "d1/paragraphs.json", JSON.stringify(original))
  const target = original.find(row => row.text === COMBINED)!
  const question = { pid: target.pid, question: "Who may sign the inspection register?" }
  const marked = `${SPLIT[1]} [open item: confirm who may sign the inspection register]`
  const reply = { amendments: [{ pid: target.pid, operation: "split", updated: [SPLIT[0], marked], reason: SPLIT_REASON, sources: SUPPORT.split }], questions: [question] }
  const accepted = checkChanges(directory, JSON.stringify(reply), original, { document: "d1" })
  expect(accepted.issues).toEqual([])
  expect(accepted.changes?.questions).toHaveLength(1)
  writeQuestions(directory, accepted.changes!.questions)
  replaceStageOpenItems(directory, "changes", "d1", accepted.openItems)
  writeLeanText(directory, "d1/changes.json", JSON.stringify(accepted.changes))
  const applied = await applyAmendments({ docxPath: leanFile(directory, "d1/body.docx"), outPath: leanFile(directory, "d1/body-applied.docx"), amendments: accepted.changes!.amendments, merge: async () => "unused" })
  expect(applied.unapplied).toEqual([])
  const final = await paragraphRows(leanFile(directory, "d1/body-applied.docx"))
  writeLeanText(directory, "d1/paragraphs-applied.json", JSON.stringify(final))
  writeLeanText(directory, "d1/apply.json", JSON.stringify(applied))
  writeLeanText(directory, "drafting-notes.md", `# Drafting notes\n\n${accepted.openItems.map(item => `${item.id}: ${item.detail}`).join("\n\n")}`)
  expect(finalGate(directory).status).toBe("pass")
  expect(finalGate(directory).delivery).toBe("review_draft")
  expect(finalGate(directory).disclosed).toContainEqual(expect.objectContaining({ kind: "question" }))
  const silent = { ...reply, amendments: [{ ...reply.amendments[0]!, updated: SPLIT }] }
  expect(checkChanges(directory, JSON.stringify(silent), original, { document: "d1" }).issues.join(" ")).toContain("open item")
  const deletion = { ...reply, amendments: [{ ...reply.amendments[0]!, operation: "delete", updated: "" }] }
  expect(checkChanges(directory, JSON.stringify(deletion), original, { document: "d1" }).issues.join(" ")).toContain("open item")
  writeLeanText(directory, "d1/paragraphs-applied.json", JSON.stringify(final.map(row => row.text === marked ? { ...row, text: SPLIT[1] } : row)))
  expect(finalGate(directory).status).toBe("blocked")
  writeLeanText(directory, "d1/paragraphs-applied.json", JSON.stringify(final))
  expect(finalGate(directory).status).toBe("pass")
  writeLeanText(directory, "d1/apply.json", JSON.stringify({ ...applied, applied: applied.applied.map(row => ({ ...row, result_pids: [] })) }))
  expect(finalGate(directory).status).toBe("blocked")
}, 120_000)
