import { afterEach, describe, expect, test } from "bun:test"
import { readFileSync } from "node:fs"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { docxodus } from "../../lib/docxodus"
import type { Amendment } from "../../lib/drafting-amendment"
import { renderMarkdownToDocx } from "../../lib/render"
import { runApply } from "./apply-stage"
import { checkChanges } from "./change-checks"
import { finalBodyParagraphs, readChangeList, readParagraphs, type ApplyResult, type ChangeExplanation, type ChangeList, type ParagraphRow } from "./checks"
import { recordLeanDocuments } from "./deliverables"
import { finalGate } from "./gate"
import { packageReviewPrompt } from "./package-review"
import { parsePlan, writeLeanPlan } from "./plan"
import { stagePrompt } from "./prompts"
import { appliedChangeEvidence, appliedChangeEvidenceFromRecords, exportAppliedChangeComments } from "./review-evidence"
import { runLean, stageInputSha } from "./runner"
import { acceptStageReply, paragraphRows } from "./stages"
import { leanFile, readLeanJson, readLeanState, replaceStageOpenItems, updateLeanStage, writeLeanBytes, writeLeanText } from "./state"
import { AGENT, APPLIED, DRAFT, DRAFTED, SOURCE, SOURCE_TEXT, SUPPORT, cleanupMatters, happyPath, matter, pidOf, registering, scriptedModel } from "./test-support"

afterEach(cleanupMatters)

test("delivery context projects the files registration uses without treating the plan as user intent", async () => {
  const directory = matter()
  const request = "Revise input-terms.docx; supply an annotated comparison named delivery-review.docx and a separate explanatory note."
  writeLeanText(directory, "request.md", request)
  const document = { title: "Delivery terms", purpose: "Revise delivery", parts: [{ heading: "Delivery", scope: "Approved period" }], mode: "revise" as const, deliverable: "redline" as const, file: "delivery-review.docx" }
  writeLeanPlan(directory, [document])
  const bytes = (await renderMarkdownToDocx("# Delivery terms\n\n## Delivery\n\nDelivery is due within twelve days.")).bytes
  writeLeanBytes(directory, "d1/body-applied.docx", bytes)
  writeLeanBytes(directory, "d1/redline.docx", bytes)
  writeLeanBytes(directory, "drafting-notes.docx", bytes)
  const projected = JSON.parse(/<delivery_artifacts>\n([\s\S]*?)\n<\/delivery_artifacts>/.exec(packageReviewPrompt(directory))?.[1] ?? "[]") as { kind: string; planned_filename: string; file_status: string }[]
  expect(projected.map(row => [row.kind, row.planned_filename, row.file_status])).toEqual([
    ["body", "delivery-review - clean.docx", "present"],
    ["redline", "delivery-review.docx", "present"],
    ["drafting_notes", "delivery-review - drafting notes.docx", "present"],
  ])
  const registered = await recordLeanDocuments(directory, { ...registering(), sessions: scriptedModel(happyPath).port }, true)
  expect(registered.documents.flatMap(row => row.deliverable ?? [])).toEqual(projected.map(row => row.planned_filename))
  // Change only the model's selection: review must see the resulting mismatch,
  // without relabelling that selection as the lawyer's requested output.
  const admitted = parsePlan("Acting for: not stated\n# Delivery terms\nForm: agreement\nMode: revise\nDeliverable: body\nFile: input-terms.docx\n## Delivery\nApproved period", request)
  expect(admitted.issues).toEqual([])
  expect(admitted.value?.documents[0]?.file).toBe("input-terms.docx")
  writeLeanPlan(directory, admitted.value!.documents)
  for (const prompt of [stagePrompt(directory, "finalize"), packageReviewPrompt(directory)]) {
    expect(prompt).toContain("<delivery_artifacts>")
    expect(prompt).toContain('"planned_filename": "input-terms - redline.docx"')
    expect(prompt).toContain('"planned_file": "input-terms.docx"')
    expect(prompt).not.toContain('"requested_file": "input-terms.docx"')
    expect(prompt).toContain("delivery-review.docx")
    expect(prompt).toContain('"kind": "drafting_notes"')
  }
  expect(parsePlan("Acting for: not stated\n# Delivery terms\nForm: agreement\nFile: input-terms.docx\n## Delivery", "Revise input-terms.docx and deliver the revised body as input-terms.docx.").issues).toEqual([])
}, 120_000)

test("memo and final reviewer retain source arrays independently of application mode", () => {
  const directory = matter()
  writeLeanPlan(directory, [{ title: "Supply note", purpose: "Record terms", parts: [{ heading: "Delivery", scope: "Agreed delivery" }] }])
  const amendment = { id: "a-source", pid: "p111111111111", existing_sha256: "initial", updated: "Supported revised wording.", reason: "Carry the supplied delivery period.", sources: SUPPORT }
  const changes = (sources: typeof SUPPORT) => writeLeanText(directory, "d1/changes.json", JSON.stringify({ amendments: [{ ...amendment, sources }], questions: [] }))
  changes(SUPPORT)
  const apply = (mode: "code" | "already_applied") => writeLeanText(directory, "d1/apply.json", JSON.stringify({ applied: [{ id: amendment.id, pid: amendment.pid, mode }], unapplied: [] }))
  apply("code")
  const supported = appliedChangeEvidence(directory, "d1")[0]!
  expect(supported.evidence_basis).toBe("source-passages-provided")
  expect(supported.sources).toEqual(SUPPORT)
  for (const prompt of [stagePrompt(directory, "finalize"), packageReviewPrompt(directory)]) {
    expect(prompt).toContain(SOURCE)
    expect(prompt).toContain(SUPPORT[0]!.quote)
    expect(prompt).toContain('"application_mode": "code"')
    expect(prompt).toContain('"evidence_basis": "source-passages-provided"')
  }
  apply("already_applied")
  expect(appliedChangeEvidence(directory, "d1")[0]!.evidence_basis).toBe(supported.evidence_basis)
  expect(appliedChangeEvidence(directory, "d1")[0]!.sources).toEqual(SUPPORT)
  const key = stageInputSha(directory, "package-review")
  changes([])
  expect(appliedChangeEvidence(directory, "d1")[0]!.evidence_basis).toBe("no-source-passage")
  expect(appliedChangeEvidence(directory, "d1")[0]!.sources).toEqual([])
  expect(stageInputSha(directory, "package-review")).not.toBe(key)
  expect(packageReviewPrompt(directory)).toContain('"evidence_basis": "no-source-passage"')
})


test("review reads the native comparison separately from the clean body and keys its actual bytes", async () => {
  const directory = matter()
  writeLeanPlan(directory, [{ title: "Equipment terms", purpose: "Revise rent", mode: "revise", deliverable: "redline", parts: [{ heading: "Rent", scope: "Approved amount" }] }])
  const before = (await renderMarkdownToDocx("# Equipment terms\n\n## Rent\n\nThe monthly rent is 2400 units." )).bytes
  const after = (await renderMarkdownToDocx("# Equipment terms\n\n## Rent\n\nThe monthly rent is 2700 units." )).bytes
  const redline = await (await docxodus()).compareDocuments(before, after, { authorName: "doc.haus" })
  writeLeanText(directory, "d1/paragraphs-applied.json", JSON.stringify([{ pid: "p111111111111", sha256: "final", text: "The monthly rent is 2700 units." }]))
  writeLeanBytes(directory, "d1/redline.docx", redline)
  writeLeanText(directory, "d1/apply.json", JSON.stringify({ applied: [], unapplied: [], comments: { author: "Doc.Haus", attached: [], unresolved: [] } }))
  const prompt = packageReviewPrompt(directory)
  expect(prompt).toContain("<native_export")
  expect(prompt).toContain('"delivery_registration": "pending"')
  expect(prompt).toContain('"deleted_text"')
  expect(prompt).toContain("2400")
  expect(prompt).toContain("2700")
  expect(prompt).toContain('"unapplied_amendment_count": 0')
  expect(stagePrompt(directory, "finalize")).toContain("<native_export")
  const packageKey = stageInputSha(directory, "package-review")
  const memoKey = stageInputSha(directory, "finalize")
  writeLeanBytes(directory, "d1/redline.docx", after)
  expect(stageInputSha(directory, "package-review")).not.toBe(packageKey)
  expect(stageInputSha(directory, "finalize")).not.toBe(memoKey)
  expect(packageReviewPrompt(directory)).toContain('"deleted_text": []')
}, 120_000)


test("native outcome counts keep withheld proposals separate from amendment application and initial explanations", () => {
  const directory = matter()
  writeLeanPlan(directory, [{ title: "Supply note", purpose: "Record terms", parts: [{ heading: "Delivery", scope: "Delivery" }] }])
  writeLeanText(directory, "d1/paragraphs.json", JSON.stringify([{ pid: "p111111111111", sha256: "initial", text: "Delivery term" }]))
  writeLeanText(directory, "d1/changes.json", JSON.stringify({ amendments: [], explanations: [{ id: "i1", pid: "p111111111111", reason: "Initial explanation", sources: SUPPORT }], questions: [], withheld: [{ amendment: { id: "a1", pid: "p111111111111", existing_sha256: "initial", updated: "Unverified proposal", reason: "Not supported", sources: [] }, issues: ["Missing confirmation marker"] }] }))
  writeLeanText(directory, "d1/apply.json", JSON.stringify({ applied: [], unapplied: [] }))
  const prompt = packageReviewPrompt(directory)
  expect(prompt).toContain('"applied_amendment_count": 0')
  expect(prompt).toContain('"unapplied_amendment_count": 0')
  expect(prompt).toContain('"withheld_proposal_count": 1')
  expect(prompt).toContain('"change_origin": "initial-draft-revision"')
})

describe("comments integration", () => {
  const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

  // A change list applied without review can hold overlapping paragraph
  // proposals. Exercise that apply/notes boundary directly: runLean batches
  // consolidate replacements before change-review, so they never reach it.
  async function applyLegacyList(directory: string, model: ReturnType<typeof scriptedModel>, draft = DRAFT) {
    const plan = parsePlan(`Acting for: not stated\n${draft.replace(/^# .+$/m, "$&\nForm: other")}`)
    expect(plan.issues).toEqual([])
    writeLeanPlan(directory, plan.value!.documents)
    const rendered = await acceptStageReply(directory, "d1.draft", draft)
    expect(rendered.issues).toEqual([])
    // Written directly: the changes step refuses two amendments to one paragraph.
    const changes = checkChanges(directory, await model.port.prompt({
      directory, sessionID: "ses_legacy_changes", agent: AGENT.changes, text: stagePrompt(directory, "d1.changes"),
    }), readParagraphs(directory, "d1"), { document: "d1" })
    expect(changes.issues).toEqual([])
    writeLeanText(directory, "d1/changes.json", `${JSON.stringify(changes.changes, null, 2)}\n`)
    replaceStageOpenItems(directory, "changes", "d1", changes.openItems)
    const applied = await runApply(directory, readLeanState(directory)!, "d1")
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
    }], [{ id: "a1", pid: located.pid, mode: "code" }])
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
})

describe("review evidence records", () => {
  const passage = [{ source_id: "instructions.txt", quote: "Keep the register digitally." }]
  const paragraph = (pid: string, text: string): ParagraphRow => ({ pid, text, sha256: `hash-${pid}` })
  const changes = (amendments: Amendment[] = [], explanations: ChangeExplanation[] = []): ChangeList => ({ amendments, explanations, questions: [] })

  test("record evidence uses reconciled current wording and document order for replacements and splits", () => {
    const list = changes([
      { id: "replace", pid: "p1", existing_sha256: "before", updated: "Proposed wording.", reason: "Use digital storage.", sources: passage },
      { id: "split", pid: "p2", existing_sha256: "before", operation: "split", updated: ["Proposed first.", "Proposed second."], reason: "Separate the duties.", sources: passage },
    ])
    const apply: ApplyResult = { applied: [
      { id: "replace", pid: "p1", mode: "code" },
      { id: "split", pid: "p2", mode: "code", result_pids: ["p2b", "p2"] },
    ], unapplied: [] }
    const current = [paragraph("other", "Unchanged."), paragraph("p1", "Reconciled replacement."), paragraph("p2", "Reconciled first."), paragraph("p2b", "Reconciled second.")]
    const before = structuredClone({ list, apply, current })
    const evidence = appliedChangeEvidenceFromRecords(list, apply, current, [])
    expect(evidence[0]).toMatchObject({ operation: "replace", application_mode: "code", result_paragraphs: [{ pid: "p1", text: "Reconciled replacement." }] })
    expect(evidence[1]).toMatchObject({ operation: "split", application_mode: "code", result_paragraphs: [
      { pid: "p2", text: "Reconciled first." }, { pid: "p2b", text: "Reconciled second." },
    ] })
    expect(evidence[0].deleted_text).toBeUndefined()
    expect(Object.hasOwn(evidence[0], "deleted_text")).toBe(true)
    expect({ list, apply, current }).toEqual(before)
  })

  test("deletion evidence retains initial text separately from later reconciled results", () => {
    const list = changes([
      { id: "delete", pid: "p1", existing_sha256: "before", operation: "delete", updated: "", reason: "Remove the instruction.", sources: passage },
      { id: "retained", pid: "p2", existing_sha256: "before", operation: "delete", updated: "", reason: "Proposed removal.", sources: [] },
      { id: "missing-original", pid: "p3", existing_sha256: "before", operation: "delete", updated: "", reason: "Remove another instruction.", sources: [] },
    ])
    const apply: ApplyResult = { applied: [
      { id: "delete", pid: "p1", mode: "code", result_pids: [] },
      { id: "retained", pid: "p2", mode: "code", result_pids: ["p2"] },
      { id: "missing-original", pid: "p3", mode: "code", result_pids: [] },
    ], unapplied: [] }
    const evidence = appliedChangeEvidenceFromRecords(list, apply, [paragraph("p2", "Later reconciliation retained this wording.")], [paragraph("p1", "Original deleted instruction."), paragraph("p2", "Original retained instruction.")])
    expect(evidence[0]).toMatchObject({ operation: "delete", deleted_text: "Original deleted instruction.", result_paragraphs: [] })
    expect(evidence[1]).toMatchObject({ operation: "delete", deleted_text: "Original retained instruction.", result_paragraphs: [{ pid: "p2", text: "Later reconciliation retained this wording." }] })
    expect(evidence[2]).toMatchObject({ deleted_text: "", result_paragraphs: [] })
  })

  test("an explicit empty result list never falls back to the original paragraph id", () => {
    const list = changes([{ id: "a1", pid: "p1", existing_sha256: "before", updated: "Proposed wording.", reason: "Revise wording.", sources: [] }])
    const current = [paragraph("p1", "Current wording.")]
    expect(appliedChangeEvidenceFromRecords(list, { applied: [{ id: "a1", pid: "p1", mode: "code" }], unapplied: [] }, current, [])[0].result_paragraphs).toEqual([{ pid: "p1", text: "Current wording." }])
    expect(appliedChangeEvidenceFromRecords(list, { applied: [{ id: "a1", pid: "p1", mode: "code", result_pids: [] }], unapplied: [] }, current, [])[0].result_paragraphs).toEqual([])
  })

  test("source support stays distinct from application mode and native comment status", () => {
    const list = changes([
      { id: "code", pid: "p1", existing_sha256: "before", updated: "Code result.", reason: "Supported change.", sources: passage },
      { id: "differs", pid: "p2", existing_sha256: "before", updated: "Differing proposal.", reason: "Another supported change.", sources: passage },
      { id: "already", pid: "p3", existing_sha256: "before", updated: "Existing result.", reason: "Unsupported change.", sources: [] },
    ])
    const attached = { id: "already", comment_id: "c1", paragraph_index: 2, paragraph_text: "Existing result.", text: "An attached explanation does not supply a source." }
    const unresolved = { id: "differs", reason: "Final wording differs from the proposal.", text: "Proposed explanation." }
    const apply: ApplyResult = { applied: [
      { id: "code", pid: "p1", mode: "code" },
      { id: "differs", pid: "p2", mode: "code" },
      { id: "already", pid: "p3", mode: "already_applied" },
      { id: "absent", pid: "p4", mode: "code" },
    ], unapplied: [{ id: "not-applied", pid: "p5", reason: "stale" }], comments: {
      author: "Doc.Haus", attached: [attached], unresolved: [unresolved, { id: "already", reason: "Earlier unresolved record.", text: "Earlier explanation." }],
    } }
    const evidence = appliedChangeEvidenceFromRecords(list, apply, [], [])
    expect(evidence.map(row => row.application_mode)).toEqual(["code", "code", "already_applied", "code"])
    expect(evidence.map(row => row.evidence_basis)).toEqual(["source-passages-provided", "source-passages-provided", "no-source-passage", "amendment-unavailable"])
    expect(evidence[0].native_comment).toEqual({ status: "unavailable" })
    expect(evidence[1].native_comment).toEqual({ status: "unresolved", ...unresolved })
    expect(evidence[2].native_comment).toEqual({ status: "attached", author: "Doc.Haus", ...attached })
    expect(evidence[3]).toEqual({ id: "absent", pid: "p4", change_origin: "review-amendment", application_mode: "code", operation: "unavailable", result_paragraphs: [], deleted_text: undefined, evidence_basis: "amendment-unavailable", reason: "", sources: [], native_comment: { status: "unavailable" } })
    expect(evidence.map(row => row.sources)).toEqual([passage, passage, [], []])
  })

  test("initial revision explanations survive missing apply records and retain their own comment evidence", () => {
    const list = changes([], [
      { id: "initial-supported", pid: "p1", reason: "Explains the initial revision.", sources: passage },
      { id: "initial-unavailable", pid: "p2", reason: "Explains wording no longer present.", sources: [] },
    ])
    const current = [paragraph("p1", "Current reconciled initial revision.")]
    const withoutApply = appliedChangeEvidenceFromRecords(list, undefined, current, [])
    expect(withoutApply[0]).toEqual({ id: "initial-supported", pid: "p1", change_origin: "initial-draft-revision", application_mode: "initial-draft", operation: "explanation", result_paragraphs: [{ pid: "p1", text: "Current reconciled initial revision." }], deleted_text: undefined, evidence_basis: "source-passages-provided", reason: "Explains the initial revision.", sources: passage, native_comment: { status: "unavailable" } })
    expect(withoutApply[1]).toMatchObject({ evidence_basis: "no-source-passage", result_paragraphs: [], native_comment: { status: "unavailable" } })
    const unresolved = { id: "initial-supported", reason: "Anchor not located.", text: "Explanation remains for review." }
    const apply: ApplyResult = { applied: [], unapplied: [], comments: { author: "Doc.Haus", attached: [], unresolved: [unresolved] } }
    expect(appliedChangeEvidenceFromRecords(list, apply, current, [])[0].native_comment).toEqual({ status: "unresolved", ...unresolved })
    expect(appliedChangeEvidenceFromRecords(undefined, undefined, current, [])).toEqual([])
    expect(appliedChangeEvidenceFromRecords(undefined, apply, current, [])).toEqual([])
    expect(appliedChangeEvidence(matter(), "d1")).toEqual([])
  })

  test("captured native records and the directory wrapper agree after the real scripted coordinator applies a revision", async () => {
    const directory = matter()
    expect((await runLean(directory, { sessions: scriptedModel(happyPath).port })).status).toBe("completed")
    const list = readChangeList(directory, "d1")
    const apply = readLeanJson(directory, "d1/apply.json") as ApplyResult
    const current = finalBodyParagraphs(directory, "d1")
    const initial = readParagraphs(directory, "d1")
    expect(apply.applied).toHaveLength(1)
    const evidence = appliedChangeEvidenceFromRecords(list, apply, current, initial)
    expect(evidence).toEqual(appliedChangeEvidence(directory, "d1"))
    expect(evidence[0]).toMatchObject({ change_origin: "review-amendment", evidence_basis: "source-passages-provided", sources: SUPPORT, result_paragraphs: [{ pid: apply.applied[0].pid, text: APPLIED }], native_comment: { status: "attached" } })
  }, 120_000)
})
