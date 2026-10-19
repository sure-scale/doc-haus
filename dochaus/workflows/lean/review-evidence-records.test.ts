import { afterEach, expect, test } from "bun:test"
import type { Amendment } from "../../lib/drafting-amendment"
import { finalBodyParagraphs, readChangeList, readParagraphs, type ApplyResult, type ChangeExplanation, type ChangeList, type ParagraphRow } from "./checks"
import { appliedChangeEvidence, appliedChangeEvidenceFromRecords } from "./review-evidence"
import { runLean } from "./runner"
import { readLeanJson } from "./state"
import { APPLIED, SUPPORT, cleanupMatters, happyPath, matter, scriptedModel } from "./test-support"

afterEach(cleanupMatters)

const passage = [{ source_id: "instructions.txt", quote: "Keep the register digitally." }]
const paragraph = (pid: string, text: string): ParagraphRow => ({ pid, text, sha256: `hash-${pid}` })
const changes = (amendments: Amendment[] = [], explanations: ChangeExplanation[] = []): ChangeList => ({ amendments, explanations, questions: [], shared_paragraphs: [] })

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
