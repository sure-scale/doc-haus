import { expect, test } from "bun:test"
import start from "../workflows/lean/fixtures/start.json"
import blockedGate from "../workflows/lean/fixtures/status-blocked-gate.json"
import blockedQuestion from "../workflows/lean/fixtures/status-blocked-question.json"
import completed from "../workflows/lean/fixtures/status-completed.json"
import failed from "../workflows/lean/fixtures/status-failed.json"
import orphaned from "../workflows/lean/fixtures/status-orphaned.json"
import { classifyLeanObservation, isLeanDraftingStart, isLeanDraftingStatus, leanDeliverableDocx, leanProgressLine } from "./support/lean-observation"

// The fixtures are drafting-status bodies recorded from the real ingest lean
// start/status path with a scripted model (no live calls).

function lean(value: unknown) {
  if (!isLeanDraftingStatus(value)) throw new Error("fixture is not a lean drafting status")
  return value
}

test("a lean start response is recognised and its run is observed while in flight", () => {
  expect(isLeanDraftingStart(start)).toBe(true)
  expect(classifyLeanObservation(lean(start.status))).toBe("observing")
  expect(isLeanDraftingStart({ ...start, ownerSessionID: "ses_other" })).toBe(false)
  expect(isLeanDraftingStart({ mode: "started", run: { run_id: "run", owner_session_id: "ses_owner" } })).toBe(false)
})

test("a completed lean run settles observation and yields its registered Word deliverables to stage", () => {
  const status = lean(completed)
  expect(classifyLeanObservation(status)).toBe("complete")
  expect(leanDeliverableDocx(status)).toEqual(["Supply note.docx", "Supply note - redline.docx", "Supply note - drafting notes.docx"])
  expect(leanProgressLine(status)).toBe("completed d1.draft=done/3calls d1.review=done d1.changes=done d1.apply=done finalize=done gate=done")
})

test("a review draft delivered with unresolved checklist findings is reported apart from a clean completion and still stages its deliverables", () => {
  const findings = lean({ ...completed, gate: { ...completed.gate, delivery: "review_draft_with_findings" } })
  expect(classifyLeanObservation(findings)).toBe("complete_with_findings")
  expect(leanDeliverableDocx(findings)).toEqual(["Supply note.docx", "Supply note - redline.docx", "Supply note - drafting notes.docx"])
  // A plain review draft is still a completion.
  expect(classifyLeanObservation(lean({ ...completed, gate: { ...completed.gate, delivery: "review_draft" } }))).toBe("complete")
})

test("a completed run of two documents stages every registered Word file", () => {
  const [body, redline, notes, markdown] = completed.documents.documents
  const second = { document: "d2", title: "Covering letter" }
  const status = lean({
    ...completed,
    documents: {
      ...completed.documents,
      documents: [body, redline, { ...body, ...second, deliverable: "Covering letter.docx" }, { ...redline, ...second, deliverable: "Covering letter - redline.docx" }, notes, markdown],
    },
  })
  expect(leanDeliverableDocx(status)).toEqual([
    "Supply note.docx", "Supply note - redline.docx", "Covering letter.docx", "Covering letter - redline.docx", "Supply note - drafting notes.docx",
  ])
})

test("failed and blocked runs stop observation and have nothing to stage", () => {
  expect(classifyLeanObservation(lean(failed))).toBe("failed")
  expect(leanProgressLine(lean(failed))).toStartWith("failed d1.draft=failedx4 d1.review=pending")
  for (const blocked of [blockedGate, blockedQuestion]) {
    expect(classifyLeanObservation(lean(blocked))).toBe("blocked")
    expect(() => leanDeliverableDocx(lean(blocked))).toThrow("completed lean drafting status has no delivered document list")
  }
  expect(() => leanDeliverableDocx(lean(failed))).toThrow("completed lean drafting status has no delivered document list")
})

test("a running state no coordinator owns is reported as inactive, not observed forever", () => {
  expect(orphaned).toMatchObject({ status: "running", in_flight: false })
  expect(classifyLeanObservation(lean(orphaned))).toBe("inactive")
})

test("a legacy drafting status is not read as a lean status", () => {
  expect(isLeanDraftingStatus({ active: true, runId: "run", ownerSessionID: "ses_owner", stage: { current: "draft" } })).toBe(false)
  expect(isLeanDraftingStatus({ ...completed, profile: "legacy" })).toBe(false)
})
