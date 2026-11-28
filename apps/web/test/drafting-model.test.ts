import { expect, test } from "bun:test"
import blockedQuestion from "../../../dochaus/workflows/lean/fixtures/status-blocked-question.json"
import blockedGate from "../../../dochaus/workflows/lean/fixtures/status-blocked-gate.json"
import completed from "../../../dochaus/workflows/lean/fixtures/status-completed.json"
import failed from "../../../dochaus/workflows/lean/fixtures/status-failed.json"
import { isLeanDraftingStatus, type LeanDocument, type LeanDraftingStatus, type LeanOpenItem } from "../src/api/ingest/lean"
import {
  addedDraftNote,
  attentionItems,
  capGroups,
  checklistFindingsCount,
  cleanTitle,
  currentStep,
  deliverableGroups,
  draftedGroups,
  type DraftedGroup,
  type DraftedOutput,
  finalCheckNote,
  groupOpenItems,
  leanSettledSummary,
  openItemClass,
  openItemsSentence,
  openItemTag,
  openQuestions,
  questionGroups,
  queuedNote,
  unresolvedChecklistFindings,
  runPosition,
  runStepName,
  stageElapsed,
  stepCause,
  stepGroups,
  stepName,
  updateProgressLabel,
} from "../src/components/drafting/model"
import { runHeadline } from "../src/components/drafting/RunHeader"
import { stageMilestones, stepActivityLabel } from "../src/components/narration/activity"
import { resultChecks } from "../src/components/narration/outcome"

function lean(value: unknown): LeanDraftingStatus {
  if (!isLeanDraftingStatus(value)) throw new Error("fixture is not a lean drafting status")
  return value
}

const item = (id: string, kind: string, status: LeanOpenItem["status"] = "open"): LeanOpenItem => ({ id, stage: "draft", kind, detail: id, status })

test("steps read as plain names, run-level and per document", () => {
  expect(["d1.draft", "d1.review", "d3.changes", "d3.apply", "finalize", "gate"].map(stepName)).toEqual([
    "Draft",
    "Review",
    "Reconcile amendments",
    "Apply changes",
    "Drafting notes",
    "Final check",
  ])
  expect(stepName("d2.cross_check")).toBe("Cross check")
})

test("equivalent topic labels keep separate decisions in one group and retain the first display label", () => {
  const question = lean(blockedQuestion).questions[0]!
  const groups = questionGroups([
    { ...question, id: "q1", group: "Payment terms" },
    { ...question, id: "q2", group: "payment   terms" },
    { ...question, id: "q3", group: "Café terms" },
    { ...question, id: "q4", group: "CAFE\u0301 TERMS" },
  ])
  expect(groups.map((group) => [group.label, group.questions.map((row) => row.id)])).toEqual([
    ["Payment terms", ["q1", "q2"]],
    ["Café terms", ["q3", "q4"]],
  ])
})

test("planned titles lose markdown ticks; an untitled document gets its position, or Draft before any plan", () => {
  expect(cleanTitle("  `Cover`   memo ")).toBe("Cover memo")
  const status = lean({ ...completed, plan: null })
  expect(stepGroups(status).map((group) => group.title)).toEqual(["Draft", "Finishing"])
  const untitled = lean({ ...completed, plan: { documents: [{ id: "d9", title: "Other", purpose: "" }] } })
  expect(stepGroups(untitled)[0]!.title).toBe("Document 1")
  const ticked = lean({ ...completed, plan: { documents: [{ id: "d1", title: "`Note`", purpose: "" }] } })
  expect(stepGroups(ticked)[0]!.title).toBe("Note")
})

test("the current step is the first not done, with its place in the walk", () => {
  expect(currentStep(lean(failed))).toEqual({ key: "d1.draft", position: 1, total: 6 })
  expect(currentStep(lean(blockedGate))).toEqual({ key: "gate", position: 6, total: 6 })
  expect(currentStep(lean(completed))).toBeUndefined()
})

test("a step's elapsed time runs to its finish, or to now while it runs", () => {
  const row = { status: "running" as const, attempts: 1, started_at: "2026-10-11T09:00:00.000Z", finished_at: null }
  expect(stageElapsed(row, Date.parse("2026-10-11T09:04:12.000Z"))).toBe(252_000)
  expect(stageElapsed({ ...row, status: "done", finished_at: "2026-10-11T09:00:30.000Z" }, Date.now())).toBe(30_000)
  expect(stageElapsed({ ...row, status: "pending", started_at: null }, Date.now())).toBeUndefined()
})

test("a stop's cause lands on the step it stopped at", () => {
  expect(stepCause(lean(failed), "d1.draft")).toStartWith("outline: checker failed")
  expect(stepCause(lean(failed), "d1.review")).toBeUndefined()
  expect(stepCause(lean(blockedGate), "gate")).toBe("1 open item still to resolve.")
})

test("open items group by kind in first-seen order and cap across groups", () => {
  const items = [item("a", "quote_miss"), item("b", "placeholder"), item("c", "quote_miss"), item("d", "new_kind"), item("e", "placeholder")]
  const groups = groupOpenItems(items)
  expect(groups.map((group) => [group.kind, group.items.map((row) => row.id)])).toEqual([
    ["quote_miss", ["a", "c"]],
    ["placeholder", ["b", "e"]],
    ["new_kind", ["d"]],
  ])
  expect(capGroups(groups, 3).map((group) => [group.kind, group.items.length, group.total])).toEqual([
    ["quote_miss", 2, 2],
    ["placeholder", 1, 2],
  ])
  expect(capGroups(groups, 5).flatMap((group) => group.items)).toHaveLength(5)
})

test("attention items follow the final check once it ran for this pass", () => {
  expect(attentionItems(lean(blockedGate)).map((row) => row.status)).toEqual(["open"])
  const disclosed = lean({ ...completed, gate: { ...completed.gate, items: [item("x", "placeholder", "disclosed"), item("y", "omission", "resolved")] } })
  expect(attentionItems(disclosed).map((row) => row.id)).toEqual(["x"])
})

function withFindings(unresolved: LeanOpenItem[], others: LeanOpenItem[] = []) {
  const items = [...others, ...unresolved]
  return lean({ ...completed, gate: { ...completed.gate, delivery: "review_draft_with_findings", disclosed: items, items } })
}

const unresolvedFinding = (n: number, checklist_class?: string): LeanOpenItem => ({
  id: `checklist-check:checklist_unresolved:${n}`,
  stage: "checklist-check",
  kind: "checklist_unresolved",
  detail: `Checklist requirement not satisfied: Requirement ${n}. Finding: Still absent after the repair round.`,
  status: "disclosed",
  ref: { checklist_item: `item-${n}`, ...(checklist_class ? { checklist_class } : {}), document: "d1" },
})

test("unresolved checklist findings sort first, stay uncapped ahead of other items, and map tags and classes", () => {
  const omission = unresolvedFinding(1, "omission")
  const conflict = unresolvedFinding(2, "source_conflict")
  const open: LeanOpenItem = { ...item("c1", "checklist_open", "disclosed"), stage: "checklist-check" }
  const others = [item("p1", "placeholder", "disclosed"), open, item("o1", "omission", "disclosed")]
  const groups = groupOpenItems([...others, omission, conflict])
  expect(groups.map((group) => group.kind)).toEqual(["checklist_unresolved", "placeholder", "checklist_open", "omission"])
  expect(capGroups(groups, 2).map((group) => [group.kind, group.items.length])).toEqual([["checklist_unresolved", 2]])
  expect(openItemTag(omission)).toBe("Not satisfied")
  expect(openItemTag(open)).toBe("Left open")
  expect(openItemTag(others[0]!)).toBe("In the notes")
  expect(
    ["omission", "source_conflict", "not_inspectable", "missing_inspection", "departure", undefined].map((value) => openItemClass(unresolvedFinding(9, value))),
  ).toEqual([
    "left out of the draft",
    "conflicts with another source requirement",
    "cannot be established from document text",
    "not established by the check",
    "deliberate departure",
    undefined,
  ])
  expect(openItemClass({ ...open, ref: { checklist_class: "omission" } })).toBeUndefined()
  expect(leanSettledSummary([...others, omission, conflict])).toContain("2 not satisfied")
  expect(leanSettledSummary([open, others[0]!])).toContain("1 left open")
})

test("a review draft with unresolved checklist findings distinguishes itself from a clean completion and a plain review draft", () => {
  const now = Date.parse(completed.updated_at)
  const two = withFindings([unresolvedFinding(1, "omission"), unresolvedFinding(2, "source_conflict")], [item("p1", "placeholder", "disclosed"), item("o1", "omission", "disclosed")])
  const one = withFindings([unresolvedFinding(1)])
  const plain = lean({ ...completed, gate: { ...completed.gate, delivery: "review_draft", items: [item("p1", "placeholder", "disclosed")] } })

  expect(unresolvedChecklistFindings(two).map((row) => row.id)).toEqual([unresolvedFinding(1).id, unresolvedFinding(2).id])
  expect(checklistFindingsCount(two)).toBe(2)
  expect(checklistFindingsCount(plain)).toBeUndefined()
  expect(checklistFindingsCount(lean(completed))).toBeUndefined()
  expect(checklistFindingsCount(lean({ ...completed, gate: { ...completed.gate, delivery: "review_draft_with_findings" } }))).toBe(0)

  expect(runHeadline(two, now)).toMatchObject({ tone: "done", title: "Review draft — 2 unresolved checklist findings" })
  expect(runHeadline(one, now).title).toBe("Review draft — 1 unresolved checklist finding")
  expect(runHeadline(lean({ ...completed, gate: { ...completed.gate, delivery: "review_draft_with_findings" } }), now).title).toBe(
    "Review draft — unresolved checklist findings",
  )
  expect(runHeadline(plain, now).title).toBe("Ready for review")
  expect(runHeadline(lean(completed), now).title).toBe("Ready for review")

  expect(openItemsSentence(two)).toContain("2 unresolved checklist findings")
  expect(openItemsSentence(one)).toContain("1 unresolved checklist finding")
  expect(finalCheckNote(two)).toContain("2 unresolved checklist findings")
  expect(finalCheckNote(one)).toContain("1 unresolved checklist finding")
  expect(finalCheckNote(plain)).toContain("review draft")
  expect(finalCheckNote(lean({ ...two, changed_after_check: { since: completed.gate.checked_at, documents: ["Supply note.docx"], missing: [] } }))).toContain(
    "Changed after final check",
  )
})

test("drafted outputs nest the redline and notes under a single draft", () => {
  expect(draftedGroups(lean(completed))).toEqual([
    {
      main: { name: "Supply note.docx", role: "draft" },
      extras: [
        { name: "Supply note - redline.docx", role: "redline" },
        { name: "Supply note - drafting notes.docx", role: "notes" },
      ],
    },
  ])
  expect(draftedGroups(lean(blockedGate))).toEqual([])
  expect(draftedGroups(undefined)).toEqual([])

  const body = completed.documents.documents[0]!
  const redline = completed.documents.documents[1]!
  const two = lean({
    ...completed,
    documents: {
      ...completed.documents,
      documents: [
        ...completed.documents.documents,
        { ...body, document: "d2", deliverable: "Second.docx" },
        { ...redline, document: "d2", deliverable: "Second - redline.docx" },
      ],
    },
  })
  expect(draftedGroups(two).map((group) => [group.main.name, group.extras.map((row) => row.name)])).toEqual([
    ["Supply note.docx", ["Supply note - redline.docx"]],
    ["Second.docx", ["Second - redline.docx"]],
    ["Supply note - drafting notes.docx", []],
  ])
})

function memoRun({
  combined = false,
  body = "Supply note - clean.docx",
  reported = body,
  file = "Supply note.docx",
  saved = file,
  rows: more = [],
}: { combined?: boolean; body?: string; reported?: string; file?: string; saved?: string | null; rows?: LeanDocument[] } = {}) {
  const rows = completed.documents.documents.map((row) => {
    if (row.kind === "drafting_notes") return { ...row, deliverable: "Supply note - cover note.docx" }
    if (row.kind === "body" && combined) return { ...row, deliverable: body }
    return row
  })
  const savedRow: LeanDocument[] =
    combined && saved ? [{ kind: "combined", document: "d1", title: "Supply note", path: saved, sha256: "c".repeat(64), deliverable: saved }] : []
  return lean({
    ...completed,
    memo: "cover note",
    documents: { ...completed.documents, documents: [...rows, ...more, ...savedRow] },
    ...(combined
      ? { combined_output: { document: "d1", name: "Supply note", file, body: reported, notes: "Supply note - cover note.docx" } }
      : {}),
  })
}

const memoNotes: DraftedOutput = { name: "Supply note - cover note.docx", role: "notes" }

test("a requested memo groups as the run's notes by its row kind and keeps its own name", () => {
  const status = memoRun()
  const groups: DraftedGroup[] = [{ main: { name: "Supply note.docx", role: "draft" }, extras: [{ name: "Supply note - redline.docx", role: "redline" }, memoNotes] }]
  expect(draftedGroups(status)).toStrictEqual(groups)
  expect(deliverableGroups(["Supply note.docx", "Supply note - redline.docx", "Supply note - cover note.docx"], status)).toStrictEqual(groups)
})

test("a combined document's one file leads its group only once both its parts are delivered", () => {
  const combined = memoRun({ combined: true })
  const body: DraftedOutput = { name: "Supply note - clean.docx", role: "draft" }
  const redline: DraftedOutput = { name: "Supply note - redline.docx", role: "redline" }
  expect(draftedGroups(combined)).toStrictEqual([{ main: body, extras: [redline, memoNotes], combined: "Supply note.docx" }])
  expect(draftedGroups(memoRun({ combined: true, body: "Supply note - clean (2).docx" }))[0]!.combined).toBe("Supply note.docx")

  const without = (kind: string) =>
    lean({ ...combined, documents: { ...combined.documents!, documents: combined.documents!.documents.filter((row) => row.kind !== kind) } })
  expect(draftedGroups(without("drafting_notes"))[0]).toStrictEqual({ main: body, extras: [redline] })
  expect(draftedGroups(without("body")).some((group) => "combined" in group)).toBe(false)
  expect(draftedGroups(without("combined"))).toStrictEqual([{ main: body, extras: [redline, memoNotes] }])
  expect(draftedGroups(lean({ ...combined, documents: { ...combined.documents!, status: "blocked" } }))).toEqual([])
  expect(draftedGroups(memoRun({ combined: true, body: "Supply note.docx", reported: "Supply note - clean.docx", saved: null }))[0]).not.toHaveProperty("combined")

  const names = ["Supply note - clean.docx", "Supply note - redline.docx", "Supply note - cover note.docx"]
  expect(deliverableGroups(names, combined)).toStrictEqual([{ main: body, extras: [redline, memoNotes] }])
  expect(deliverableGroups([...names, "Supply note.docx"], combined)).toStrictEqual([{ main: body, extras: [redline, memoNotes], combined: "Supply note.docx" }])
  expect(deliverableGroups([...names, "Supply note.docx"], undefined).filter((group) => group.main.name === "Supply note.docx")).toHaveLength(1)
})

test("a combined file belongs to the exact body and notes ingest names, never to another \"- clean\" row", () => {
  const stray: LeanDocument = { ...(completed.documents.documents[0] as LeanDocument), document: "d2", title: "Older supply note", deliverable: "Supply note - clean.docx" }
  const status = memoRun({ combined: true, body: "Supply note - clean (2).docx", rows: [stray] })
  const groups = draftedGroups(status)
  expect(groups.filter((group) => group.combined).map((group) => [group.main.name, group.combined])).toEqual([["Supply note - clean (2).docx", "Supply note.docx"]])
  expect(groups.find((group) => group.main.name === "Supply note - clean.docx")).not.toHaveProperty("combined")

  const reportedAs = (output: Partial<NonNullable<LeanDraftingStatus["combined_output"]>>) =>
    draftedGroups(lean({ ...status, combined_output: { ...status.combined_output!, ...output } })).some((group) => group.combined)
  expect(reportedAs({})).toBe(true)
  expect(reportedAs({ document: "d2" })).toBe(false)
  expect(reportedAs({ body: "Supply note - clean.docx" })).toBe(false)
  expect(reportedAs({ body: "supply note - clean (2).docx" })).toBe(false)
  expect(reportedAs({ body: "Supply note - clean (3).docx" })).toBe(false)
  expect(reportedAs({ notes: "Supply note - drafting notes.docx" })).toBe(false)
  expect(draftedGroups(lean({ ...status, combined_output: null })).some((group) => group.combined)).toBe(false)
  expect(isLeanDraftingStatus({ ...status, combined_output: { name: "Supply note", file: "Supply note.docx" } })).toBe(false)
})

function authoredMemoRun({ file = "Supply note.docx", saved = file, rows: more = [] }: { file?: string; saved?: string | null; rows?: LeanDocument[] } = {}) {
  const [body, redline, ...notes] = completed.documents.documents
  const memoRows: LeanDocument[] = [
    { kind: "body", document: "d2", title: "Cover memo", path: ".dochaus/lean/d2/body-applied.docx", sha256: "e".repeat(64), deliverable: "Cover memo.docx" },
    { kind: "redline", document: "d2", title: "Cover memo", path: ".dochaus/lean/d2/redline.docx", sha256: "f".repeat(64), deliverable: "Cover memo - redline.docx" },
  ]
  const savedRow: LeanDocument[] = saved ? [{ kind: "combined", document: "d1", title: "Supply note", path: saved, sha256: "c".repeat(64), deliverable: saved }] : []
  return lean({
    ...completed,
    memo: "cover memo",
    plan: { documents: [{ id: "d1", title: "Supply note", purpose: "" }, { id: "d2", title: "Cover memo", purpose: "" }] },
    documents: { ...completed.documents, documents: [{ ...body!, deliverable: "Supply note - clean.docx" }, redline!, ...memoRows, ...notes, ...more, ...savedRow] },
    combined_output: { document: "d1", name: "Supply note", file, body: "Supply note - clean.docx", notes: "Cover memo.docx" },
  })
}

const authoredDraft: DraftedOutput = { name: "Supply note - clean.docx", role: "draft" }
const authoredRedline: DraftedOutput = { name: "Supply note - redline.docx", role: "redline" }
const authoredMemo: DraftedOutput = { name: "Cover memo.docx", role: "draft" }
const authoredMemoRedline: DraftedOutput = { name: "Cover memo - redline.docx", role: "redline" }
const authoredNotes: DraftedOutput = { name: "Supply note - drafting notes.docx", role: "notes" }
const authoredNames = ["Supply note - clean.docx", "Supply note - redline.docx", "Cover memo.docx", "Cover memo - redline.docx", "Supply note - drafting notes.docx", "Supply note.docx"]

test("a combined file whose memo part is the authored memo body is the run's current combined file", () => {
  const status = authoredMemoRun()
  const groups: DraftedGroup[] = [
    { main: authoredDraft, extras: [authoredRedline], combined: "Supply note.docx" },
    { main: authoredMemo, extras: [authoredMemoRedline] },
    { main: authoredNotes, extras: [] },
  ]
  expect(draftedGroups(status)).toStrictEqual(groups)
  expect(deliverableGroups(authoredNames, status)).toStrictEqual(groups)
  expect(deliverableGroups(authoredNames, status).flatMap((group) => [group.main, ...group.extras]).filter((output) => output.role === "earlier_combined")).toEqual([])
  const withoutDraft = deliverableGroups(authoredNames.filter((name) => name !== "Supply note - clean.docx" && name !== "Supply note - redline.docx"), status)
  expect(withoutDraft[0]).toStrictEqual({ main: { name: "Supply note.docx", role: "combined" }, extras: [] })

  const reportedAs = (output: Partial<NonNullable<LeanDraftingStatus["combined_output"]>>) =>
    draftedGroups(lean({ ...status, combined_output: { ...status.combined_output!, ...output } })).some((group) => group.combined)
  expect(reportedAs({})).toBe(true)
  expect(reportedAs({ notes: "Cover memo (2).docx" })).toBe(false)
  expect(reportedAs({ notes: "cover memo.docx" })).toBe(false)
  expect(reportedAs({ notes: "Cover memo - redline.docx" })).toBe(false)
  expect(reportedAs({ body: "Supply note - clean (2).docx" })).toBe(false)
  expect(reportedAs({ document: "d2" })).toBe(false)

  const dropped = lean({ ...status, documents: { ...status.documents!, documents: status.documents!.documents.filter((row) => !(row.kind === "body" && row.document === "d2")) } })
  expect(draftedGroups(dropped).some((group) => group.combined)).toBe(false)
  expect(deliverableGroups(authoredNames, dropped).flatMap((group) => [group.main, ...group.extras]).find((output) => output.name === "Supply note.docx")?.role).toBe("earlier_combined")
})

test("an authored-memo combined file keeps the current file on its draft and an earlier kept file beside it", () => {
  const earlier: LeanDocument = { kind: "combined", document: "d1", title: "Supply note", path: "Supply note.docx", sha256: "b".repeat(64), deliverable: "Supply note.docx" }
  const status = authoredMemoRun({ file: "Supply note (2).docx", rows: [earlier] })
  const groups: DraftedGroup[] = [
    { main: authoredDraft, extras: [authoredRedline], combined: "Supply note (2).docx" },
    { main: authoredMemo, extras: [authoredMemoRedline] },
    { main: authoredNotes, extras: [] },
    { main: { name: "Supply note.docx", role: "earlier_combined" }, extras: [] },
  ]
  expect(draftedGroups(status)).toStrictEqual(groups)
  expect(deliverableGroups([...authoredNames, "Supply note (2).docx"], status)).toStrictEqual(groups)
})

test("addedDraftNote explains an extra planned draft only when the request named other files", () => {
  const second: LeanDocument = { kind: "body", document: "d2", title: "Second", path: ".dochaus/lean/d2/body-applied.docx", sha256: "d".repeat(64), deliverable: "Second.docx" }
  const run = (plan: Array<{ id: string; title: string; purpose: string; file?: string }>) =>
    lean({ ...completed, documents: { ...completed.documents, documents: [...completed.documents.documents, second] }, plan: { documents: plan } })
  const named = { id: "d1", title: "Supply note", purpose: "Sets out the supply terms.", file: "Supply note.docx" }
  const added = { id: "d2", title: "Second", purpose: "Records the approval the first document relies on" }
  const status = run([named, added])
  expect(addedDraftNote(status, "Second.docx")).toBe("Prepared in addition to the file you named. Records the approval the first document relies on.")
  expect(addedDraftNote(status, "Supply note.docx")).toBeUndefined()
  expect(addedDraftNote(run([named, { ...added, file: "Second.docx" }]), "Second.docx")).toBeUndefined()
  expect(addedDraftNote(run([{ id: "d1", title: "Supply note", purpose: "Sets out the supply terms." }, added]), "Second.docx")).toBeUndefined()
})

test("unsaved, mismatched, and earlier combined rows classify as earlier_combined rather than current combined or draft", () => {
  const unsaved: LeanDocument = { kind: "combined", document: "d1", title: "Supply note", path: "Supply note (2).docx", sha256: "a".repeat(64), deliverable: null }
  const failedSave = memoRun({ combined: true, body: "Supply note - clean (2).docx", saved: null, rows: [unsaved] })
  expect(draftedGroups(failedSave).some((group) => group.combined)).toBe(false)
  const failedNames = ["Supply note - clean.docx", "Supply note - clean (2).docx", "Supply note - redline.docx", "Supply note - cover note.docx", "Supply note.docx"]
  const failedGroups = deliverableGroups(failedNames, failedSave)
  expect(failedGroups.some((group) => group.combined)).toBe(false)
  expect(failedGroups.flatMap((group) => [group.main, ...group.extras]).filter((output) => output.role === "combined")).toEqual([])

  expect(draftedGroups(memoRun({ combined: true, saved: "Supply note (3).docx" })).some((group) => group.combined)).toBe(false)

  const earlier: LeanDocument = { kind: "combined", document: "d1", title: "Supply note", path: "Supply note.docx", sha256: "b".repeat(64), deliverable: "Supply note.docx" }
  const replaced = memoRun({ combined: true, body: "Supply note - clean (2).docx", file: "Supply note (2).docx", rows: [earlier] })
  const body: DraftedOutput = { name: "Supply note - clean (2).docx", role: "draft" }
  const redline: DraftedOutput = { name: "Supply note - redline.docx", role: "redline" }
  const kept: DraftedGroup = { main: { name: "Supply note.docx", role: "earlier_combined" }, extras: [] }
  expect(draftedGroups(replaced)).toStrictEqual([{ main: body, extras: [redline, memoNotes], combined: "Supply note (2).docx" }, kept])
  const allNames = [...failedNames, "Supply note (2).docx"]
  expect(deliverableGroups(allNames, replaced)).toStrictEqual([
    { main: body, extras: [redline, memoNotes], combined: "Supply note (2).docx" },
    kept,
    { main: { name: "Supply note - clean.docx", role: "draft" }, extras: [] },
  ])
  const unlisted = memoRun({ combined: true, body: "Supply note - clean (2).docx", file: "Supply note (2).docx" })
  expect(deliverableGroups(allNames, unlisted).flatMap((group) => [group.main, ...group.extras]).find((output) => output.name === "Supply note.docx")?.role).not.toBe("combined")

  const withoutBody = deliverableGroups(["Supply note - redline.docx", "Supply note - cover note.docx", "Supply note.docx"], memoRun({ combined: true }))
  const outputs = withoutBody.flatMap((group) => [group.main, ...group.extras, ...(group.combined ? [{ name: group.combined, role: "combined" }] : [])])
  expect(outputs.filter((output) => output.name === "Supply note.docx")).toEqual([{ name: "Supply note.docx", role: "combined" }])

  const names = ["Supply note - clean.docx", "Supply note - redline.docx", "Supply note - cover note.docx", "Supply note.docx"]
  const roleOf = (status: LeanDraftingStatus) =>
    deliverableGroups(names, status).flatMap((group) => [group.main, ...group.extras]).find((output) => output.name === "Supply note.docx")?.role
  const unsavedSame: LeanDocument = { kind: "combined", document: "d1", title: "Supply note", path: "Supply note.docx", sha256: "a".repeat(64), deliverable: null }
  const held = lean({ ...memoRun({ combined: true, saved: null, rows: [unsavedSame] }), combined_output: null })
  const stopped = memoRun({ combined: true, saved: null, rows: [{ ...unsavedSame, deliverable: "Supply note.docx" }] })
  const blocked = lean({ ...stopped, status: "blocked", documents: { ...stopped.documents!, status: "blocked" }, combined_output: null })
  const unsavedFailed = lean({ ...held, status: "failed", documents: { ...held.documents!, status: "blocked" } })
  for (const status of [held, blocked, unsavedFailed]) {
    expect(draftedGroups(status).some((group) => group.combined)).toBe(false)
    expect(roleOf(status)).toBe("earlier_combined")
  }
  expect(roleOf(lean({ ...held, documents: { ...held.documents!, documents: held.documents!.documents.filter((row) => row.kind !== "combined") } }))).toBe("draft")
})

test("a requested memo renames drafting notes across steps, summaries, milestones and checks; null matches absent", () => {
  const items = [item("p1", "placeholder", "disclosed"), item("o1", "omission", "disclosed"), item("o2", "omission", "disclosed")]
  const plain = lean({ ...completed, gate: { ...completed.gate, items } })
  const memo = lean({ ...plain, memo: "cover memo" })
  const finalize = (status: LeanDraftingStatus) => stepGroups(status).flatMap((group) => group.steps).find((step) => step.key === "finalize")!.name
  expect([stepName("finalize"), runStepName(plain, "finalize"), runStepName({ memo: null }, "finalize"), finalize(plain)]).toEqual(
    ["Drafting notes", "Drafting notes", "Drafting notes", "Drafting notes"],
  )
  expect([runStepName(memo, "finalize"), finalize(memo)]).toEqual(["Cover memo", "Cover memo"])
  expect(runStepName(memo, "gate")).toBe("Final check")
  expect(leanSettledSummary(items)).toContain("drafting notes")
  expect(leanSettledSummary(items, "cover memo")).toContain("cover memo")
  expect(resultChecks(plain).map((check) => check.text).some((t) => t.includes("drafting notes"))).toBe(true)
  expect(resultChecks(memo).map((check) => check.text).some((t) => t.includes("cover memo"))).toBe(true)
  const milestone = (status: LeanDraftingStatus) => stageMilestones(status).find((line) => line.step === "finalize")!.text
  expect([milestone(plain), milestone(memo)]).toEqual(["Drafting notes written", "Cover memo written"])
  const running = (status: LeanDraftingStatus) =>
    lean({ ...status, status: "running", in_flight: true, stages: { ...status.stages, finalize: { ...status.stages.finalize!, status: "running", finished_at: null }, gate: { ...status.stages.gate!, status: "pending", started_at: null, finished_at: null } } })
  expect([stepActivityLabel(running(plain), "finalize"), stepActivityLabel(running(memo), "finalize")]).toEqual(["Writing the drafting notes", "Writing the cover memo"])
  expect([runPosition(running(plain))!.text, runPosition(running(memo))!.text]).toEqual(["Finishing · Drafting notes", "Finishing · Cover memo"])

  for (const raw of [completed, blockedGate, failed, blockedQuestion, { ...completed, gate: { ...completed.gate, items } }]) {
    const absent = lean(raw)
    const none = lean({ ...raw, memo: null })
    expect(stageMilestones(none)).toStrictEqual(stageMilestones(absent))
    expect(resultChecks(none)).toStrictEqual(resultChecks(absent))
    expect(stepGroups(none)).toStrictEqual(stepGroups(absent))
  }
})

test("a run without a combined file treats null and absent combined_output identically", () => {
  const plain = lean(completed)
  const reported = lean({ ...completed, combined_output: null })
  expect(draftedGroups(reported)).toStrictEqual(draftedGroups(plain))
  const names = ["Old.docx", "Old - drafting notes.docx", "Supply note.docx", "Supply note - redline.docx", "Supply note - drafting notes.docx"]
  expect(deliverableGroups(names, reported)).toStrictEqual(deliverableGroups(names, plain))
})

test("the current position names the document of several, or the finishing step", () => {
  const pending = { status: "pending", attempts: 0, started_at: null, finished_at: null }
  const done = completed.stages["d1.draft"]
  const plan = { documents: [{ id: "d1", title: "Supply note", purpose: "" }, { id: "d2", title: "`Covering` letter", purpose: "" }] }
  const second = lean({
    ...completed,
    status: "running",
    in_flight: true,
    plan,
    stages: {
      "d1.draft": done, "d1.review": done, "d1.changes": done, "d1.apply": done,
      "d2.draft": { ...pending, status: "running" }, "d2.review": pending, "d2.changes": pending, "d2.apply": pending, finalize: pending, gate: pending,
    },
  })
  expect(runPosition(second)).toEqual({ text: "Document 2 of 2 · Draft", document: "Covering letter" })
  const finishing = lean({ ...second, stages: { ...second.stages, "d2.draft": done, "d2.review": done, "d2.changes": done, "d2.apply": done, finalize: { ...pending, status: "running" } } })
  expect(runPosition(finishing)).toEqual({ text: "Finishing · Drafting notes" })
})

test("open items read as one sentence, questions apart from placeholders", () => {
  const status = (items: LeanOpenItem[]) => lean({ ...blockedGate, open_items: items, gate: { ...blockedGate.gate, items } })
  expect(openItemsSentence(lean(completed))).toBe("No open items remain.")
  expect(openItemsSentence(lean(blockedQuestion))).toBe("Open items: 1 question for you.")
  expect(openItemsSentence(status([item("a", "placeholder"), item("b", "placeholder"), item("c", "quote")]))).toBe(
    "Open items: 2 placeholders to fill and 1 other item to check.",
  )
  expect(openItemsSentence(status([]))).toBeUndefined()
})

test("a draft changed or missing after the final check stops repeating the check's verdict", () => {
  const after = (documents: string[], missing: string[]) =>
    lean({ ...completed, changed_after_check: { since: completed.gate.checked_at, documents, missing } })
  const changed = after(["Supply note.docx"], [])
  const missing = after(["Supply note.docx"], ["Supply note.docx"])
  const both = after(["Supply note.docx", "Side letter.docx"], ["Supply note.docx"])

  expect(finalCheckNote(lean(completed))).toContain("ready for signature")
  expect(finalCheckNote(lean({ ...completed, changed_after_check: null }))).toContain("ready for signature")
  expect(finalCheckNote(lean({ ...completed, gate: { ...completed.gate, delivery: "review_draft" } }))).toContain("review draft")
  expect(finalCheckNote(changed)).toContain("Changed after final check")
  expect(finalCheckNote(lean({ ...changed, gate: { ...completed.gate, delivery: "review_draft" } }))).toContain("Changed after final check")
  expect(finalCheckNote(missing)).toContain("Missing after final check")
  expect(finalCheckNote(after(["Supply note.docx", "Side letter.docx"], ["Supply note.docx", "Side letter.docx"]))).toContain("delivered drafts")
  expect(finalCheckNote(both)).toContain("Changed or missing after final check")

  expect(runHeadline(changed, Date.now())).toMatchObject({ tone: "done", title: "Ready for review" })
  expect(runHeadline(changed, Date.now()).meta).toContain("Changed after final check")
  expect(runHeadline(lean(completed), Date.now()).meta).not.toContain("Changed after final check")
  expect(runHeadline(missing, Date.now()).meta).toContain("Missing after final check")
  expect(runHeadline(missing, Date.now()).meta).not.toContain("Changed after final check")
  expect(runHeadline(both, Date.now()).meta).toContain("Changed or missing after final check")
})

test("a stopped run's headline classifies failure, gate blocks, and provider setup/credit stops", () => {
  const stopped = runHeadline(lean(failed), Date.now())
  expect(stopped.title).toBe("Paused at draft")
  expect(stopped.meta).toContain("Its check did not pass")
  expect(stopped.meta.join(" ")).not.toContain("checker failed")
  expect(runHeadline(lean(blockedGate), Date.now()).meta[0]).toMatch(/^1 item to resolve, starting with [a-z]/)

  const setup = runHeadline(lean({ ...failed, stop_reason: "d1.draft: Model setup is not finished.", provider_stop: "setup" }), Date.now())
  expect(setup).toMatchObject({ tone: "paused", title: "Paused: Settings need attention" })
  const credit = runHeadline(lean({ ...failed, stop_reason: "d1.draft: The AI account has run out of credit.", provider_stop: "credit" }), Date.now())
  expect(credit).toMatchObject({ tone: "paused", title: "Paused: the AI account needs credit" })
})

test("evidence stages stay in preparation, document review and finishing order", () => {
  const pending = { status: "pending" as const, attempts: 0, started_at: null, finished_at: null }
  const stages = Object.fromEntries(["discrepancy", "intake", "research", "research-review", "spine", "d1.draft", "d1.review", "d1.review-independent", "d1.review-wording", "d1.review-client-disadvantage", "d1.change-plan", "d1.change-prepare", "d1.changes", "d1.change-review", "d1.apply", "finalize", "package-review", "gate"].map((key) => [key, pending]))
  const status = lean({ ...completed, status: "running", stages })
  expect(stepGroups(status).map((group) => [group.id, group.steps.map((step) => step.key)])).toEqual([
    ["preparation", ["discrepancy", "intake", "research", "research-review", "spine"]],
    ["d1", ["d1.draft", "d1.review", "d1.review-independent", "d1.review-wording", "d1.review-client-disadvantage", "d1.change-plan", "d1.change-prepare", "d1.changes", "d1.change-review", "d1.apply"]],
    ["run", ["finalize", "package-review", "gate"]],
  ])
  expect(runPosition(status)?.text).toBe("Preparation · Assess source discrepancies")
  expect(stepGroups(lean({ ...status, plan: null }))[1]?.title).toBe("Draft")
})

test("updates and clarifications track explicit execution progress and queued answers", () => {
  const update = { id: "u1", text: "Use short headings.", status: "included" as const }
  expect(updateProgressLabel(update)).toBe("Included in drafting instructions.")
  expect(updateProgressLabel({ ...update, status: "queued" })).toStartWith("Update received")
  expect(updateProgressLabel({ ...update, progress: "received" })).toStartWith("Update received")
  expect(updateProgressLabel({ ...update, progress: "applying" })).toStartWith("Update is being applied")
  expect(updateProgressLabel({ ...update, progress: "applied" })).toStartWith("Update applied")
  expect(updateProgressLabel({ ...update, progress: "requires-rework" })).toStartWith("Update requires rework")

  const clarification = { kind: "clarification" as const, id: "c1", text: "Inspect the supplied source.", status: "included" as const }
  expect(updateProgressLabel(clarification)).toContain("included for inspection")
  expect(updateProgressLabel({ ...clarification, progress: "applied" })).toStartWith("Clarification reviewed")
  expect(updateProgressLabel({ ...clarification, progress: "applying" })).toContain("being inspected")
  expect(updateProgressLabel({ ...clarification, status: "queued" })).toContain("received")

  const base = lean(blockedQuestion as unknown as LeanDraftingStatus)
  expect(queuedNote({ ...base, queued_answer_ids: [] })).toBeUndefined()
  expect(queuedNote({ ...base, queued_answer_ids: ["a"] })).toContain("Answer received")
  expect(queuedNote({ ...base, queued_answer_ids: ["a", "b"] })).toContain("Answers received")
})

test("an essential question during amendment planning keeps the run waiting for an answer", () => {
  const pending = { status: "pending" as const, attempts: 0, started_at: null, finished_at: null }
  const done = { ...pending, status: "done" as const }
  const status = lean({
    ...blockedQuestion,
    stop_reason: "d1.change-plan: essential question requires an answer",
    questions: blockedQuestion.questions.map((question) => ({ ...question, stage: "change-plan", document: "d1" })),
    stages: { "d1.review-client-disadvantage": done, "d1.change-plan": { ...pending, status: "blocked" }, "d1.change-prepare": pending, "d1.changes": pending, "d1.change-review": pending, "d1.apply": pending, finalize: pending, gate: pending },
  })
  expect(currentStep(status)).toEqual({ key: "d1.change-plan", position: 2, total: 8 })
  expect(runPosition(status)?.text).toBe("Assess review findings")
  expect(runHeadline(status, Date.now()).title).toBe("Waiting for your answers")
  expect(openQuestions(status)).toHaveLength(1)
  expect(stepCause(status, "d1.change-plan")).toBe("essential question requires an answer")
  expect(openQuestions(lean({ ...status, queued_answer_ids: status.questions.map((question) => question.id) }))).toEqual([])
})
