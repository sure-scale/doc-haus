import { expect, test } from "bun:test"
import { createElement } from "react"
import { MemoryRouter } from "react-router-dom"
import { renderToStaticMarkup } from "react-dom/server"
import blockedQuestion from "../../../dochaus/workflows/lean/fixtures/status-blocked-question.json"
import blockedGate from "../../../dochaus/workflows/lean/fixtures/status-blocked-gate.json"
import completed from "../../../dochaus/workflows/lean/fixtures/status-completed.json"
import failed from "../../../dochaus/workflows/lean/fixtures/status-failed.json"
import { isLeanDraftingStatus, type LeanDraftingStatus, type LeanOpenItem, type LeanQuestion } from "../src/api/ingest/lean"
import DocumentUpload from "../src/components/DocumentUpload"
import {
  attentionItems,
  capGroups,
  checklistFindingsCount,
  cleanTitle,
  currentStep,
  draftedGroups,
  finalCheckNote,
  groupOpenItems,
  leanSettledSummary,
  openItemClass,
  openItemsSentence,
  openItemTag,
  openQuestions,
  unresolvedChecklistFindings,
  runPosition,
  stageElapsed,
  stepCause,
  stepGroups,
  stepName,
  updateProgressLabel,
} from "../src/components/drafting/model"
import OpenItems from "../src/components/drafting/OpenItems"
import QuestionDock from "../src/components/drafting/QuestionDock"
import RunHeader, { runHeadline } from "../src/components/drafting/RunHeader"

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
  // A stage this build does not know still reads as words, not a key.
  expect(stepName("d2.cross_check")).toBe("Cross check")
})

test("planned titles lose markdown ticks; an untitled document gets its position", () => {
  expect(cleanTitle("  `Cover`   memo ")).toBe("Cover memo")
  const status = lean({ ...completed, plan: null })
  expect(stepGroups(status).map((group) => group.title)).toEqual(["Document 1", "Finishing"])
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
  expect(groups.map((group) => [group.label, group.items.map((row) => row.id)])).toEqual([
    ["Quotes not found in the sources", ["a", "c"]],
    ["Placeholders to fill", ["b", "e"]],
    ["New kind", ["d"]],
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

test("a checklist requirement left open shows in the delivered review draft's details list", () => {
  const actions = { busy: false, error: "", resuming: false, resume: async () => true, answer: async () => true }
  const detail =
    "Checklist requirement left open: Keep the note to about one page. Finding: Accepted body text measures {\"paragraphs\":28,\"words\":626,\"characters\":3900} do not clearly exceed about one page. Limitation: Rendered pagination was not verified."
  const open: LeanOpenItem = { id: "checklist-check:checklist_open:1", stage: "checklist-check", kind: "checklist_open", detail, status: "disclosed", ref: { document: "d1" } }
  const status = lean({ ...completed, gate: { ...completed.gate, delivery: "review_draft", disclosed: [open], items: [open, item("p1", "placeholder", "disclosed")] } })
  expect(attentionItems(status)).toContainEqual(open)
  expect(groupOpenItems([open])[0]!.label).toBe("Requirements left open by the checklist check")
  const html = renderToStaticMarkup(createElement(RunHeader, { status, actions, own: true, onOpen: () => {} }))
  expect(html).toContain("Requirements left open by the checklist check")
  expect(html).toContain("626")
  expect(html).toContain("Rendered pagination was not verified.")
  expect(html).toContain('<span class="open-item-tag">Left open</span>')
  // Only the placeholder is named in the drafting notes.
  expect(html.match(/<span class="open-item-tag">In the notes<\/span>/g)).toHaveLength(1)
  expect(html).toContain("a review draft, not yet ready for signature")
  expect(leanSettledSummary(status.gate!.items!)).toBe(
    "Final check settled 2 open items: 1 disclosed in the drafting notes, 1 left open by the checklist check.",
  )
})

/** A completed run whose final check delivered a review draft with unresolved
 * checklist findings: requirements the delivered draft does not satisfy. */
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

test("unresolved checklist findings sort first, read as not satisfied and carry their class in plain words", () => {
  const omission = unresolvedFinding(1, "omission")
  const conflict = unresolvedFinding(2, "source_conflict")
  const open: LeanOpenItem = { ...item("c1", "checklist_open", "disclosed"), stage: "checklist-check" }
  const others = [item("p1", "placeholder", "disclosed"), open, item("o1", "omission", "disclosed")]
  const groups = groupOpenItems([...others, omission, conflict])
  expect(groups.map((group) => group.kind)).toEqual(["checklist_unresolved", "placeholder", "checklist_open", "omission"])
  expect(groups[0]!.label).toBe("Unresolved checklist findings")
  // The list is capped; the findings are never the part cut off.
  expect(capGroups(groups, 2).map((group) => [group.kind, group.items.length])).toEqual([["checklist_unresolved", 2]])
  expect(openItemTag(omission)).toBe("Not satisfied")
  expect(openItemTag(open)).toBe("Left open")
  expect(openItemTag(others[0]!)).toBe("In the notes")
  expect(
    ["omission", "source_conflict", "not_inspectable", "missing_inspection", "departure", undefined].map((value) => openItemClass(unresolvedFinding(9, value))),
  ).toEqual([
    "not fixed by the repair round",
    "conflicts with another source requirement",
    "cannot be established from document text",
    "not inspected",
    "deliberate departure",
    undefined,
  ])
  // Only an unresolved finding carries a class sub-label.
  expect(openItemClass({ ...open, ref: { checklist_class: "omission" } })).toBeUndefined()
  expect(leanSettledSummary([...others, omission, conflict])).toBe(
    "Final check settled 5 open items: 2 not satisfied (unresolved checklist findings), 2 disclosed in the drafting notes, 1 left open by the checklist check.",
  )
})

test("a review draft with unresolved checklist findings reads differently from a clean completion and a plain review draft", () => {
  const actions = { busy: false, error: "", resuming: false, resume: async () => true, answer: async () => true }
  const render = (status: LeanDraftingStatus) =>
    renderToStaticMarkup(createElement(RunHeader, { status, actions, own: true, onOpen: () => {} }))
  const now = Date.parse(completed.updated_at)
  const two = withFindings([unresolvedFinding(1, "omission"), unresolvedFinding(2, "source_conflict")], [item("p1", "placeholder", "disclosed"), item("o1", "omission", "disclosed")])
  const one = withFindings([unresolvedFinding(1)])
  const plain = lean({ ...completed, gate: { ...completed.gate, delivery: "review_draft", items: [item("p1", "placeholder", "disclosed")] } })

  expect(unresolvedChecklistFindings(two).map((row) => row.id)).toEqual([unresolvedFinding(1).id, unresolvedFinding(2).id])
  expect(checklistFindingsCount(two)).toBe(2)
  expect(checklistFindingsCount(plain)).toBeUndefined()
  expect(checklistFindingsCount(lean(completed))).toBeUndefined()
  // The delivery alone marks the run, even if its items did not reach the status.
  expect(checklistFindingsCount(lean({ ...completed, gate: { ...completed.gate, delivery: "review_draft_with_findings" } }))).toBe(0)

  expect(runHeadline(two, now)).toMatchObject({ tone: "done", title: "Review draft — 2 unresolved checklist findings" })
  expect(runHeadline(one, now).title).toBe("Review draft — 1 unresolved checklist finding")
  expect(runHeadline(lean({ ...completed, gate: { ...completed.gate, delivery: "review_draft_with_findings" } }), now).title).toBe(
    "Review draft — unresolved checklist findings",
  )
  expect(runHeadline(plain, now).title).toBe("Ready for review")
  expect(runHeadline(lean(completed), now).title).toBe("Ready for review")

  expect(openItemsSentence(two)).toBe("Open items: 2 unresolved checklist findings, 1 placeholder to fill and 1 other item to check.")
  expect(openItemsSentence(one)).toBe("Open items: 1 unresolved checklist finding.")
  expect(finalCheckNote(two)).toBe(
    "Final workflow check: a review draft with 2 unresolved checklist findings, not yet ready for signature. The draft does not satisfy those requirements.",
  )
  expect(finalCheckNote(one)).toBe(
    "Final workflow check: a review draft with 1 unresolved checklist finding, not yet ready for signature. The draft does not satisfy that requirement.",
  )
  expect(finalCheckNote(plain)).toBe("Final workflow check passed: a review draft, not yet ready for signature.")
  // A draft changed after the check is not what it checked.
  expect(finalCheckNote(lean({ ...two, changed_after_check: { since: completed.gate.checked_at, documents: ["Supply note.docx"] } }))).toBe(
    "Changed after final check: review the draft.",
  )

  const html = render(two)
  expect(html).toContain("Review draft — 2 unresolved checklist findings")
  expect(html).not.toContain("Ready for review")
  expect(html).toContain('class="run-head-icon is-done is-findings"')
  expect(render(plain)).toContain('class="run-head-icon is-done"')
  expect(html).toContain('<span class="open-item-tag is-unresolved">Not satisfied</span>')
  expect(html).toContain('<span class="open-item-class">not fixed by the repair round</span>')
  expect(html).toContain('<span class="open-item-class">conflicts with another source requirement</span>')
  // The findings head the open items, ahead of what was recorded first.
  expect(html.indexOf("Unresolved checklist findings")).toBeLessThan(html.indexOf("Placeholders to fill"))
  expect(html).toContain("The draft does not satisfy those requirements.")
})

test("the open items list never caps away an unresolved checklist finding and counts them in its summary", () => {
  const findings = Array.from({ length: 6 }, (_, i) => unresolvedFinding(i + 1, "missing_inspection"))
  const others = [item("p1", "placeholder", "disclosed"), item("p2", "placeholder", "disclosed")]
  const html = renderToStaticMarkup(createElement(OpenItems, { items: [...others, ...findings], titles: new Map() }))
  expect(html).toContain('<span class="open-items-unresolved">6 not satisfied</span>')
  expect(html.match(/class="open-item-tag is-unresolved"/g)).toHaveLength(6)
  expect(html.match(/<span class="open-item-class">not inspected<\/span>/g)).toHaveLength(6)
  // The cap still folds the rest behind Show all.
  expect(html).not.toContain("Placeholders to fill")
  expect(html).toContain("Show all 8")
  const plain = renderToStaticMarkup(createElement(OpenItems, { items: others, titles: new Map() }))
  expect(plain).not.toContain("open-items-unresolved")
  expect(plain).not.toContain("Show all")
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

test("the document timeline orders files newest first and preserves origins and downloads", () => {
  const names = ["supply-terms.docx", "Supply note.docx", "Supply note - redline.docx", "Supply note - drafting notes.docx"]
  // Ingest records which files a drafting run delivered.
  const documents = names.map((name, id) => ({ id, name, doc_path: name, created_at: 1_760_000_000_000 + id * 60_000, role: id === 0 ? ("source" as const) : ("deliverable" as const) }))
  const render = (drafting?: LeanDraftingStatus, files = documents) =>
    renderToStaticMarkup(createElement(DocumentUpload, { matterId: "m", documents: files, onUploaded: () => {}, onView: () => {}, drafting }))
  const grouped = render(lean(completed))
  expect(grouped).toContain('class="matter-files-collection"')
  expect(grouped).toContain('aria-label="Document timeline, newest first"')
  for (let index = 1; index < names.length; index++) {
    expect(grouped.indexOf(names[index])).toBeLessThan(grouped.indexOf(names[index - 1]))
  }
  expect(grouped.match(/class="document-origin is-created">Created by Doc.Haus<\/span>/g)).toHaveLength(3)
  expect(grouped).toMatch(/class="doc-row is-uploaded">[\s\S]*supply-terms\.docx[\s\S]*class="document-origin">Uploaded<\/span>/)
  for (const name of names.slice(1)) {
    expect(grouped).toContain(`download="${name}"`)
    expect(grouped).toContain(`aria-label="Download ${name}"`)
  }
  expect(grouped).toContain('aria-label="Remove supply-terms.docx"')
  expect(grouped).not.toContain('download="supply-terms.docx"')
  expect(grouped.match(/<span class="matter-files-count">\d+<\/span>/g)).toEqual(['<span class="matter-files-count">4</span>'])
  expect(grouped.match(/class="matter-files-collection"/g)).toHaveLength(1)
  expect(grouped).toContain("Upload documents")
  // A file ingest does not mark as delivered is the lawyer's own, whatever its name.
  const sourcesOnly = render(lean(completed), documents.map((document) => ({ ...document, role: "source" as const })))
  expect(sourcesOnly.match(/class="document-origin">Uploaded/g)).toHaveLength(4)
  expect(sourcesOnly).not.toContain("Created by Doc.Haus")
  expect(sourcesOnly).not.toContain('class="drafted-group"')
  for (const name of names) expect(sourcesOnly).toContain(`aria-label="Remove ${name}"`)
})

test("a delivered run with questions left counts them in the header and docks them above the composer", () => {
  const actions = { busy: false, error: "", resuming: false, resume: async () => true, answer: async () => true }
  const render = (status: LeanDraftingStatus) =>
    renderToStaticMarkup(createElement(RunHeader, { status, actions, own: true, onOpen: () => {} }))
  const status = lean({ ...completed, questions: blockedQuestion.questions })
  const delivered = render(status)
  expect(delivered).toContain("Ready for review")
  expect(delivered).toContain("1 question for you")
  // The dock's Answer is the colour cue for the question; opening the draft stays secondary.
  expect(delivered).toMatch(/class="run-head-cta">Open draft</)
  expect(render(lean(completed))).toMatch(/class="primary run-head-cta">Open draft</)

  const dock = (questions: LeanQuestion[]) =>
    renderToStaticMarkup(
      createElement(QuestionDock, { matterId: "m", questions, answered: 0, titles: new Map(), busy: false, error: "", onSend: async () => true }),
    )
  const one = dock(openQuestions(status))
  expect(one).toContain("The drafting run has a question for you")
  expect(one).toMatch(/aria-haspopup="dialog"[^>]*>Answer</)
  // The dialog holds no question until it is opened.
  expect(one).toContain('<dialog class="question-dialog"')
  expect(one).not.toContain("Your answer")
  expect(dock([])).toBe("")
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

test("a draft changed after the final check stops repeating the check's verdict", () => {
  const actions = { busy: false, error: "", resuming: false, resume: async () => true, answer: async () => true }
  const render = (status: LeanDraftingStatus) =>
    renderToStaticMarkup(createElement(RunHeader, { status, actions, own: true, onOpen: () => {} }))
  const changed = lean({ ...completed, changed_after_check: { since: completed.gate.checked_at, documents: ["Supply note.docx"] } })
  expect(finalCheckNote(lean(completed))).toBe("Final workflow check passed: ready for signature.")
  expect(finalCheckNote(lean({ ...completed, gate: { ...completed.gate, delivery: "review_draft" } }))).toBe(
    "Final workflow check passed: a review draft, not yet ready for signature.",
  )
  expect(finalCheckNote(changed)).toBe("Changed after final check: review the draft.")
  expect(finalCheckNote(lean({ ...changed, gate: { ...completed.gate, delivery: "review_draft" } }))).toBe("Changed after final check: review the draft.")
  expect(finalCheckNote(lean({ ...completed, changed_after_check: null }))).toBe("Final workflow check passed: ready for signature.")
  expect(render(lean(completed))).toContain("Final workflow check passed: ready for signature.")
  const header = render(changed)
  expect(header).toContain("Changed after final check: review the draft.")
  expect(header).not.toContain("ready for signature")
  expect(runHeadline(changed, Date.now())).toMatchObject({ tone: "done", title: "Ready for review" })
  expect(runHeadline(changed, Date.now()).meta).toContain("Changed after final check")
  expect(runHeadline(lean(completed), Date.now()).meta).not.toContain("Changed after final check")
})

test("a stopped run's header says what happened in plain words; the checker's wording stays on its step", () => {
  const stopped = runHeadline(lean(failed), Date.now())
  expect(stopped.title).toBe("Paused at draft")
  expect(stopped.meta).toContain("Its check did not pass")
  expect(stopped.meta.join(" ")).not.toContain("checker failed")
  const gate = runHeadline(lean(blockedGate), Date.now())
  expect(gate.meta[0]).toMatch(/^1 item to resolve, starting with [a-z]/)
})

test("elsewhere the run's one line leads with the questions still owed", () => {
  const actions = { busy: false, error: "", resuming: false, resume: async () => true, answer: async () => true }
  const render = (status: LeanDraftingStatus) =>
    renderToStaticMarkup(createElement(MemoryRouter, null, createElement(RunHeader, { status, actions, own: false, onOpen: () => {} })))
  const asked = render(lean({ ...completed, questions: blockedQuestion.questions }))
  expect(asked).toContain("1 question for you")
  expect(asked).toContain("Open run")
  expect(render(lean(completed))).not.toContain("for you")
})


test("evidence stages stay in preparation, document review and finishing order", () => {
  const pending = { status: "pending" as const, attempts: 0, started_at: null, finished_at: null }
  const stages = Object.fromEntries(["discrepancy", "intake", "research", "research-review", "spine", "d1.draft", "d1.review", "d1.review-independent", "d1.review-wording", "d1.review-client-disadvantage", "d1.change-plan", "d1.change-prepare", "d1.changes", "d1.change-review", "d1.apply", "finalize", "package-review", "gate"].map(key => [key, pending]))
  const status = lean({ ...completed, status: "running", stages })
  expect(stepGroups(status).map(group => [group.id, group.steps.map(step => step.key)])).toEqual([
    ["preparation", ["discrepancy", "intake", "research", "research-review", "spine"]],
    ["d1", ["d1.draft", "d1.review", "d1.review-independent", "d1.review-wording", "d1.review-client-disadvantage", "d1.change-plan", "d1.change-prepare", "d1.changes", "d1.change-review", "d1.apply"]],
    ["run", ["finalize", "package-review", "gate"]],
  ])
  expect(runPosition(status)?.text).toBe("Preparation · Assess source discrepancies")
  expect(stepGroups(lean({ ...status, plan: null }))[1]?.title).toBe("Document 1")
  expect(stepName("d1.review-independent")).toBe("Independent function and fidelity review")
  expect(stepName("d1.review-wording")).toBe("Review problematic wording")
  expect(stepName("d1.review-client-disadvantage")).toBe("Review client disadvantage")
  expect(stepName("package-review")).toBe("Final semantic and package review")
})

test("an update claims application only from explicit execution evidence", () => {
  const update = { id: "u1", text: "Use short headings.", status: "included" as const }
  expect(updateProgressLabel(update)).toBe("Included in drafting instructions.")
  expect(updateProgressLabel({ ...update, status: "queued" })).toStartWith("Update received")
  expect(updateProgressLabel({ ...update, progress: "received" })).toStartWith("Update received")
  expect(updateProgressLabel({ ...update, progress: "applying" })).toStartWith("Update is being applied")
  expect(updateProgressLabel({ ...update, progress: "applied" })).toStartWith("Update applied")
  expect(updateProgressLabel({ ...update, progress: "requires-rework" })).toStartWith("Update requires rework")
})


test("an essential question during amendment planning keeps the run waiting for an answer", () => {
  const pending = { status: "pending" as const, attempts: 0, started_at: null, finished_at: null }
  const done = { ...pending, status: "done" as const }
  const status = lean({
    ...blockedQuestion,
    stop_reason: "d1.change-plan: essential question requires an answer",
    questions: blockedQuestion.questions.map(question => ({ ...question, stage: "change-plan", document: "d1" })),
    stages: { "d1.review-client-disadvantage": done, "d1.change-plan": { ...pending, status: "blocked" }, "d1.change-prepare": pending, "d1.changes": pending, "d1.change-review": pending, "d1.apply": pending, finalize: pending, gate: pending },
  })
  expect(currentStep(status)).toEqual({ key: "d1.change-plan", position: 2, total: 8 })
  expect(runPosition(status)?.text).toBe("Assess review findings")
  expect(runHeadline(status, Date.now()).title).toBe("Waiting for your answers")
  expect(openQuestions(status)).toHaveLength(1)
  expect(stepCause(status, "d1.change-plan")).toBe("essential question requires an answer")
  expect(["d1.change-plan", "d1.change-prepare", "d1.changes", "d1.change-review"].map(stepName)).toEqual([
    "Assess review findings", "Prepare amendments", "Reconcile amendments", "Check amended document",
  ])
  expect(openQuestions(lean({ ...status, queued_answer_ids: status.questions.map(question => question.id) }))).toEqual([])
})


test("clarification progress describes inspection instead of applied document changes", () => {
  const update = { kind: "clarification" as const, id: "c1", text: "Inspect the supplied source.", status: "included" as const }
  expect(updateProgressLabel(update)).toContain("included for inspection")
  expect(updateProgressLabel({ ...update, progress: "applied" })).toStartWith("Clarification reviewed")
  expect(updateProgressLabel({ ...update, progress: "applying" })).toContain("being inspected")
  expect(updateProgressLabel({ ...update, status: "queued" })).toContain("received")
})
