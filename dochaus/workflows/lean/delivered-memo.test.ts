import { afterEach, expect, test } from "bun:test"
import { readFileSync } from "node:fs"
import path from "node:path"
import { readDocxCommentParagraphs } from "../../lib/docx-comments"
import type { ChecklistFinding, ChecklistItem } from "./checklist-contract"
import { checklistRefinementSourceBasis } from "./checklist-flow"
import { requestedMemo } from "./deliverables"
import { memoMode, withoutInternalIDs, withoutRefMarkers } from "./delivered-memo"
import { deliveredNotes, finalGate } from "./gate"
import { readLeanPlan, writeLeanPlan } from "./plan"
import { openItem } from "./quotes"
import { runLean } from "./runner"
import { acceptStageReply } from "./stages"
import { leanFile, readLeanState, readLeanText, readOpenItems, updateLeanStage, writeLeanText, writeOpenItems, type OpenItem } from "./state"
import { AGENT, DRAFT, cleanupMatters, drafter, happyPath, matter, pidOf, registering, sameSession, scriptedModel, type Reply } from "./test-support"

// The drafting notes, a requested memo or not, are delivered at the gate
// without their "[ref: …]" markers, and the finding lines the gate composes
// for them name no internal id. The working drafting notes keep their markers
// for the notes check, the gate and the checklist check.

afterEach(cleanupMatters)

/** The plan as the outline leaves it for a requested memo. */
const requestMemo = (directory: string) =>
  writeLeanText(directory, "plan.json", JSON.stringify({ ...readLeanPlan(directory)!, memo: "cover memo", notes_file: "Supply note - cover memo.docx" }))
const docxText = (bytes: Uint8Array) => readDocxCommentParagraphs(bytes).map(row => row.paragraph_text).join("\n")

test("strips reference markers and whole known ids, tidies only the lines that changed and keeps unknown text", () => {
  const ids = ["research:research_gap:abc123def0", "apply:unapplied_amendment:0123456789", "a1", "w1", "p0123456789ab", "c1", "c2", "d1", "source-0123456789ab-3"]
  const text = [
    "# Memo",
    "",
    "No authority was found for the delivery period; confirm it before use. [ref: research:research_gap:abc123def0]",
    "- apply:unapplied_amendment:0123456789: the payment change could not be applied.",
    "- **c1** — The note states the delivery period.",
    "The proposed change (a1) was withheld [p0123456789ab].",
    "The clause c10 and item c1.u1.1 stay; see C1 and section 4.2.",
    "See items `c2`, `source-0123456789ab-3`.",
    "[ref: apply:unapplied_amendment:0123456789]",
    "  kept line with trailing spaces   ",
    "Windows line\r",
    "Both points stand [Ref: c1] [refs: c2, a1].",
    "Document d1 (see c1, c2) is ready.",
    "An unknown note [see q-delivery] stays.",
  ].join("\n")
  expect(withoutInternalIDs(text, ids).split("\n")).toEqual([
    "# Memo",
    "",
    "No authority was found for the delivery period; confirm it before use.",
    "- the payment change could not be applied.",
    "- The note states the delivery period.",
    "The proposed change was withheld.",
    "The clause c10 and item c1.u1.1 stay; see C1 and section 4.2.",
    "See items.",
    "  kept line with trailing spaces   ",
    "Windows line\r",
    "Both points stand.",
    "Document is ready.",
    "An unknown note [see q-delivery] stays.",
  ])
  // Memo prose loses its markers whole and nothing else.
  expect(withoutRefMarkers("# Cover memo\n\nConfirm the period. [ref: c1]\n- [ref: a1]\nItem c1 stays.\n")).toBe("# Cover memo\n\nConfirm the period.\nItem c1 stays.\n")
})

test("reads known ids joined by a slash or dash as the memo check does and removes them whole, with their joints", () => {
  const ids = ["a1", "a2", "a3", "w1"]
  const text = [
    "Three review changes were applied (a1-a3).",
    "- **a1–a3**: the payment changes stand.",
    "The memo leads the pack (a2/w1; a1–a2).",
    // Joined to a word, inside a longer id or in a path they are not ids.
    "Rows a1-x, xa1-a2, a12/a13 and path/a1 stay.",
  ].join("\n")
  expect(withoutInternalIDs(text, ids).split("\n")).toEqual([
    "Three review changes were applied.",
    "- the payment changes stand.",
    "The memo leads the pack.",
    "Rows a1-x, xa1-a2, a12/a13 and path/a1 stay.",
  ])
})

const REQUIREMENT = "source-0123456789ab-2"
const OPEN = "research:research_gap:abc123def0"

test("leaves text inside straight or curly double quotes byte for byte, ids, spacing and punctuation included", () => {
  const ids = ["a1", "a2", "a3", "d2", "i1", PID]
  const quoted = `Source: request: "Lots a1-a3 and Unit d2 are excluded; see note i1." and “${PID} ,  (a1)”`
  expect(withoutInternalIDs([
    quoted,
    `The body says "d2  and a1 go together" but the request says “a1/a2 as one”; a3 differs.`,
    // A quote of nothing but an id is that id, not a quotation.
    `See "a2".`,
  ].join("\n"), ids).split("\n")).toEqual([
    quoted,
    `The body says "d2  and a1 go together" but the request says “a1/a2 as one”; a proposed change differs.`,
    "See a proposed change.",
  ])
})

test("a quote of nothing but an id is that id in curly quotes as in straight ones", () => {
  // Was delivered as written: only a straight quote was read as its id.
  expect(withoutInternalIDs([`The draft omits “d2” and cites “a1” and “${PID}”.`, `“a1” was withheld; “a1 and d2” stay quoted.`].join("\n"), ["a1", "d2", PID]).split("\n")).toEqual([
    "The draft omits a document and cites a proposed change and a paragraph.",
    "A proposed change was withheld; “a1 and d2” stay quoted.",
  ])
})

test("an id followed by a word naming its kind in the run's number only labels that word, and gives way to \"the\" where nothing leads it", () => {
  const ids = ["a1", "a2", "w1", PID]
  const cases = [
    // Was "The proposed change change … the paragraph paragraph …".
    [`The a1 change was withheld and the ${PID} paragraph is silent.`, "The change was withheld and the paragraph is silent."],
    ["Its a1/a2 amendments and each w1 change differ.", "Its amendments and each change differ."],
    ["a1 change was withheld; a1/a2 changes were not.", "The change was withheld; the changes were not."],
    // In the other number the word is a verb.
    ["a1 changes the price; a1 and a2 change it; each a1 changes it.", "A proposed change changes the price; proposed changes change it; each proposed change changes it."],
  ]
  expect(withoutInternalIDs(cases.map(([text]) => text).join("\n"), ids).split("\n")).toEqual(cases.map(([, expected]) => expected))
})

test("after an auxiliary, a modal or \"let\", the word naming a single id's kind is a verb, not the id's label", () => {
  const ids = ["a1", "a2", "w1"]
  const cases = [
    // Was "Does the change the price? Would the change it?".
    ["Open decisions: Does a1 change the price? Would a1 or w1 change it?", "Open decisions: Does a proposed change change the price? Would a proposed change change it?"],
    ["Did a1 change it, and could w1 change it?", "Did a proposed change change it, and could a proposed change change it?"],
    ["It must not let a1 change it; should w1 change it?", "It must not let a proposed change change it; should a proposed change change it?"],
    // A plural kind word, or one after a name, still only takes the label.
    ["Do a1/a2 changes apply? Under the Will a1 change applies.", "Do the changes apply? Under the Will change applies."],
  ]
  expect(withoutInternalIDs(cases.map(([text]) => text).join("\n"), ids).split("\n")).toEqual(cases.map(([, expected]) => expected))
})

test("an id directly after a name is that name's and is named bare; after a verb or a joining word it keeps its article", () => {
  const ids = ["a1", PID]
  const cases = [
    // Was "The Lease a paragraph omits it; under the Deed of Variation a paragraph …".
    [`The Lease ${PID} omits it; under the Deed of Variation ${PID} the period differs.`, "The Lease paragraph omits it; under the Deed of Variation paragraph the period differs."],
    [`Each Supply note ${PID} and the Acme a1 differ.`, "Each Supply note paragraph and the Acme proposed change differ."],
    [`The Lease omits a1; the Lease omitted ${PID}; the Lease and a1 differ.`, "The Lease omits a proposed change; the Lease omitted a paragraph; the Lease and a proposed change differ."],
  ]
  expect(withoutInternalIDs(cases.map(([text]) => text).join("\n"), ids).split("\n")).toEqual(cases.map(([, expected]) => expected))
})

test("a preposition, conjunction or auxiliary after a name ends it, so an id after that word keeps its article", () => {
  const ids = ["a1", "w1", PID]
  const cases = [
    // Was "The Tenant under paragraph pays; the Lease unless proposed change …".
    [`The Tenant under ${PID} pays; the Lease unless a1 is applied differs.`, "The Tenant under a paragraph pays; the Lease unless a proposed change is applied differs."],
    [`The Buyer within ${PID} and the Seller after a1 agree; the Lease before w1 and the Deed until a1 differ.`, "The Buyer within a paragraph and the Seller after a proposed change agree; the Lease before a proposed change and the Deed until a proposed change differ."],
    [`The Seller had a1 applied; the Buyer may ${PID} cite; the Lease if w1 is withheld differs; the Deed but a1 differs.`, "The Seller had a proposed change applied; the Buyer may a paragraph cite; the Lease if a proposed change is withheld differs; the Deed but a proposed change differs."],
  ]
  expect(withoutInternalIDs(cases.map(([text]) => text).join("\n"), ids).split("\n")).toEqual(cases.map(([, expected]) => expected))
})

test("names an id left in prose by what it is, with its article, number and capital, and drops one that only labels a word naming it", () => {
  const ids = ["a1", "a2", "a3", "w1", "i1", "d4", PID, REQUIREMENT, OPEN]
  const cases = [
    [`It conflicts with ${REQUIREMENT}, which requires 60 days.`, "It conflicts with another requirement, which requires 60 days."],
    ["Reason: a1 omits it.", "Reason: A proposed change omits it."],
    ["a1-a3 changed the price", "Proposed changes changed the price"],
    ["a2 was withheld, so clause 4 is missing; a1–a3 did not change the 30-day period.", "A proposed change was withheld, so clause 4 is missing; proposed changes did not change the 30-day period."],
    [`The draft omits it. ${PID} is silent and the d4 text differs.`, "The draft omits it. A paragraph is silent and the document text differs."],
    // Ids listed with "or" name one of them; kinds are named in their order.
    [`Either w1 or i1 settles ${OPEN}.`, "Either a proposed change or an explanation settles an open item."],
    [`See cl. ${REQUIREMENT} and Acme PLC. a1 follows.`, "See cl. another requirement and Acme PLC. A proposed change follows."],
    // After a word that already names it, or alone in brackets, an id is only
    // a label and goes whole.
    [`Paragraph ${PID} and changes a1/w1 differ (see i1).`, "Paragraph and changes differ."],
    ["Changes a1–a3 were applied; w1 was withheld (i1).", "Changes were applied; a proposed change was withheld."],
  ]
  expect(withoutInternalIDs(cases.map(([text]) => text).join("\n"), ids).split("\n")).toEqual(cases.map(([, expected]) => expected))
})

const DOCUMENT = { title: "Supply note", purpose: "Record supplied terms", parts: [{ heading: "Delivery", scope: "Delivery" }] }

test("finalize keeps a requested memo's markers in the working notes; only the gate's delivery drops them", async () => {
  const directory = matter()
  writeLeanPlan(directory, [DOCUMENT])
  const gap = openItem("research", "research_gap", "No authority was found for the delivery period.", {})
  writeOpenItems(directory, [gap])
  requestMemo(directory)
  expect(requestedMemo(directory)).toBeDefined()
  expect(memoMode(directory)).toBe(true)
  const reply = `# Cover memo\n\nNo authority was found for the delivery period; confirm it before use. [ref: ${gap.id}]\n`
  expect((await acceptStageReply(directory, "finalize", reply)).issues).toEqual([])
  expect(readLeanText(directory, "drafting-notes.md")).toBe(reply)
  expect(docxText(readFileSync(leanFile(directory, "drafting-notes.docx")))).toContain(`[ref: ${gap.id}]`)
  const gate = finalGate(directory)
  expect(gate).toMatchObject({ status: "pass", blocking: [] })
  expect(deliveredNotes(directory, gate)).toBe("# Cover memo\n\nNo authority was found for the delivery period; confirm it before use.\n")

  // Without a requested memo the notes are delivered the same way.
  writeLeanPlan(directory, [DOCUMENT])
  expect(memoMode(directory)).toBe(false)
  expect(deliveredNotes(directory, finalGate(directory))).toBe("# Cover memo\n\nNo authority was found for the delivery period; confirm it before use.\n")
})

test("a requested memo delivered without findings keeps each quotation whole on a line that loses its marker", () => {
  const directory = matter()
  writeLeanPlan(directory, [DOCUMENT])
  const gap = openItem("research", "research_gap", "No authority was found for the delivery period.", {})
  writeOpenItems(directory, [gap])
  requestMemo(directory)
  writeLeanText(directory, "drafting-notes.md", ["# Cover memo", "",
    `The request says "Deliver within thirty days of the order." and no authority confirms it. [ref: ${gap.id}]`,
    `The schedule defines “Goods , Services” as "the goods  listed" . [ref: ${gap.id}]`, ""].join("\n"))
  // Was "The request says 17 and no authority confirms it.": each quotation
  // on such a line was replaced by its offset.
  expect(deliveredNotes(directory, finalGate(directory))).toBe(["# Cover memo", "",
    `The request says "Deliver within thirty days of the order." and no authority confirms it.`,
    `The schedule defines “Goods , Services” as "the goods  listed".`, ""].join("\n"))
})

const REQUEST_LOCATOR = ".dochaus/lean/questions.json answer (delivery-start)"
const PID = "p0123456789ab"
const requirement = (check: string, quote: string, locator: string): ChecklistItem =>
  ({ check, basis: "directive", source_refs: [{ document: "request", locator, quote }], qualifications: [], target: { document: "d1", role: "Supply note", treatment: "operative" } })
const ITEMS = [
  { id: "c1", item: requirement("The note states the delivery period.", "Deliver within thirty days of the order.", REQUEST_LOCATOR) },
  { id: "c2", item: requirement("The note states when delivery starts.", "Deliver within ten days of the order.", "request.md") },
]
const finding = (item_id: string, extra: Partial<ChecklistFinding> = {}): ChecklistFinding =>
  ({ item_id, status: "unsatisfied", reason: `Paragraph ${PID} omits it.`, evidence: [], source_refs: [], inspection_scope: "Current body paragraphs only.", limitation: "Comments not inspected.", ...extra })
const NOTES = "# Cover memo\n\nThe delivery period follows the order. [ref: c2]\n"

function checklistMatter() {
  const directory = matter()
  writeLeanPlan(directory, [{ ...DOCUMENT, file: "supply-note.docx" }])
  writeLeanText(directory, "d1/paragraphs.json", JSON.stringify([{ pid: PID, sha256: "sha", text: "The supplier shall deliver the goods." }]))
  writeLeanText(directory, "checklist/extraction.json", JSON.stringify({ expected_sources: ["request"], sources: [{ key: "request", status: "succeeded", items: ITEMS }] }))
  writeLeanText(directory, "checklist/checks.json", JSON.stringify({ extraction_complete: true, expected_items: ["c1", "c2"],
    checks: [finding("c1", { conflicting_source_refs: [ITEMS[1]!.item.source_refs[0]!] }), finding("c2")]
      .map(row => ({ item_id: row.item_id, input_sha: "sha", status: "succeeded", issues: [], finding: row })) }))
  writeLeanText(directory, "checklist/report.md", "# Checklist handoff findings\n")
  updateLeanStage(directory, readLeanState(directory)!, "checklist-check", { status: "done" })
  writeLeanText(directory, "drafting-notes.md", NOTES)
  return directory
}

const MEMO_FINDINGS = ["## Requirements not confirmed as met", "",
  "These requirements were not confirmed as met in this draft and need the lawyer's review.", ""]
// A memo finding's status in plain words: what is not confirmed and what the
// lawyer should check, never the workflow's own label naming its stages.
const OMISSION = "Status: not met in this draft; check whether the draft should be changed to meet it."
const CONFLICT = "Status: conflicts with another requirement in the sources; check which one governs."
const PROCESS_WORDS = ["repair", "omission", "reinspect", "source adjudication", "allowance", "the check ", "checker", "finalize", "Checklist", "Finding:"]

test("without a requested memo the findings lead in a memo's plain words, with no internal id, then the drafting notes without their markers", () => {
  const plain = checklistMatter()
  const gate = finalGate(plain)
  expect(gate).toMatchObject({ status: "pass", delivery: "review_draft_with_findings", blocking: [] })
  const delivered = deliveredNotes(plain, gate)!
  expect(delivered).toBe(["# Checklist findings not satisfied", "",
    "These checklist requirements were not established as satisfied. Each is disclosed with this review draft for review; none is treated as satisfied.", "",
    `- The note states the delivery period. ${CONFLICT} Reason: Paragraph omits it. Limitation: Comments not inspected. Source: request: "Deliver within thirty days of the order."`,
    `- The note states when delivery starts. ${OMISSION} Reason: Paragraph omits it. Limitation: Comments not inspected. Source: request: "Deliver within ten days of the order."`, "",
    "# Cover memo", "", "The delivery period follows the order.", ""].join("\n"))
  for (const text of ["c1", "c2", PID, REQUEST_LOCATOR, "[ref", "Finding:", "Conflicting"]) expect(delivered).not.toContain(text)
  expect(readLeanText(plain, "drafting-notes.md")).toBe(NOTES)
})

test("a requested memo comes first; its findings close it in the requirement's own words and sources, a conflict said without naming either side, with no internal id", () => {
  const memo = checklistMatter()
  requestMemo(memo)
  const gate = finalGate(memo)
  expect(gate).toMatchObject({ status: "pass", delivery: "review_draft_with_findings", blocking: [] })
  const delivered = deliveredNotes(memo, gate)!
  expect(delivered).toBe(["# Cover memo", "", "The delivery period follows the order.", "", ...MEMO_FINDINGS,
    `- The note states the delivery period. ${CONFLICT} Reason: Paragraph omits it. Limitation: Comments not inspected. Source: request: "Deliver within thirty days of the order."`,
    `- The note states when delivery starts. ${OMISSION} Reason: Paragraph omits it. Limitation: Comments not inspected. Source: request: "Deliver within ten days of the order."`, ""].join("\n"))
  expect(delivered.split("\n")[0]).toBe("# Cover memo")
  // Every finding stays disclosed; none is worded from the checker's detail.
  expect(gate.items.filter(item => item.kind === "checklist_unresolved")).toHaveLength(2)
  for (const text of ["c1", "c2", PID, REQUEST_LOCATOR, "in the instructions", "[ref", ...PROCESS_WORDS]) expect(delivered.split(MEMO_FINDINGS[0]!)[1]).not.toContain(text)
  expect(readLeanText(memo, "drafting-notes.md")).toBe(NOTES)
})

const MEMO_LEFT_OPEN = ["## Requirements left open", "", "These requirements were left open in this draft and need the lawyer's review.", ""]
const LEFT_OPEN = "Status: left open in this draft; settle what it leaves open and check the draft against it."
const saveFindings = (directory: string, rows: ChecklistFinding[]) => writeLeanText(directory, "checklist/checks.json", JSON.stringify({ extraction_complete: true,
  expected_items: rows.map(row => row.item_id), checks: rows.map(row => ({ item_id: row.item_id, input_sha: "sha", status: "succeeded", issues: [], finding: row })) }))

test("a requested memo's requirements left open close it, after any unresolved finding, in the requirement's own words with no internal id", () => {
  const memo = checklistMatter()
  requestMemo(memo)
  const open = finding("c2", { status: "open", open_decisions: ["Whether delivery starts at dispatch"] })
  saveFindings(memo, [finding("c1"), open])
  const gate = finalGate(memo)
  expect(gate).toMatchObject({ status: "pass", delivery: "review_draft_with_findings", blocking: [] })
  expect(gate.disclosed.map(item => [item.kind, item.ref.checklist_item])).toEqual([["checklist_unresolved", "c1"], ["checklist_open", "c2"]])
  const line = `- The note states when delivery starts. ${LEFT_OPEN} Reason: Paragraph omits it. Limitation: Comments not inspected. Open decisions: Whether delivery starts at dispatch. Source: request: "Deliver within ten days of the order."`
  const delivered = deliveredNotes(memo, gate)!
  expect(delivered).toBe(["# Cover memo", "", "The delivery period follows the order.", "", ...MEMO_FINDINGS,
    `- The note states the delivery period. ${OMISSION} Reason: Paragraph omits it. Limitation: Comments not inspected. Source: request: "Deliver within thirty days of the order."`, "",
    ...MEMO_LEFT_OPEN, line, ""].join("\n"))
  for (const text of ["c1", "c2", PID, REQUEST_LOCATOR, "[ref", ...PROCESS_WORDS]) expect(delivered.split(MEMO_FINDINGS[0]!)[1]).not.toContain(text)
  // With nothing unresolved, the requirements left open close the memo alone.
  saveFindings(memo, [finding("c1", { status: "satisfied" }), open])
  const alone = deliveredNotes(memo, finalGate(memo))!
  expect(alone).toBe(["# Cover memo", "", "The delivery period follows the order.", "", ...MEMO_LEFT_OPEN, line, ""].join("\n"))
  // An open item the checklist output no longer holds is read back from its detail.
  writeLeanText(memo, "checklist/extraction.json", JSON.stringify({ expected_sources: [], sources: [] }))
  expect(deliveredNotes(memo, gate)!.split("\n").at(-2)).toBe(line.replace(/ Source: .*$/, ""))
})

test("a memo finding whose checklist item cannot be found takes its requirement from its detail, cut of its conflict list and ids, with its plain status", () => {
  const memo = checklistMatter()
  requestMemo(memo)
  const gate = finalGate(memo)
  const [c1, c2] = ["c1", "c2"].map(id => gate.items.find(item => item.kind === "checklist_unresolved" && item.ref.checklist_item === id)!.detail)
  expect(c1).toEndWith(`Paragraph ${PID} omits it. Limitation: Comments not inspected. Conflicting checklist items: c2.`)
  // The checklist output read for the notes no longer holds the items.
  writeLeanText(memo, "checklist/extraction.json", JSON.stringify({ expected_sources: [], sources: [] }))
  const delivered = deliveredNotes(memo, gate)!
  expect(delivered).toBe(["# Cover memo", "", "The delivery period follows the order.", "", ...MEMO_FINDINGS,
    // c1's status already names the conflict, so no clause repeats it.
    `- The note states the delivery period. ${CONFLICT} Reason: Paragraph omits it. Limitation: Comments not inspected.`,
    `- The note states when delivery starts. ${OMISSION} Reason: Paragraph omits it. Limitation: Comments not inspected.`, ""].join("\n"))
  for (const text of [PID, "c2", REQUEST_LOCATOR, "Conflicting", ...PROCESS_WORDS]) expect(delivered.split(MEMO_FINDINGS[0]!)[1]).not.toContain(text)
  // A detail of another shape is still delivered as written, less its ids.
  const other = { ...gate.items.find(item => item.ref.checklist_item === "c2")!, detail: `Paragraph ${PID} needs the lawyer's review.` }
  expect(deliveredNotes(memo, { ...gate, items: [other] })!.split("\n").at(-2)).toBe("- Paragraph needs the lawyer's review.")
})

test("a memo finding keeps its source quote whole and names the requirement and paragraph its reason cites in plain words", () => {
  const directory = checklistMatter()
  const [own, other] = ["source-0123456789ab-1", REQUIREMENT]
  const quote = `Deliver the d1 goods per ${PID} within ten days (see ${other}).`
  const items = [{ id: own, item: requirement("The note states the delivery period.", quote, "request.md") }, { id: other, item: requirement("The note states when delivery starts.", "Deliver within sixty days.", "request.md") }]
  writeLeanText(directory, "checklist/extraction.json", JSON.stringify({ expected_sources: ["request"], sources: [{ key: "request", status: "succeeded", items }] }))
  writeLeanText(directory, "checklist/checks.json", JSON.stringify({ extraction_complete: true, expected_items: [own, other],
    checks: [finding(own, { reason: `It conflicts with ${other}, which requires 60 days; ${PID} omits it.` }), finding(other, { status: "satisfied" })]
      .map(row => ({ item_id: row.item_id, input_sha: "sha", status: "succeeded", issues: [], finding: row })) }))
  requestMemo(directory)
  expect(deliveredNotes(directory, finalGate(directory))!.split("\n").at(-2))
    .toBe(`- The note states the delivery period. ${OMISSION} Reason: It conflicts with another requirement, which requires 60 days; a paragraph omits it. Limitation: Comments not inspected. Source: request: "${quote}"`)
})

test("a memo finding's status says, in plain words for each class, what is not confirmed and what to check, never the stage that left it", () => {
  const memo = checklistMatter()
  requestMemo(memo)
  const gate = finalGate(memo)
  const c2 = gate.items.find(item => item.kind === "checklist_unresolved" && item.ref.checklist_item === "c2")!
  const rest = " The note states when delivery starts. Finding: Paragraph omits it. Limitation: Comments not inspected."
  const cases = [
    ["omission", "omission not fixed by the repair round", OMISSION],
    ["omission", "omission; source adjudication declined the repair", OMISSION],
    ["omission", "omission; the repair allowance was used up", OMISSION],
    ["missing_inspection", "not established by the check after reinspection", "Status: could not be confirmed from the draft; check the draft against it."],
    ["missing_inspection", "not established by the check; not yet reinspected", "Status: could not be confirmed from the draft; check the draft against it."],
    ["not_inspectable", "cannot be established from document text", "Status: cannot be confirmed from the document text; check it directly."],
    ["departure", "deliberate departure", "Status: the draft departs from it on purpose; check that the departure is intended."],
  ] as const
  for (const [kind, label, status] of cases) {
    const finding = { ...c2, detail: `Checklist requirement unresolved (${label}):${rest}`, ref: { ...c2.ref, checklist_class: kind } }
    const line = deliveredNotes(memo, { ...gate, items: [finding] })!.split("\n").at(-2)!
    expect(line).toBe(`- The note states when delivery starts. ${status} Reason: Paragraph omits it. Limitation: Comments not inspected. Source: request: "Deliver within ten days of the order."`)
    for (const text of PROCESS_WORDS) expect(line).not.toContain(text)
  }
  // A refined parent carries its reconciled coverage instead of a class, and
  // its coverage issues as its reason.
  for (const [coverage, status] of [["not satisfied", OMISSION], ["could not be established", "Status: could not be confirmed from the draft; check the draft against it."]]) {
    const { checklist_class: _, ...ref } = c2.ref
    const parent = { ...c2, ref, detail: `Checklist requirement unresolved across its refined checks (${coverage}): The note states when delivery starts. Coverage: A refined check remains adverse.` }
    expect(deliveredNotes(memo, { ...gate, items: [parent] })!.split("\n").at(-2))
      .toBe(`- The note states when delivery starts. ${status} Reason: A refined check remains adverse. Source: request: "Deliver within ten days of the order."`)
  }
  // Without a requested memo the notes give the same plain status.
  writeLeanPlan(memo, [{ ...DOCUMENT, file: "supply-note.docx" }])
  expect(deliveredNotes(memo, { ...gate, items: [{ ...c2, detail: `Checklist requirement unresolved (omission not fixed by the repair round):${rest}` }] })!.split("\n"))
    .toContain(`- The note states when delivery starts. ${OMISSION} Reason: Paragraph omits it. Limitation: Comments not inspected. Source: request: "Deliver within ten days of the order."`)
})

// A refined parent's coverage issues in the memo: one plain reason for each
// kind of issue, saying only what that issue records, without the checker's
// wording or the ids of the checks it names.
const MISSING = "Status: could not be confirmed from the draft; check the draft against it."
const NO_FINDING = "A check of part of it did not complete."
const INSPECT = "Whether part of it was carried out, or applies, could not be confirmed from the draft."
const ADOPTION = "No current text or file in the draft shows that part of it was adopted."
const DEFERRED = "Part of it depends on a document-wide property that was not decided."
const CHECKER_WORDS = ["fulfillment", "applicability", "requires inspection", "operative", "artifact", "deferred", "document-properties", "finding", "Coverage", "-refined-", "-comparison-"]

test("a refined parent's memo finding keeps its coverage reasons, whether or not its checklist item is found", () => {
  const memo = checklistMatter()
  requestMemo(memo)
  const gate = finalGate(memo)
  const c2 = gate.items.find(item => item.kind === "checklist_unresolved" && item.ref.checklist_item === "c2")!
  const { checklist_class: _, ...ref } = c2.ref
  const parent = { ...c2, ref, detail: `Checklist requirement unresolved across its refined checks (could not be established): The note states when delivery starts. Coverage: c2: no completed finding. The delivery-start check needs inspection.` }
  // Wording the checker does not write stays as written.
  const reason = `Reason: ${NO_FINDING} The delivery-start check needs inspection.`
  const line = () => deliveredNotes(memo, { ...gate, items: [parent] })!.split("\n").at(-2)
  const found = line()
  // The checklist output read for the notes no longer holds the item: the
  // requirement comes from the detail, and its coverage reasons stay.
  writeLeanText(memo, "checklist/extraction.json", JSON.stringify({ expected_sources: [], sources: [] }))
  expect([found, line()]).toEqual([
    `- The note states when delivery starts. ${MISSING} ${reason} Source: request: "Deliver within ten days of the order."`,
    `- The note states when delivery starts. ${MISSING} ${reason}`,
  ])
})

test("a refined parent's coverage read back from its detail gives each kind of issue once, in plain words, with no ids", () => {
  const memo = checklistMatter()
  requestMemo(memo)
  const gate = finalGate(memo)
  const c2 = gate.items.find(item => item.kind === "checklist_unresolved" && item.ref.checklist_item === "c2")!
  const { checklist_class: _, ...ref } = c2.ref
  const [k1, k2, k3] = ["c2-refined-0123456789ab-1", "c2-comparison-0123456789ab", "c2-comparison-ba9876543210"]
  const issues = [`${k1}: action fulfillment or applicability requires inspection.`, `${k2}: adoption lacks current operative or artifact evidence.`,
    `${k3}: action fulfillment or applicability requires inspection.`, `c2: ${k2}, ${k3} deferred a document-wide property, but no document-properties finding decided it.`]
  const parent = { ...c2, ref, detail: `Checklist requirement unresolved across its refined checks (could not be established): The note states when delivery starts. Coverage: ${issues.join(" ")}` }
  const line = deliveredNotes(memo, { ...gate, items: [parent] })!.split("\n").at(-2)!
  expect(line).toBe(`- The note states when delivery starts. ${MISSING} Reason: ${INSPECT} ${ADOPTION} ${DEFERRED} Source: request: "Deliver within ten days of the order."`)
  for (const text of CHECKER_WORDS) expect(line).not.toContain(text)
})

test("a refined parent's finding gives its recorded coverage issues in plain words, with or without a memo; its disclosure keeps the checker's", () => {
  const directory = matter()
  writeLeanPlan(directory, [{ ...DOCUMENT, file: "supply-note.docx" }])
  writeLeanText(directory, "d1/paragraphs.json", JSON.stringify([{ pid: PID, sha256: "sha", text: "The supplier shall deliver the goods." }]))
  writeLeanText(directory, "checklist/extraction.json", JSON.stringify({ expected_sources: ["request"], sources: [{ key: "request", status: "succeeded", items: ITEMS }] }))
  // c2 is refined into four checks: one adopts without current text, two
  // need inspection and one defers a document-wide property that no check
  // decided.
  const quote = ITEMS[1]!.item.source_refs[0]!
  const children = ([["The note adopts the start date.", "adopt"], ["The note adopts the start trigger.", "adopt"], ["The note adopts the start notice.", "adopt"], ["The note's start wording is consistent.", "assess"]] as const)
    .map(([check, required_action], index) => ({ id: `c2-refined-0123456789ab-${index + 1}`, item: { ...requirement(check, quote.quote, quote.locator), required_action } }))
  const unsettled = { status: "indeterminate" as const, action_status: "indeterminate" as const, reason: "The start trigger was not inspected." }
  writeLeanText(directory, "checklist/checks.json", JSON.stringify({ extraction_complete: true, expected_items: ["c1", ...children.map(row => row.id)],
    checks: [finding("c1", { status: "satisfied" }), finding(children[0]!.id, { status: "satisfied", action_status: "fulfilled" }), finding(children[1]!.id, unsettled),
      finding(children[2]!.id, unsettled), finding(children[3]!.id, { status: "satisfied", action_status: "fulfilled", deferred_to: "document_properties" })]
      .map(row => ({ item_id: row.item_id, input_sha: "sha", status: "succeeded", issues: [], finding: row })) }))
  writeLeanText(directory, "checklist/report.md", "# Checklist handoff findings\n")
  updateLeanStage(directory, readLeanState(directory)!, "checklist-check", { status: "done" })
  writeLeanText(directory, "drafting-notes.md", NOTES)
  writeLeanText(directory, "checklist/refinements.json", JSON.stringify({ parents: [{ parent_id: "c2", parent: ITEMS[1]!.item, scope: "cited_sources",
    source_basis: checklistRefinementSourceBasis(directory), expected_units: ["u1"], units: [{ unit_id: "u1", input_sha: "sha", status: "succeeded",
      result: { parent_id: "c2", unit_id: "u1", disposition: "refined", reason: "Focused duties.", items: children.map(row => row.item), source_refs: [quote], limitations: [] },
      children, issues: [] }] }] }))
  const gate = finalGate(directory)
  expect(gate).toMatchObject({ status: "pass", delivery: "review_draft_with_findings", blocking: [] })
  const parent = gate.items.find(item => item.kind === "checklist_unresolved" && item.ref.checklist_item === "c2")!
  const [k1, k2, k3, k4] = children.map(row => row.id)
  // The gate's disclosure keeps the checker's own issues, ids and all.
  expect(parent.detail).toBe(`Checklist requirement unresolved across its refined checks (could not be established): The note states when delivery starts. Coverage: ${k1}: adoption lacks current operative or artifact evidence. ${k2}: action fulfillment or applicability requires inspection. ${k3}: action fulfillment or applicability requires inspection. c2: ${k4} deferred a document-wide property, but no document-properties finding decided it.`)
  const plain = deliveredNotes(directory, gate)!
  requestMemo(directory)
  const line = deliveredNotes(directory, gate)!.split("\n").find(row => row.startsWith("- The note states when delivery starts."))!
  expect(line).toBe(`- The note states when delivery starts. ${MISSING} Reason: ${ADOPTION} ${INSPECT} ${DEFERRED} Source: request: "Deliver within ten days of the order."`)
  for (const text of CHECKER_WORDS) expect(line).not.toContain(text)
  // Without a requested memo the notes give it in the same plain words.
  expect(plain.split("\n")).toContain(line)
})

test("a memo finding's reason is its recorded finding, even where the requirement itself names a field", () => {
  const memo = checklistMatter()
  const items = [ITEMS[0]!, { id: "c2", item: { ...ITEMS[1]!.item, check: "The schedule states Coverage: 5 million per claim." } }]
  writeLeanText(memo, "checklist/extraction.json", JSON.stringify({ expected_sources: ["request"], sources: [{ key: "request", status: "succeeded", items }] }))
  requestMemo(memo)
  const gate = finalGate(memo)
  const line = () => deliveredNotes(memo, gate)!.split("\n").at(-2)
  const found = line()
  writeLeanText(memo, "checklist/extraction.json", JSON.stringify({ expected_sources: [], sources: [] }))
  expect([found, line()]).toEqual([
    `- The schedule states Coverage: 5 million per claim. ${OMISSION} Reason: Paragraph omits it. Limitation: Comments not inspected. Source: request: "Deliver within ten days of the order."`,
    `- The schedule states Coverage: 5 million per claim. ${OMISSION} Reason: Paragraph omits it. Limitation: Comments not inspected.`,
  ])
})

test("a memo finding names each planned document by its title wherever the requirement put its id, and names a parenthesised target unless it is the memo", () => {
  const directory = matter()
  writeLeanPlan(directory, [{ ...DOCUMENT, file: "supply-pack.docx" }, { title: "Supply agreement", purpose: "Full terms", parts: [{ heading: "Price", scope: "Price" }] }])
  requestMemo(directory)
  const items = [
    { id: "c1", item: { ...requirement("For d1, across the pack containing the cover memo (notes) and the agreement (d1), is d1 delivered in the single Word file named supply-pack.docx?", "Deliver everything as one file named supply-pack.docx.", "request.md"), target: { document: "d1", role: "Supply note", treatment: "operative" as const } } },
    { id: "c2", item: requirement("d1 and d2 state the same delivery period; the cover memo comes first, followed by d1 (scope: memo plus d1, d2).", "Deliver within thirty days of the order.", "request.md") },
  ]
  writeLeanText(directory, "d1/paragraphs.json", JSON.stringify([{ pid: PID, sha256: "sha", text: "The supplier shall deliver the goods." }]))
  writeLeanText(directory, "checklist/extraction.json", JSON.stringify({ expected_sources: ["request"], sources: [{ key: "request", status: "succeeded", items }] }))
  writeLeanText(directory, "checklist/checks.json", JSON.stringify({ extraction_complete: true, expected_items: ["c1", "c2"],
    checks: [finding("c1", { reason: "d1 body names no file. However the inspected d1 properties name none, as d2's do not." }), finding("c2", { reason: `Paragraph ${PID} of d2 and the d1 text differ.` })]
      .map(row => ({ item_id: row.item_id, input_sha: "sha", status: "succeeded", issues: [], finding: row })) }))
  writeLeanText(directory, "checklist/report.md", "# Checklist handoff findings\n")
  updateLeanStage(directory, readLeanState(directory)!, "checklist-check", { status: "done" })
  writeLeanText(directory, "drafting-notes.md", NOTES)
  const gate = finalGate(directory)
  expect(gate).toMatchObject({ status: "pass", delivery: "review_draft_with_findings", blocking: [] })
  expect(deliveredNotes(directory, gate)!.split(MEMO_FINDINGS[0]!)[1]!.split("\n").filter(line => line.startsWith("- "))).toEqual([
    `- For the Supply note, across the pack containing the cover memo and the agreement (the Supply note), is the Supply note delivered in the single Word file named supply-pack.docx? ${OMISSION} Reason: The Supply note body names no file. However the inspected Supply note properties name none, as the Supply agreement's do not. Limitation: Comments not inspected. Source: request: "Deliver everything as one file named supply-pack.docx."`,
    `- The Supply note and the Supply agreement state the same delivery period; the cover memo comes first, followed by the Supply note (scope: memo plus the Supply note, the Supply agreement). ${OMISSION} Reason: Paragraph of the Supply agreement and the Supply note text differ. Limitation: Comments not inspected. Source: request: "Deliver within thirty days of the order."`,
  ])
})

/** The memo finding line delivered for each requirement text, in a matter
 * planning the Supply note, the Supply agreement and the Deed of Variation, in
 * that order. Its checklist item is not found, so it takes its requirement from
 * its detail. */
function titledLine(ref: Partial<OpenItem["ref"]> = {}) {
  const directory = checklistMatter()
  requestMemo(directory)
  const gate = finalGate(directory)
  const c2 = gate.items.find(item => item.kind === "checklist_unresolved" && item.ref.checklist_item === "c2")!
  writeLeanPlan(directory, [DOCUMENT, { title: "Supply agreement", purpose: "Full terms", parts: [{ heading: "Price", scope: "Price" }] },
    { title: "The Deed of Variation", purpose: "Varies the terms", parts: [{ heading: "Variation", scope: "Variation" }] }])
  requestMemo(directory)
  const line = (text: string) => deliveredNotes(directory, { ...gate, items: [{ ...c2, detail: `Checklist requirement unresolved (omission): ${text}`,
    ref: { ...c2.ref, checklist_item: "c9", checklist_class: "omission" as const, ...ref } }] })!.split("\n").at(-2)
  return (cases: string[][]) => expect(cases.map(([text]) => line(text!))).toEqual(cases.map(([, expected]) => `- ${expected} ${OMISSION}`))
}

test("a memo finding names documents whose ids are joined by a slash or dash, keeps a parenthesised target, and reads plainly around a title", () => {
  titledLine()([
    // Ids joined by "/", or by an en dash from a later document to an earlier
    // one, each name their document; a hyphen joins just the two; a longer id
    // or a word is left as written.
    ["d1/d2 state the same delivery period.", "The Supply note/the Supply agreement state the same delivery period."],
    ["Cross-references between d2-d1 and d3–d2 agree.", "Cross-references between the Supply agreement and the Supply note and the Deed of Variation–the Supply agreement agree."],
    ["Rows d12 and d1x stay.", "Rows d12 and d1x stay."],
    // A parenthesis names its targets, but not the memo alone or the document
    // whose title it follows.
    ["The notice period is stated (d2).", "The notice period is stated (the Supply agreement)."],
    ["The memo (notes) and the Supply note (d1) agree, as do both terms (d1/d3).", "The memo and the Supply note agree, as do both terms (the Supply note/the Deed of Variation)."],
    ["The price is set (d2 and notes).", "The price is set (the Supply agreement and this memo)."],
    // A title's own article is not doubled or capitalized mid-sentence.
    ["For d3, check the d3 schedule.", "For the Deed of Variation, check the Deed of Variation schedule."],
    // An abbreviation's period does not start a sentence.
    ["d3 is signed, e.g. d2 is not.", "The Deed of Variation is signed, e.g. the Supply agreement is not."],
    // An id before a capitalized name labels that name: it follows it as its
    // title, or goes where the name already gives the title.
    ["The d1 Lease and the d3 Deed of Variation are signed.", "The Lease (the Supply note) and the Deed of Variation are signed."],
  ])
})

test("a memo finding names a range of planned ids as every planned document from the first to the last, in plan order", () => {
  titledLine()([
    // Was "The Supply note–the Deed of Variation state …": only the ends.
    ["d1–d3 state the same delivery period.", "The Supply note, the Supply agreement and the Deed of Variation state the same delivery period."],
    // The first title takes the article a single id would.
    ["Check the d1-d2 cross-references.", "Check the Supply note and the Supply agreement cross-references."],
    ["The terms agree (d1–d3).", "The terms agree (the Supply note, the Supply agreement and the Deed of Variation)."],
    // Reversed or equal ends, an end that is no planned id, a slash and a longer
    // run of joined ids are each named, as before.
    ["d3–d1 and d2-d2 differ.", "The Deed of Variation–the Supply note and the Supply agreement-the Supply agreement differ."],
    ["d1-d4 and d1/d3 differ.", "The Supply note-d4 and the Supply note/the Deed of Variation differ."],
    ["d1-d2-d3 differ.", "The Supply note-the Supply agreement-the Deed of Variation differ."],
  ])
})

test("a parenthesis after a document's title drops only that document's id and names its other targets as joined to it", () => {
  titledLine()([
    // Was "The Supply Agreement (the Supply note) requires notice.", read as
    // another name for it; before that, the Supply note was lost.
    ["The Supply Agreement (d1/d2) requires notice.", "The Supply Agreement (and the Supply note) requires notice."],
    // Was "The Supply note (this memo) records it."; before that, the memo was lost.
    ["The Supply note (d1, notes) records it.", "The Supply note (and this memo) records it."],
    ["The Supply note (d1) records it; the Deed of Variation (d1, d2 and d3) varies it.", "The Supply note records it; the Deed of Variation (and the Supply note and the Supply agreement) varies it."],
    ["The Supply agreement (d2/notes) is signed; the Supply note (d1 or d3) is not.", "The Supply agreement (and this memo) is signed; the Supply note (or the Deed of Variation) is not."],
  ])
})

test("a memo finding keeps a title's article where the id opens its own clause, takes a determiner's modifier, and capitalizes after no abbreviation", () => {
  titledLine()([
    // After a preposition and a determiner, a noun that opens a clause is not
    // the id's. Was "In the event Supply note is signed" and "to the extent
    // Supply agreement is silent".
    ["In the event d1 is signed, d2 follows.", "In the event the Supply note is signed, the Supply agreement follows."],
    ["The price applies to the extent d2 is silent.", "The price applies to the extent the Supply agreement is silent."],
    ["In each case d1 is signed; the executed d1 is filed.", "In each case the Supply note is signed; the executed Supply note is filed."],
    // Any other modifier is the id's.
    ["Under the executed d1, clause 3 differs; in the executed d1 the price is fixed.", "Under the executed Supply note, clause 3 differs; in the executed Supply note the price is fixed."],
    // Any determiner takes one modifier. Was "Each delivered the Supply note".
    ["Each delivered d1 is filed; their final d2 and its signed d3 differ.", "Each delivered Supply note is filed; their final Supply agreement and its signed Deed of Variation differ."],
    // A word with no vowel before a period is an abbreviation. Was "See cl. The
    // Supply note and Sch. The Supply agreement".
    ["See cl. d1 and Sch. d2; the term is met. d3 is signed.", "See cl. the Supply note and Sch. the Supply agreement; the term is met. The Deed of Variation is signed."],
  ])
})

test("a memo finding names two planned ids joined by a hyphen as just those two; only an en dash runs through the documents between", () => {
  titledLine()([
    // Was "Check the Supply note, the Supply agreement and the Deed of
    // Variation cross-references.": the hyphen added the document between.
    ["Check the d1-d3 cross-references.", "Check the Supply note and the Deed of Variation cross-references."],
    ["d3-d1 differ; d1–d3 agree.", "The Deed of Variation and the Supply note differ; the Supply note, the Supply agreement and the Deed of Variation agree."],
  ])
})

test("a memo finding names the memo joined to a planned id with that document's title, and an id joined to a word by its title", () => {
  titledLine()([
    // Was "(d2/this memo)", "(this memo/d2)" and "(the Supply agreement–this
    // memo)": the memo was named before the id, which then stayed raw.
    ["The price is set (d2/notes).", "The price is set (the Supply agreement and this memo)."],
    ["The price is set (notes/d2).", "The price is set (this memo and the Supply agreement)."],
    ["The price is set (d2–notes).", "The price is set (the Supply agreement and this memo)."],
    // Were left as "d1/notes", "notes-d2", "d1-specific" and "d2-only".
    ["d1/notes agree on price; notes-d2 differ.", "The Supply note and this memo agree on price; this memo and the Supply agreement differ."],
    ["d1-specific and the d2-only terms differ.", "Supply note-specific and the Supply agreement-only terms differ."],
  ])
})

test("a memo finding capitalizes after an acronym that ends a sentence; only a lowercase or short mixed-case word with no vowel is an abbreviation", () => {
  titledLine()([
    // Was "the GDPR. the Supply note records", and likewise after each acronym.
    ["The data terms follow the GDPR. d1 records the processor.", "The data terms follow the GDPR. The Supply note records the processor."],
    ["Tax is paid to HMRC. d2 records it; the adviser is Smith LLP. d3 appoints it.", "Tax is paid to HMRC. The Supply agreement records it; the adviser is Smith LLP. The Deed of Variation appoints it."],
    ["The buyer is Acme PLC. d2 names it.", "The buyer is Acme PLC. The Supply agreement names it."],
    ["The claim is A v. d1 terms; see cl. d2 and Sch. d3.", "The claim is A v. the Supply note terms; see cl. the Supply agreement and Sch. the Deed of Variation."],
  ])
})

test("a memo finding gives a title its article where the word after the determiner heads a phrase of its own, and otherwise takes that word as the id's", () => {
  titledLine()([
    // Was "Under the executed the Supply note is" and "in the current the
    // Supply agreement is": an auxiliary after the id dropped the modifier.
    ["Under the executed d1 is a schedule; in the current d2 is no notice clause.", "Under the executed Supply note is a schedule; in the current Supply agreement is no notice clause."],
    // Was "In the event Supply note and" and "To the extent Supply agreement's".
    ["In the event d1 and d2 are signed, d3 follows.", "In the event the Supply note and the Supply agreement are signed, the Deed of Variation follows."],
    ["To the extent d2's terms are silent, d1 applies.", "To the extent the Supply agreement's terms are silent, the Supply note applies."],
    // A plural noun heads its own phrase. Was "The documents Supply note and".
    ["The documents d1 and d2 must agree on price.", "The documents the Supply note and the Supply agreement must agree on price."],
    ["The previous d1 and the various d2 differ.", "The previous Supply note and the various Supply agreement differ."],
  ])
})

test("a memo finding names an id left just after a document's title as that document's, and drops one whose kind the next word names", () => {
  titledLine({ amendment_id: "a1" })([
    // Was "The Supply note a paragraph omits it; the Supply agreement a
    // proposed change amendment differs.".
    [`d1 ${PID} omits it; the d2 a1 amendment differs.`, "The Supply note paragraph omits it; the Supply agreement amendment differs."],
    [`Under d3 ${PID}, d2 a1 is silent.`, "Under the Deed of Variation paragraph, the Supply agreement proposed change is silent."],
    // A title followed by a verb or a joining word leaves the id its article.
    [`d3 omits a1; d1 and ${PID} differ.`, "The Deed of Variation omits a proposed change; the Supply note and a paragraph differ."],
  ])
})

const PLANNED = [DOCUMENT, { title: "Supply agreement", purpose: "Full terms", parts: [{ heading: "Price", scope: "Price" }] },
  { title: "The Deed of Variation", purpose: "Varies the terms", parts: [{ heading: "Variation", scope: "Variation" }] }]

/** The memo finding line delivered for one checked requirement, its source
 * quote and its recorded reason, in a matter planning `documents` whose gate
 * also records the withheld changes a1 and w1. */
function checkedLine(check: string, quote: string, reason: string, documents = PLANNED) {
  const directory = matter()
  writeLeanPlan(directory, documents)
  requestMemo(directory)
  writeLeanText(directory, "d1/paragraphs.json", JSON.stringify([{ pid: PID, sha256: "sha", text: "The supplier shall deliver the goods." }]))
  writeLeanText(directory, "checklist/extraction.json", JSON.stringify({ expected_sources: ["request"], sources: [{ key: "request", status: "succeeded", items: [{ id: "c1", item: requirement(check, quote, "request.md") }] }] }))
  writeLeanText(directory, "checklist/checks.json", JSON.stringify({ extraction_complete: true, expected_items: ["c1"],
    checks: [{ item_id: "c1", input_sha: "sha", status: "succeeded", issues: [], finding: finding("c1", { reason }) }] }))
  writeLeanText(directory, "checklist/report.md", "# Checklist handoff findings\n")
  updateLeanStage(directory, readLeanState(directory)!, "checklist-check", { status: "done" })
  writeLeanText(directory, "drafting-notes.md", NOTES)
  const gate = finalGate(directory)
  const withheld = ["a1", "w1"].map(id => openItem("apply", "unapplied_amendment", `Change ${id} was withheld.`, { amendment_id: id }))
  return deliveredNotes(directory, { ...gate, items: [...gate.items, ...withheld] })!.split("\n").at(-2)
}

test("a memo finding's source quote reaches the reader verbatim, though it holds quotation marks and ids of its own", () => {
  // Was "…takes "Unit the Supply agreement" and "Lot a proposed change, the
  // Supply note" as defined.": its inner marks re-paired, exposing its words.
  const quote = `The tenant takes "Unit d2" and "Lot a1 , d1" as defined.`
  expect(checkedLine("d1 states the delivery period.", quote, `Paragraph ${PID} omits it.`))
    .toBe(`- The Supply note states the delivery period. ${OMISSION} Reason: Paragraph omits it. Limitation: Comments not inspected. Source: request: "${quote}"`)
})

test("a memo finding names a planned id joined to another run id by its title, and that id by its kind, the two joined by \"and\"", () => {
  // Was "The Supply note/a1", "the Supply agreement-w1" and "a1/the Deed of
  // Variation": the title took the joint and left the other id standing raw.
  expect(checkedLine("d1/a1 state the delivery period.", "Deliver within ten days.", `d2-w1 differ; so do a1/d3 and d1/${PID}.`))
    .toBe(`- The Supply note and a proposed change state the delivery period. ${OMISSION} Reason: The Supply agreement and a proposed change differ; so do a proposed change and the Deed of Variation and the Supply note and a paragraph. Limitation: Comments not inspected. Source: request: "Deliver within ten days."`)
})

test("a memo finding names a planned id quoted alone by its title, straight or curly, but keeps a source quote holding only that id", () => {
  // Was "omits a document and “d3”, and documents differ" and the source "a document".
  expect(checkedLine("The note states the delivery period.", "d2", `The draft omits "d2" and “d3”, and "d1/d3" differ.`))
    .toBe(`- The note states the delivery period. ${OMISSION} Reason: The draft omits the Supply agreement and the Deed of Variation, and the Supply note/the Deed of Variation differ. Limitation: Comments not inspected. Source: request: "d2"`)
})

test("a memo finding keeps a document's title as written: a quotation mark in it pairs with nothing, and an id-shaped word in it stays", () => {
  const [note, agreement, deed] = PLANNED
  // Was "Reason: Paragraph p0123456789ab omits it.": the title's mark paired
  // with the source quote's and shielded the paragraph id.
  expect(checkedLine("d1 states the delivery period.", "Deliver within ten days.", `Paragraph ${PID} omits it.`, [{ ...note!, title: `Spec for 3" pipe` }, agreement!, deed!]))
    .toBe(`- The Spec for 3" pipe states the delivery period. ${OMISSION} Reason: Paragraph omits it. Limitation: Comments not inspected. Source: request: "Deliver within ten days."`)
  // Was "The Schedule a proposed change plans".
  expect(checkedLine("d2 states the delivery period.", "Deliver within ten days.", `Paragraph ${PID} omits it.`, [note!, { ...agreement!, title: "Schedule a1 plans" }, deed!]))
    .toBe(`- The Schedule a1 plans states the delivery period. ${OMISSION} Reason: Paragraph omits it. Limitation: Comments not inspected. Source: request: "Deliver within ten days."`)
})

test("a refined parent's memo finding with no coverage issue recorded gives a plain reason, recorded or read back", () => {
  const directory = matter()
  writeLeanPlan(directory, [{ ...DOCUMENT, file: "supply-note.docx" }])
  writeLeanText(directory, "d1/paragraphs.json", JSON.stringify([{ pid: PID, sha256: "sha", text: "The supplier shall deliver the goods." }]))
  writeLeanText(directory, "checklist/extraction.json", JSON.stringify({ expected_sources: ["request"], sources: [{ key: "request", status: "succeeded", items: ITEMS }] }))
  // c2 is refined into two checks; one is not met, and coverage records no issue.
  const quote = ITEMS[1]!.item.source_refs[0]!
  const children = ["The note states the start date.", "The note states the start trigger."]
    .map((check, index) => ({ id: `c2-refined-0123456789ab-${index + 1}`, item: { ...requirement(check, quote.quote, quote.locator), required_action: "assess" as const } }))
  writeLeanText(directory, "checklist/checks.json", JSON.stringify({ extraction_complete: true, expected_items: ["c1", ...children.map(row => row.id)],
    checks: [finding("c1", { status: "satisfied" }), finding(children[0]!.id, { status: "satisfied", action_status: "fulfilled" }), finding(children[1]!.id, { action_status: "not_fulfilled" })]
      .map(row => ({ item_id: row.item_id, input_sha: "sha", status: "succeeded", issues: [], finding: row })) }))
  writeLeanText(directory, "checklist/report.md", "# Checklist handoff findings\n")
  updateLeanStage(directory, readLeanState(directory)!, "checklist-check", { status: "done" })
  writeLeanText(directory, "drafting-notes.md", NOTES)
  writeLeanText(directory, "checklist/refinements.json", JSON.stringify({ parents: [{ parent_id: "c2", parent: ITEMS[1]!.item, scope: "cited_sources",
    source_basis: checklistRefinementSourceBasis(directory), expected_units: ["u1"], units: [{ unit_id: "u1", input_sha: "sha", status: "succeeded",
      result: { parent_id: "c2", unit_id: "u1", disposition: "refined", reason: "Focused duties.", items: children.map(row => row.item), source_refs: [quote], limitations: [] },
      children, issues: [] }] }] }))
  const gate = finalGate(directory)
  const parent = gate.items.find(item => item.kind === "checklist_unresolved" && item.ref.checklist_item === "c2")!
  // The gate's disclosure is unchanged.
  expect(parent.detail).toBe("Checklist requirement unresolved across its refined checks (not satisfied): The note states when delivery starts. A refined check remains adverse.")
  requestMemo(directory)
  const reason = "Reason: Part of it was not confirmed as met."
  // Was "- The note states when delivery starts. Status: … Source: …", with no reason.
  expect(deliveredNotes(directory, gate)!.split("\n").find(row => row.startsWith("- The note states when delivery starts.")))
    .toBe(`- The note states when delivery starts. ${OMISSION} ${reason} Source: request: "Deliver within ten days of the order."`)
  // The checklist output no longer holds it: the reason is read from the detail.
  writeLeanText(directory, "checklist/extraction.json", JSON.stringify({ expected_sources: [], sources: [] }))
  expect(deliveredNotes(directory, { ...gate, items: [parent] })!.split("\n").at(-2)).toBe(`- The note states when delivery starts. ${OMISSION} ${reason}`)
})

test("through the run, the registered memo has no markers or ids while the working notes keep them", async () => {
  const directory = matter()
  writeLeanText(directory, "request.md", "Draft a short supply note from the supply terms, with a cover memo on the open points.")
  const draft = drafter(DRAFT)
  const memoDrafter: Reply = (prompt, index) => draft(prompt, index).replace("Acting for: not stated", "Acting for: not stated\nMemo: cover memo")
  const changes = sameSession(prompt => JSON.stringify({ amendments: [{ pid: pidOf(prompt, "The supplier shall deliver the goods."),
    updated: "The supplier shall deliver the goods in good condition [open item: confirm the condition standard chosen].", reason: "Drafting judgment.", sources: [] }], questions: [] }))
  const memo: Reply = prompt => ["# Cover memo", "", "The draft follows the supply terms.", "",
    ...[...prompt.matchAll(/<open_items>\n([\s\S]*?)\n<\/open_items>/g)].flatMap(match => [...match[1]!.matchAll(/^- (\S+):/gm)])
      .map(row => `- The condition standard is for the lawyer to confirm. [ref: ${row[1]}]`)].join("\n")
  const deliverables = registering()
  const model = scriptedModel({ ...happyPath, [AGENT.draft]: memoDrafter, [AGENT.changes]: changes, [AGENT.finalize]: memo })
  expect((await runLean(directory, { sessions: model.port, registerDeliverable: deliverables.registerDeliverable })).status).toBe("completed")

  expect(memoMode(directory)).toBe(true)
  const item = readOpenItems(directory).find(row => row.kind === "placeholder" && row.stage === "apply")!
  const working = readLeanText(directory, "drafting-notes.md")!
  expect(working).toContain(`[ref: ${item.id}]`)
  expect(readLeanText(directory, "delivered-notes.md")).toBe(withoutRefMarkers(working))
  const name = deliverables.names.find(name => name.endsWith("cover memo.docx"))!
  const delivered = docxText(readFileSync(path.join(directory, name)))
  expect(delivered).toContain("The condition standard is for the lawyer to confirm.")
  for (const text of ["[ref", item.id]) expect(delivered).not.toContain(text)
}, 120_000)
