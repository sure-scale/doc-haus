import { afterEach, describe, expect, test } from "bun:test"
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { paragraphSha256 } from "../../lib/drafting-amendment"
import { checkChanges } from "./change-checks"
import { draftStatusItems, draftStatusStatements } from "./draft-status"
import { writeLeanPlan, type PlannedDocument } from "./plan"
import { finalGate } from "./gate"
import { changesRepairSections } from "./repairs"
import { replaceStageOpenItems, writeLeanText } from "./state"
import { cleanupMatters, matter, SOURCE } from "./test-support"

// #286: an agreement's text holds its terms; its draft status, approval state
// and legend are the firm's furniture and the drafting notes. Synthetic text only.

const workspace = process.env.WORKSPACE_ROOT
const roots: string[] = []
afterEach(() => {
  cleanupMatters()
  roots.splice(0).forEach(root => rmSync(root, { recursive: true, force: true }))
  if (workspace === undefined) delete process.env.WORKSPACE_ROOT
  else process.env.WORKSPACE_ROOT = workspace
})

const STATUS = [
  "CONSULTATION DRAFT — NOT FOR SIGNATURE",
  "This draft is circulated for discussion only and has not been approved for execution by either party, and is not ready for signing.",
  "(D) Clause 6 and Schedule 3 are explicit placeholders pending the valuation, as set out in this draft.",
  "It is a working draft and is not approved for signature, completion or implementation. [open item: confirm the start date]",
  "Draft — 1 March 2026",
  // Status wording inside an open-item marker: the marker says what is missing.
  "Signed by the Supplier's authorised signatory. [open item: confirm the signatory's authority before execution — draft not ready]",
  "The Price is set out in Schedule 1 (open item: confirm the Price. This version is not approved for signature).",
  "The Term is two years. [open item: first draft; confirm the Term with the client]",
]

const OPERATIVE = [
  "4.1 The Provider shall deliver the Services from the Start Date in accordance with Schedule 1.",
  "This agreement is subject to contract and shall not bind either party until it has been signed by both parties.",
  "Nothing in this agreement gives either party authority to bind the other.",
  "The Consultant shall circulate a draft for discussion at least five Business Days before each Steering Meeting.",
  "\"Working Draft\" means the version of the Annual Plan most recently circulated under clause 7.2.",
  "The Provider shall submit the draft Specification to the Customer for approval [open item: confirm the review period].",
  // Operative conditions, defined terms and mechanisms that share status words.
  "The Guarantee is not ready for signing until the Bank's credit approval is received.",
  "The Supplier shall deliver a Draft Status Report to the Client monthly.",
  "Clause 9 is a placeholder pending the Parties' agreement of the Service Credit regime.",
  "Draft-form Schedules listed in Schedule 2 are incorporated.",
  "This is a mutual NDA covering information the Discloser provided for discussion of a possible acquisition.",
  "The Parties shall settle the bracketed provisions in Schedule 4 in good faith before signature.",
  // Markers that say what is missing, including a marker describing itself.
  "The Supplier shall notify the Customer of any defect. [open item: confirm whether the draft Specification is approved before delivery]",
  "Notices go to the address in Schedule 2. [open item: placeholder — insert the Supplier's notice address]",
  "The Customer shall pay within 30 days. [open item: confirm what happens if the draft budget is not approved by the Board]",
]

function agreement(directory: string, extra: Partial<Omit<PlannedDocument, "id">> = {}) {
  writeLeanPlan(directory, [{ title: "Services agreement", purpose: "Terms of the services", parts: [{ heading: "Terms", scope: "All terms" }], form: "agreement", ...extra }])
}

function rows(texts: string[]) {
  return texts.map((text, index) => ({ pid: `p${String(index + 1).padStart(12, "0")}`, text, sha256: paragraphSha256(text) }))
}

function firmLegend(legend: string) {
  const root = mkdtempSync(path.join(tmpdir(), "draft-status-prefs-"))
  roots.push(root)
  mkdirSync(path.join(root, ".preferences"))
  writeFileSync(path.join(root, ".preferences", "preferences.json"), JSON.stringify({ agreementLegend: legend }))
  process.env.WORKSPACE_ROOT = root
}

describe("draft-status statements in an agreement", () => {
  test("a paragraph stating the agreement's draft status or approval state is flagged", () => {
    for (const text of STATUS) expect(draftStatusStatements(text).length, text).toBeGreaterThan(0)
  })

  test("operative clauses, binding-effect terms, other documents' drafts, quoted terms and markers are not flagged", () => {
    for (const text of OPERATIVE) expect(draftStatusStatements(text), text).toEqual([])
  })

  test("the firm's legend applied by code and the lawyer's own marking in the original being revised are kept", () => {
    firmLegend("For discussion purposes only")
    const directory = matter("Draft for discussion | 1 March 2026\n\nThe Provider shall deliver the Services.")
    agreement(directory, { mode: "revise", comparison: SOURCE })
    const body = rows(["For discussion purposes only", "Draft for discussion | 1 March 2026", ...OPERATIVE])
    expect(draftStatusItems(directory, "d1", body, "apply")).toEqual([])
    // The same marking written into a new agreement is the model's, not the lawyer's.
    agreement(directory)
    expect(draftStatusItems(directory, "d1", body, "apply").map(item => item.ref.pid)).toEqual([body[1]!.pid])
  })

  test("in a revision only a sentence taken from one paragraph of the original is kept, not the same words written elsewhere", () => {
    const original = "Draft for discussion | 1 March 2026\n\nThe Provider shall deliver the Services.\n\nIt is a working draft subject to the points in the review note."
    expect(draftStatusStatements("DRAFT FOR DISCUSSION — 1 March 2026", [original])).toEqual([])
    expect(draftStatusStatements("It is a working draft subject to the points in the review note.", [original])).toEqual([])
    expect(draftStatusStatements("(B) The Supplier will supply the Goods, and this document remains a working draft.", [original]).length).toBeGreaterThan(0)
  })

  test("a letter may refer to a draft for discussion; only an agreement's own text is checked", () => {
    const directory = matter()
    writeLeanPlan(directory, [{ title: "Cover letter", purpose: "Send the draft", parts: [{ heading: "Letter", scope: "Letter" }], form: "letter" }])
    expect(draftStatusItems(directory, "d1", rows(["We enclose a draft for discussion. It is not approved for signature."]), "draft")).toEqual([])
  })

  test("status wording inside a marker reaches the change cycle; the marker's open point does not", () => {
    const directory = matter()
    agreement(directory)
    const draft = rows([STATUS[5]!, OPERATIVE[0]!])
    const items = draftStatusItems(directory, "d1", draft, "draft")
    expect(items).toEqual([expect.objectContaining({ kind: "draft_status", ref: expect.objectContaining({ pid: draft[0]!.pid, statement: "draft not ready" }) })])
    expect(items[0]!.detail).toContain("never the draft's status")
    expect(draftStatusItems(directory, "d1", rows([STATUS[5]!.replace(" — draft not ready", "")]), "draft")).toEqual([])
  })

  test("a draft's status wording reaches the change cycle as a required repair", () => {
    const directory = matter()
    agreement(directory)
    const draft = rows([STATUS[0]!, OPERATIVE[0]!])
    writeLeanText(directory, "d1/paragraphs.json", JSON.stringify(draft))
    const items = draftStatusItems(directory, "d1", draft, "draft")
    expect(items).toEqual([expect.objectContaining({ kind: "draft_status", ref: expect.objectContaining({ pid: draft[0]!.pid }) })])
    replaceStageOpenItems(directory, "draft", "d1", items)
    expect(changesRepairSections(directory, "d1").join("\n")).toContain(`${items[0]!.id} [${draft[0]!.pid}]`)
  })

  test("an amendment that writes status wording into an agreement is retried, then withheld; an operative amendment is admitted", () => {
    const directory = matter()
    agreement(directory)
    const draft = rows([OPERATIVE[0]!])
    const amend = (updated: string) => JSON.stringify({ amendments: [{ pid: draft[0]!.pid, updated, reason: "Review finding.", sources: [] }], questions: [] })
    const status = `${OPERATIVE[0]} This version remains a working draft and is not approved for signature. [open item: confirm the reporting cycle]`
    expect(checkChanges(directory, amend(status), draft, { document: "d1" }).issues.join(" ")).toContain("draft status")
    const last = checkChanges(directory, amend(status), draft, { document: "d1", withholdUnverified: true })
    expect(last.changes?.amendments).toEqual([])
    // Withheld, it is disclosed through the drafting notes with its reason.
    expect(last.openItems).toContainEqual(expect.objectContaining({ kind: "unapplied_amendment", ref: expect.objectContaining({ pid: draft[0]!.pid }),
      detail: expect.stringContaining("states the agreement's own draft status or approval state") }))
    expect(checkChanges(directory, amend(`${OPERATIVE[0]} The Provider shall report monthly. [open item: confirm the reporting cycle]`), draft, { document: "d1" }).issues).toEqual([])
  })

  test("status wording left in the final agreement holds delivery until the drafting notes disclose it", () => {
    const directory = matter()
    agreement(directory)
    const final = rows([STATUS[1]!, OPERATIVE[0]!])
    writeLeanText(directory, "d1/paragraphs-applied.json", JSON.stringify(final))
    const items = draftStatusItems(directory, "d1", final, "apply")
    replaceStageOpenItems(directory, "apply", "d1", items)
    writeLeanText(directory, "drafting-notes.md", "# Drafting notes\n\nNothing to report.\n")
    expect(finalGate(directory).blocking.map(item => item.kind)).toContain("draft_status")
    writeLeanText(directory, "drafting-notes.md", `# Drafting notes\n\nThe agreement still states its draft status in its opening paragraph [ref: ${items[0]!.id}].\n`)
    expect(finalGate(directory).disclosed).toContainEqual(expect.objectContaining({ id: items[0]!.id, status: "disclosed" }))
  })
})
