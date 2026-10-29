import { expect, test } from "bun:test"
import type { ChronologyAccount } from "../../../dochaus/lib/chronology"
import { chronologyMatchCandidates } from "./chronology-generation"
import { newChronologyEvent } from "./chronology-store"

function account(title: string, detail = "", status: ChronologyAccount["status"] = "reported"): ChronologyAccount {
  return {
    id: title,
    title,
    detail,
    date: { original: "Undated", precision: "undated" },
    evidence: [],
    people: [],
    privilege: "uncertain",
    significance: "background",
    status,
    gaps: [],
  }
}
function saved(id: string, report: ChronologyAccount) {
  return newChronologyEvent({ id, title: report.title, detail: report.detail, date: report.date, accounts: [report] })
}

test("core act wording recalls an obligation despite repeated parties and project context", () => {
  const context =
    "Fiona Park at Ardent Civil discusses the north district programme, access planning, consultation and the project team's next steps."
  const incoming = account(
    "Freight tunnel closure required before reopening",
    `${context} The instruction requires closing the freight tunnel before residents can return.`,
    "required",
  )
  incoming.people = ["Fiona Park", "Ardent Civil"]
  const target = saved(
    "closure",
    account("Order for tunnel closure", "Traffic manager orders freight tunnel closure.", "required"),
  )
  // The lawyer's current heading need not repeat the source's event title.
  target.title = "Access issue awaiting review"
  const distractors = [
    "Fiona Park sends Ardent Civil procurement email",
    "Fiona Park approves Ardent Civil laboratory invoice",
    "Fiona Park files Ardent Civil insurance certificate",
    "Fiona Park schedules Ardent Civil staff induction",
    "Fiona Park receives Ardent Civil budget forecast",
    "Fiona Park signs Ardent Civil consultancy terms",
    "Fiona Park requests Ardent Civil staffing report",
    "Fiona Park circulates Ardent Civil project minutes",
  ].map((title, index) => saved(`context-${index}`, account(title, context)))
  const candidates = chronologyMatchCandidates(incoming, [...distractors, target])
  expect(candidates).toContain(target)
  expect(candidates.length).toBeLessThanOrEqual(6)
  // Recall is not an identity decision: a completion and the obligation can
  // both be offered for the model's separate completion/conflict decision.
  const completion = saved("completed", account("Freight tunnel closure completed", "Workers closed the tunnel."))
  expect(chronologyMatchCandidates(incoming, [...distractors, target, completion])).toContain(completion)
})

test("superseded events are excluded and a genuinely unrelated date match is empty", () => {
  const incoming = account("Clinic equipment disposed", "2037-04-18")
  const old = saved("superseded", account("Clinic equipment disposed", "2037-04-18"))
  old.supersededBy = ["replacement"]
  const unrelated = saved("floor", account("Warehouse flooring replaced", "2037-04-18"))
  expect(chronologyMatchCandidates(incoming, [old, unrelated])).toEqual([])
})
