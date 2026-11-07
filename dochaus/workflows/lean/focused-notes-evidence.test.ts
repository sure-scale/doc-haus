import { expect, test } from "bun:test"
import { currentNotesPids, notesCurrentEvidence, notesEvidenceUnits, checkNotesTargets, type NotesBody, type NotesTarget } from "./focused-notes-evidence"
import type { ParagraphRow } from "./checks"

const row = (pid: string, text: string): ParagraphRow => ({ pid, text, sha256: "synthetic", draft_part: 0 })
const body: NotesBody = { document: "d1", initial: [row("p111111111111", "Acceptance is immediate.")], current: [
  row("p222222222222", "The custodian accepts the container."),
  row("p333333333333", "Acceptance takes effect only after inspection."),
], applied: [{ id: "a1", pid: "p111111111111", mode: "code", result_pids: ["p222222222222", "p333333333333"] }] }

test("split application results locate current targets and preserve the actual generated initial baseline", () => {
  expect(currentNotesPids(body, ["p111111111111"])).toEqual({ pids: ["p222222222222", "p333333333333"], missing: [], deleted: [] })
  const context = notesCurrentEvidence(body, ["p222222222222"])
  expect(context.find(row => row.name === "generated_initial_evidence")!.text).toContain("Acceptance is immediate.")
  expect(context.find(row => row.name === "current_accepted_evidence")!.text).toContain("only after inspection")
})

test("a recorded deletion differs from an unavailable target and neither proves document-wide absence", () => {
  expect(currentNotesPids({ ...body, applied: [{ ...body.applied![0]!, result_pids: [] }] }, ["p111111111111"]))
    .toEqual({ pids: [], missing: [], deleted: ["p111111111111"] })
  expect(currentNotesPids(body, ["p444444444444"])).toEqual({ pids: [], missing: ["p444444444444"], deleted: [] })
  expect(notesCurrentEvidence(body, ["p444444444444"])[0]!.text).toContain("Missing inspection never proves missing body content")
})

test("target lookup accepts only current PIDs in planned documents and requires an honest unavailable-inspection limit", () => {
  expect(checkNotesTargets({ targets: [{ document: "d1", pids: ["p222222222222", "p222222222222"] }], limitation: "" }, [body]))
    .toEqual({ targets: [{ document: "d1", pids: ["p222222222222"] }], limitation: "", issues: [] })
  for (const value of [
    { targets: [{ document: "d9", pids: ["p222222222222"] }], limitation: "" },
    { targets: [{ document: "d1", pids: ["p111111111111"] }], limitation: "" },
    { targets: [], limitation: "" },
  ]) expect(checkNotesTargets(value, [body]).issues.length).toBeGreaterThan(0)
  expect(checkNotesTargets({ targets: [], limitation: "The current passage could not be located." }, [body]).targets).toEqual([])
})

test("overlapping locator entries merge per document before creating oversized comparison owners", () => {
  const large: NotesBody = { ...body, current: [row("p222222222222", "Container handling instructions. ".repeat(600)), body.current[1]!] }
  const other: NotesBody = { document: "d2", initial: [], current: [row("p222222222222", "A separate document retains its own inspection requirement.")] }
  const targets = [
    { document: "d1", pids: ["p222222222222"] },
    { document: "d2", pids: ["p222222222222"] },
    { document: "d1", pids: ["p222222222222", "p333333333333"] },
    { document: "d1", pids: ["p222222222222"] },
  ]
  const checked = checkNotesTargets({ targets, limitation: "Only the selected passages were inspected." }, [large, other])
  const expected: NotesTarget[] = [{ document: "d1", pids: ["p222222222222", "p333333333333"] }, { document: "d2", pids: ["p222222222222"] }]
  expect(checked).toEqual({ targets: expected, limitation: "Only the selected passages were inspected.", issues: [] })
  const units = (selected: typeof expected) => notesEvidenceUnits(selected.flatMap(target =>
    notesCurrentEvidence(target.document === "d1" ? large : other, target.pids))
    .map(row => `<${row.name}>\n${row.text}\n</${row.name}>`).join("\n\n"))
  expect(units(checked.targets!)).toEqual(units(expected))
  expect(units(checked.targets!).filter(unit => JSON.parse(/<evidence_unit_scope>\n([^\n]+)/.exec(unit.text)![1]!).document === "d1")
    .every(unit => unit.text.includes(body.current[1]!.text))).toBe(true)
})

test("oversized target questions retain the paired baseline and every selected qualification in each comparison", () => {
  const target = "Unabridged handling instruction. ".repeat(600) + "Acceptance remains conditional on inspection."
  const qualification = "No acceptance takes effect until the custodian approves the seal."
  const context = notesCurrentEvidence({ ...body, current: [row("p222222222222", target), row("p333333333333", qualification)] }, ["p222222222222"])
  const text = context.map(row => `<${row.name}>\n${row.text}\n</${row.name}>`).join("\n\n")
  const units = notesEvidenceUnits(text)
  expect(units.length).toBeGreaterThan(2)
  const focus = units.map(unit => JSON.parse(/<evidence_unit_focus>\n([\s\S]*?)\n<\/evidence_unit_focus>/.exec(unit.text)![1]!))
  expect(focus.join("")).toBe(target)
  expect(focus.every(text => text.length <= 4_000)).toBe(true)
  for (const unit of units) {
    expect(unit.text).toContain(qualification)
    expect(unit.text).toContain("Acceptance remains conditional on inspection.")
    expect(unit.text).toContain("Acceptance is immediate.")
    expect(unit.text).toContain('"applied_outcomes"')
    expect(unit.text).toContain("Do not infer a whole-document omission")
  }
})

test("an oversized deleted target remains an evidenced deletion comparison with surviving qualifications", () => {
  const initial = "Earlier container handling terms. ".repeat(600)
  const survivor = row("p333333333333", "The custodian must still inspect the seal.")
  const context = notesCurrentEvidence({ document: "d1", initial: [row("p111111111111", initial), survivor], current: [survivor],
    applied: [{ id: "a1", pid: "p111111111111", mode: "code", result_pids: [] }] }, ["p111111111111"])
  const units = notesEvidenceUnits(context.map(row => `<${row.name}>\n${row.text}\n</${row.name}>`).join("\n\n"))
  expect(units.length).toBeGreaterThan(1)
  for (const unit of units) {
    expect(unit.label).toContain("Deleted initial")
    expect(unit.text).toContain(survivor.text)
    expect(unit.text).toContain('"result_pids":[]')
    expect(unit.text).toContain('"effect":"applied deletion"')
  }
})
