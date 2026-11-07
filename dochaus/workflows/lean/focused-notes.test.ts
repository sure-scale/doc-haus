import { afterEach, expect, test } from "bun:test"
import { delegatedDecisionsText, stagePrompt } from "./prompts"
import { readLeanPlan, writeLeanPlan } from "./plan"
import { readLeanState, readLeanText, writeLeanText, writeOpenItems, writeQuestions, type OpenItem, type Question } from "./state"
import { paragraphSha256 } from "../../lib/drafting-amendment"
import { checkDraftingNotes } from "./checks"
import { assembleFocusedNotes, checkFocusedNotesSection, focusedNotesInputs, runFocusedNotes } from "./focused-notes"
import { LEAN_MAX_RETRIES, type LeanDeps, type StageSessionPort } from "./stages"
import { cleanupMatters, matter, SOURCE } from "./test-support"

afterEach(cleanupMatters)

const CHOICE: Question = {
  id: "q-layout", stage: "intake", kind: "judgment", status: "answered",
  question: "Should the handover use a table?", context: "The instructions permit either layout.",
  answer: "Use prose paragraphs.", answer_kind: "adopted", answer_class: "delegated", option_id: "o1",
  lawyer_words: "Use your judgment on the layout.",
  advice: { options: [{ id: "o1", text: "Use prose paragraphs." }, { id: "o2", text: "Use a table." }],
    recommendation: { option: "o1", rationale: "The handover has one narrative instruction.", basis: [] } },
}

test("notes preserve an adopted drafting judgment without reopening it or a source-stated client side", () => {
  const directory = matter("Keep the archive label on the container.", "Prepare a handover note for the archive custodian.")
  writeLeanPlan(directory, [{ title: "Archive handover", purpose: "Explain the container handover.", parts: [{ heading: "Container", scope: "Label instruction." }] }], undefined,
    { party: "Archive custodian", basis: "stated" })
  writeQuestions(directory, [CHOICE])
  const decision = delegatedDecisionsText(directory)
  expect(decision).toContain(CHOICE.answer!)
  expect(decision).toContain("The handover has one narrative instruction.")
  expect(decision).toContain("Alternative not taken: Use a table.")
  expect(decision).not.toContain("for the lawyer to confirm")
  const prompt = stagePrompt(directory, "finalize")
  expect(prompt).not.toContain("Where acting_for says the side was taken from the sources, say so once, with that evidence, and that the lawyer must confirm it")
  expect(prompt).not.toContain("that the lawyer delegated it, and that it is for the lawyer to confirm")
})

const OPEN_POINTS: OpenItem[] = [
  { id: "spine:omission:archive-label", stage: "spine", kind: "omission", detail: "The label colour is not supplied; the custodian must supply it.", ref: { repair_target: "memo" }, status: "open" },
  { id: "spine:omission:archive-seal", stage: "spine", kind: "omission", detail: "The seal design is not supplied; the custodian must supply it.", ref: { repair_target: "memo" }, status: "open" },
]
const row = (pid: string, text: string, part = 0) => ({ pid, text, sha256: paragraphSha256(text), draft_part: part })

function noteMatter(items: OpenItem[] = []) {
  const directory = matter("Keep the archive label on the container.", "Prepare a handover note for the archive custodian.")
  writeLeanPlan(directory, [{ title: "Archive handover", purpose: "Explain the container handover.", parts: [{ heading: "Container", scope: "Label instruction." }] }])
  const paragraphs = [row("p111111111111", "Keep the archive label on the container.")]
  writeLeanText(directory, "d1/paragraphs.json", JSON.stringify(paragraphs))
  writeLeanText(directory, "d1/paragraphs-applied.json", JSON.stringify(paragraphs))
  writeOpenItems(directory, items)
  return directory
}

function scriptedSections(answer: (input: { text: string; sessionID: string; initial: string; attempt: number }) => string | Promise<string>) {
  const prompts: { text: string; sessionID: string; initial: string; attempt: number }[] = []
  const sessions: Parameters<StageSessionPort["createSession"]>[0][] = []
  const first = new Map<string, string>()
  const port: StageSessionPort = {
    async createSession(input) { sessions.push(input); return `notes-${sessions.length}` },
    async prompt(input) {
      if (!first.has(input.sessionID)) first.set(input.sessionID, input.text)
      const call = { text: input.text, sessionID: input.sessionID, initial: first.get(input.sessionID)!, attempt: prompts.filter(row => row.sessionID === input.sessionID).length }
      prompts.push(call)
      return answer(call)
    },
  }
  return { deps: { sessions: port } as LeanDeps, prompts, sessions }
}

test("code attaches each full input reference once while prose jobs receive no reference-copying task", () => {
  const directory = noteMatter(OPEN_POINTS)
  writeQuestions(directory, [CHOICE])
  const inputs = focusedNotesInputs(directory)
  const disclosures = inputs.filter(input => input.kind === "open-item" || input.kind === "delegated")
  expect(disclosures).toHaveLength(3)
  for (const input of disclosures) {
    expect(input.text).not.toContain("[ref:")
    for (const ref of input.refs) expect(input.text).not.toContain(ref)
    expect(checkFocusedNotesSection(directory, input, "The custodian must address this point before issue.")).toEqual([])
  }
  const unmarked = "# Drafting notes\n\nThe custodian must address the open points."
  expect(checkDraftingNotes(directory, unmarked, OPEN_POINTS).issues.some(issue => issue.includes("do not mention these open items"))).toBe(true)
  const notes = assembleFocusedNotes(directory, inputs.map(input => ({ input, prose: "The custodian must address this point before issue." })))
  expect(checkDraftingNotes(directory, notes, OPEN_POINTS).issues).toEqual([])
  for (const ref of [...OPEN_POINTS.map(item => item.id), CHOICE.id]) expect(notes.split(`[ref: ${ref}]`)).toHaveLength(2)
  expect(checkFocusedNotesSection(directory, disclosures[0]!, "A point. [ref: invented]")).toEqual(["Write this section's explanation only. The coordinator attaches its reference markers; do not supply markers or coverage bookkeeping."])
})

test("section validation preserves exact quote and unsupported-authority checks", () => {
  const directory = noteMatter()
  const input = focusedNotesInputs(directory).find(input => input.kind === "passage")!
  expect(checkFocusedNotesSection(directory, input, `The instruction is "Keep the archive label on the container." [source: ${SOURCE}].`)).toEqual([])
  expect(checkFocusedNotesSection(directory, input, `The instruction is "Remove the archive label." [source: ${SOURCE}].`).some(issue => issue.includes("those words are not in that document"))).toBe(true)
  expect(checkFocusedNotesSection(directory, input, "In Lumen v. Solace, the custodian must use that label.").some(issue => issue.includes("no matter source"))).toBe(true)
  expect(checkFocusedNotesSection(directory, input, "In Lumen v. Solace, the custodian must use that label.", true).some(issue => issue.includes("no matter source"))).toBe(true)
})

test("an unsupported section retries alone and an unchanged resume reuses every accepted section", async () => {
  const directory = noteMatter(OPEN_POINTS)
  const model = scriptedSections(input => input.initial.includes(OPEN_POINTS[1]!.detail) && input.attempt === 0
    ? `The instruction is "Remove the archive label." [source: ${SOURCE}].`
    : "The custodian must address the stated container instruction.")
  const result = await runFocusedNotes(directory, readLeanState(directory)!, model.deps, () => {})
  expect(result.status).toBe("done")
  const retried = model.prompts.filter(input => input.initial.includes(OPEN_POINTS[1]!.detail))
  expect(retried).toHaveLength(2)
  expect(new Set(retried.map(input => input.sessionID)).size).toBe(1)
  expect(retried[1]!.text).toContain("Correct only this section's errors")
  for (const session of new Set(model.prompts.filter(input => !input.initial.includes(OPEN_POINTS[1]!.detail)).map(input => input.sessionID)))
    expect(model.prompts.filter(input => input.sessionID === session)).toHaveLength(1)
  const acceptedCalls = model.prompts.length
  expect((await runFocusedNotes(directory, readLeanState(directory)!, model.deps, () => {})).status).toBe("done")
  expect(model.prompts).toHaveLength(acceptedCalls)
  const notes = readLeanText(directory, "drafting-notes.md")!
  for (const item of OPEN_POINTS) expect(notes).toContain(`[ref: ${item.id}]`)
  expect(notes).not.toContain("Remove the archive label")
})

test("a section exhausting its retries preserves accepted neighbors for recovery", async () => {
  const directory = noteMatter(OPEN_POINTS)
  const broken = scriptedSections(input => input.initial.includes(OPEN_POINTS[1]!.detail)
    ? "In Lumen v. Solace, the custodian must use that label."
    : "The custodian must address the stated container instruction.")
  expect((await runFocusedNotes(directory, readLeanState(directory)!, broken.deps, () => {})).status).toBe("failed")
  expect(broken.prompts.filter(input => input.initial.includes(OPEN_POINTS[1]!.detail))).toHaveLength(LEAN_MAX_RETRIES + 1)
  const recovered = scriptedSections(() => "The seal design remains for the custodian to supply.")
  expect((await runFocusedNotes(directory, readLeanState(directory)!, recovered.deps, () => {})).status).toBe("done")
  expect(recovered.prompts).toHaveLength(1)
  expect(recovered.prompts[0]!.text).toContain(OPEN_POINTS[1]!.detail)
})

test("a changed final definition invalidates a cached explanation of its dependent passage", async () => {
  const directory = noteMatter()
  const paragraphs = [row("p111111111111", "Review Window means ten working days.", 0), row("p222222222222", "The handover follows the Review Window.", 1)]
  writeLeanText(directory, "d1/paragraphs.json", JSON.stringify(paragraphs))
  writeLeanText(directory, "d1/paragraphs-applied.json", JSON.stringify(paragraphs))
  const first = scriptedSections(() => "The handover follows the stated review window.")
  expect((await runFocusedNotes(directory, readLeanState(directory)!, first.deps, () => {})).status).toBe("done")
  const changed = [row(paragraphs[0]!.pid, "Review Window means twenty working days.", 0), paragraphs[1]!]
  writeLeanText(directory, "d1/paragraphs-applied.json", JSON.stringify(changed))
  const second = scriptedSections(() => "The handover follows the revised review window.")
  expect((await runFocusedNotes(directory, readLeanState(directory)!, second.deps, () => {})).status).toBe("done")
  expect(second.prompts.some(input => input.initial.includes("The handover follows the Review Window."))).toBe(true)
  expect(readLeanText(directory, "drafting-notes.md")).not.toContain("follows the stated review window")
})

test("unscoped notes repairs replace the linked wrong section and preserve accepted neighbors", async () => {
  const directory = noteMatter()
  const wrong = "The archive custodian is still unidentified."
  const first = scriptedSections(input => input.initial.includes("<section_kind>\npackage\n") ? wrong : "The note preserves the container label instruction.")
  expect((await runFocusedNotes(directory, readLeanState(directory)!, first.deps, () => {})).status).toBe("done")
  expect(readLeanText(directory, "drafting-notes.md")).toContain(wrong)
  const repair: OpenItem = { id: "package-review:omission:custodian", stage: "package-review", kind: "omission", status: "open", detail: "Correct the package overview: the instructions identify the archive custodian; remove the statement that the custodian is unidentified.", ref: { repair_target: "memo" } }
  writeLeanText(directory, "repairs.json", JSON.stringify({ round: 1, requested_at: "synthetic", items: [repair] }))
  const corrected = scriptedSections(input => input.initial.includes("separate repair-link job")
    ? JSON.stringify({ sections: [0], reason: "The package overview contains the erroneous audience statement." })
    : "The note is drafted for the archive custodian, as instructed.")
  expect((await runFocusedNotes(directory, readLeanState(directory)!, corrected.deps, () => {})).status).toBe("done")
  expect(corrected.prompts).toHaveLength(2)
  expect(corrected.prompts.filter(input => !input.initial.includes("separate repair-link job"))[0]!.text).toContain(`<required_repairs>\n- ${repair.detail}`)
  const notes = readLeanText(directory, "drafting-notes.md")!
  expect(notes).not.toContain(wrong)
  expect(notes).toContain("drafted for the archive custodian")
  expect(notes).not.toContain("Corrected explanation")
})

test("long final bodies split by part and passage without copying every body into each job", () => {
  const directory = noteMatter()
  const paragraphs = Array.from({ length: 15 }, (_, index) => row(`p${index.toString(16).padStart(12, "0")}`, `Passage ${index}. ${"Archive handling instruction. ".repeat(45)}`, index < 7 ? 0 : 1))
  const plan = readLeanPlan(directory)!
  writeLeanPlan(directory, plan.documents.map(({ id, ...document }) => ({ ...document, parts: [...document.parts, { heading: "Handover", scope: "Handover details." }] })))
  writeLeanText(directory, "d1/paragraphs-applied.json", JSON.stringify(paragraphs))
  const inputs = focusedNotesInputs(directory)
  const passages = inputs.filter(input => input.kind === "passage")
  expect(passages.length).toBeGreaterThan(2)
  expect(passages.flatMap(input => input.pids).toSorted()).toEqual(paragraphs.map(row => row.pid).toSorted())
  expect(passages.every(input => input.text.length < 10_000)).toBe(true)
  expect(inputs.find(input => input.kind === "package")!.text).not.toContain(paragraphs.at(-1)!.text)
  expect(passages[0]!.text).not.toContain(paragraphs.at(-1)!.text)
})

test("every relevant notes assignment receives current short human instructions and answers directly", () => {
  const directory = noteMatter(OPEN_POINTS)
  const update = "Explain the container exception in plain words in the drafting notes."
  writeLeanText(directory, "updated-notes.md", update)
  writeQuestions(directory, [{ ...CHOICE, id: "q-audience", answer_kind: "option", answer_class: "answered", question: "Who reads the handover?", answer: "The archive custodian reads the handover.", document: "d1" }])
  for (const input of focusedNotesInputs(directory)) {
    expect(input.text).toContain("<current_request>\nPrepare a handover note for the archive custodian.")
    expect(input.text).toContain(`<current_updates>\n${update}`)
    expect(input.text).toContain("The archive custodian reads the handover.")
    expect(input.text).not.toContain("q-audience")
  }
})

test("oversized instructions remain bounded and explicitly require current targeted retrieval", () => {
  const directory = noteMatter()
  const start = "Request opening for archive handling."
  const end = "In the drafting notes, explain the container exception."
  writeLeanText(directory, "request.md", `${start}\n${"Detailed source-handling instruction.\n".repeat(500)}${end}`)
  for (const input of focusedNotesInputs(directory)) {
    expect(input.text).toContain("<current_request_navigation>")
    expect(input.text).toContain(start)
    expect(input.text).toContain(end)
    expect(input.text).toContain("navigation only and do not replace")
    expect(input.text.length).toBeLessThan(12_000)
  }
})

test("changed acting-for instructions invalidate dependent cached notes without changing the body", async () => {
  const directory = noteMatter()
  const first = scriptedSections(() => "The note explains the container instruction.")
  expect((await runFocusedNotes(directory, readLeanState(directory)!, first.deps, () => {})).status).toBe("done")
  const plan = readLeanPlan(directory)!
  writeLeanPlan(directory, plan.documents.map(({ id, ...document }) => document), plan.basis, { party: "Container keeper", basis: "stated" })
  const second = scriptedSections(() => "The note explains the instruction for the container keeper.")
  expect((await runFocusedNotes(directory, readLeanState(directory)!, second.deps, () => {})).status).toBe("done")
  expect(second.prompts.some(input => input.initial.includes("<section_kind>\npassage\n"))).toBe(true)
})
