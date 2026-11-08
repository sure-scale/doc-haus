import { afterEach, expect, test } from "bun:test"
import { readLeanEvents } from "./events"
import { checkFocusedNotesCoverage, checkNotesCoverageReply, notesCoveragePrompt, NOTES_CHECK_PERMISSIONS, NOTES_COVERAGE_AGENT, type NotesCoverageVerdict } from "./focused-notes-coverage"
import { LEAN_MAX_RETRIES, type StageSessionPort } from "./stages"
import { readLeanState } from "./state"
import { cleanupMatters, matter } from "./test-support"
import { writeLeanPlan } from "./plan"

afterEach(cleanupMatters)

const assignment = { id: "item:source:unknown:seal", kind: "open-item" as const, text: "Explain that the seal design remains unknown and the custodian must supply it; keep the adopted layout choice settled." }
const proposed = "The seal design remains unknown; the custodian must supply it before using the handover."
const passed = { status: "pass" as const, reason: "The prose states the actual missing seal design and the custodian's required action." }

test("only a typed verdict with its explanation answers the coverage question", () => {
  expect(checkNotesCoverageReply(JSON.stringify(passed))).toEqual({ value: passed, issues: [] })
  const rewrite = { status: "rewrite" as const, reason: "The prose claims completion but never explains the unknown seal design or action." }
  expect(checkNotesCoverageReply(JSON.stringify(rewrite))).toEqual({ value: rewrite, issues: [] })
  for (const value of [null, [], {}, { status: "pass" }, { status: "pass", reason: " " }, { status: "unknown", reason: "Uncertain." }, { status: ["pass"], reason: "Looks complete." }, { status: ["rewrite"], reason: "Needs another explanation." }, { status: { value: "pass" }, reason: "Looks complete." }, { status: true, reason: "Looks complete." }]) {
    const checked = checkNotesCoverageReply(JSON.stringify(value))
    expect(checked.value).toBeUndefined()
    expect(checked.issues.length).toBeGreaterThan(0)
  }
  expect(checkNotesCoverageReply("The explanation looks complete.").value).toBeUndefined()
})

test("coverage reads only the scoped proposed prose and accepts honest unknowns without demanding answers or renewed confirmations", () => {
  const prompt = notesCoveragePrompt(assignment, proposed)
  expect(prompt).toContain(`<assigned_writer_packet>\n${assignment.text}`)
  expect(prompt).toContain(`<proposed_prose>\n${proposed}`)
  expect(prompt).toContain("invisible earlier text, errored output")
  expect(prompt).toContain("An honest explanation of an unresolved fact or unavailable authority can pass")
  expect(prompt).toContain("Do not require reference markers")
  expect(prompt).toContain("Preserve an adopted delegated judgment as settled")
  expect(prompt).not.toContain(assignment.id)
})

function verifier(replies: string[]) {
  const sessions: Parameters<StageSessionPort["createSession"]>[0][] = []
  const prompts: Parameters<StageSessionPort["prompt"]>[0][] = []
  const port: StageSessionPort = {
    async createSession(input) { sessions.push(input); return "notes-coverage-session" },
    async prompt(input) { prompts.push(input); return replies[Math.min(prompts.length - 1, replies.length - 1)]! },
  }
  return { sessions, prompts, port }
}

function plannedMatter() {
  const directory = matter()
  writeLeanPlan(directory, [
    { title: "Handover", purpose: "Record the handover condition.", parts: [{ heading: "Inspection", scope: "Inspection condition." }] },
    { title: "Certificate", purpose: "Record the inspection result.", parts: [{ heading: "Result", scope: "Inspection result." }] },
  ])
  return directory
}

test("a malformed verdict retries only the same tool-denied verifier and its valid pass is required", async () => {
  const directory = matter()
  const model = verifier(["The earlier response was complete.", JSON.stringify(passed)])
  expect(await checkFocusedNotesCoverage(directory, readLeanState(directory)!, { sessions: model.port }, assignment, proposed, () => {})).toEqual(passed)
  expect(model.sessions).toHaveLength(1)
  expect(model.sessions[0]!.agent).toBe(NOTES_COVERAGE_AGENT)
  expect(model.sessions[0]!.permissions).toEqual(NOTES_CHECK_PERMISSIONS)
  expect(model.sessions[0]!.permissions).toEqual([{ action: "*", resource: "*", effect: "deny" }])
  expect(model.prompts).toHaveLength(2)
  expect(model.prompts.every(prompt => prompt.agent === NOTES_COVERAGE_AGENT && prompt.sessionID === "notes-coverage-session")).toBe(true)
  expect(model.prompts[1]!.text).toContain("Correct only this explanation-check reply's format")
})

test("a valid rewrite returns once to the writer but exhausted verifier format errors never supply a verdict", async () => {
  const directory = matter()
  const rewrite = { status: "rewrite" as const, reason: "The proposed prose is about an unrelated container colour instead of the missing seal design and action." }
  const negative = verifier([JSON.stringify(rewrite)])
  expect(await checkFocusedNotesCoverage(directory, readLeanState(directory)!, { sessions: negative.port }, assignment, proposed, () => {})).toEqual(rewrite)
  expect(negative.prompts).toHaveLength(1)
  const broken = verifier([JSON.stringify({ status: "pass", reason: "" })])
  const result = await checkFocusedNotesCoverage(directory, readLeanState(directory)!, { sessions: broken.port }, assignment, proposed, () => {})
  expect(result.status).toBe("failed")
  expect(result.reason).toContain("supplied no valid verdict")
  expect(broken.sessions).toHaveLength(1)
  expect(broken.prompts).toHaveLength(LEAN_MAX_RETRIES + 1)
})

test("checker telemetry records only the bounded verdict and timings, keeping explanation prose in repair feedback", async () => {
  const directory = matter()
  const rewrite = { status: "rewrite" as const, reason: "The proposed prose omits the custodian's required action." }
  const verdicts = [passed, rewrite]
  for (const verdict of verdicts) {
    const model = verifier([JSON.stringify(verdict)])
    expect(await checkFocusedNotesCoverage(directory, readLeanState(directory)!, { sessions: model.port }, assignment, proposed, () => {})).toEqual(verdict)
  }
  const events = readLeanEvents(directory)
  const checkers = events.filter(event => event.event === "checker")
  expect(checkers.map(event => event.outcome)).toEqual(["pass", "rewrite"])
  const keys = new Set(["at", "stage", "event", "step", "attempt", "issues", "outcome", "wait_ms", "check_ms"])
  for (const event of checkers) {
    expect(Object.keys(event).filter(key => !keys.has(key))).toEqual([])
    expect(event.stage).toBe("finalize")
    expect(event.step).toBe(`${assignment.id}:explanation-check`)
    expect(event.issues).toEqual([])
    expect(event.wait_ms).toBeGreaterThanOrEqual(0)
    expect(event.check_ms).toBeGreaterThanOrEqual(0)
  }
  const log = JSON.stringify(events)
  for (const text of [assignment.text, proposed, ...verdicts.map(verdict => verdict.reason)]) expect(log).not.toContain(text)
})

test("verification separates current accepted effects from legitimate initial history and uninspected absence", () => {
  const input = { kind: "change" as const, text: "<generated_initial_evidence>Acceptance is immediate.</generated_initial_evidence>\n<current_accepted_evidence>Acceptance is conditional on inspection.</current_accepted_evidence>" }
  const prompt = notesCoveragePrompt(input, "The initial draft was immediate; the accepted revision requires inspection.")
  expect(prompt).toContain(input.text.replaceAll("<", "\\u003c"))
  expect(prompt).toContain("correctly labelled historical explanation is not a current contradiction")
  expect(prompt).toContain("stale deferred record alone cannot establish")
  expect(prompt).toContain("unavailable current inspection cannot establish document-wide absence")
  expect(prompt).toContain("Citation presence does not establish legal correctness or human approval")
})

test("a missing current dependency packet requests one bounded inspection rather than a prose rewrite", async () => {
  const directory = plannedMatter()
  const prose = "The covering note remains subject to the inspection condition in the attached certificate."
  const request = { document: "d2" as const, claim: "the inspection condition in the attached certificate", question: "What current certificate wording qualifies the inspection condition?" }
  const needed: NotesCoverageVerdict = { status: "needs-evidence", reason: "The assigned packet does not include the certificate's current condition.", request }
  expect(checkNotesCoverageReply(JSON.stringify(needed))).toEqual({ value: needed, issues: [] })
  const model = verifier([JSON.stringify(needed)])
  expect(await checkFocusedNotesCoverage(directory, readLeanState(directory)!, { sessions: model.port }, { ...assignment, kind: "change" }, prose, () => {})).toEqual(needed)
  expect(model.prompts).toHaveLength(1)
  expect(model.sessions[0]!.permissions).toEqual(NOTES_CHECK_PERMISSIONS)
  expect(model.prompts[0]!.text).toContain(prose)
})

test("an inspection verdict must name one document and one concrete claim and question", () => {
  const request = { document: "d2", claim: "the inspection condition", question: "What current wording qualifies this condition?" }
  for (const broken of [undefined, [], [request], { ...request, document: "../../other" }, { ...request, document: "d0" }, { ...request, claim: " " }, { ...request, question: "" }]) {
    expect(checkNotesCoverageReply(JSON.stringify({ status: "needs-evidence", reason: "Need the current condition.", request: broken })).value).toBeUndefined()
  }
})

test("reference-aware verdict parsing preserves strict claim case and punctuation against planned IDs", () => {
  const request = { document: "d1" as const, claim: "The seal design remains unknown", question: "What condition applies?" }
  const verdict: NotesCoverageVerdict = { status: "needs-evidence", reason: "Need the current condition.", request }
  const reference = { documents: [{ document: "d1" as const }], prose: proposed }
  expect(checkNotesCoverageReply(JSON.stringify(verdict), reference).value).toEqual(verdict)
  for (const invalid of [{ ...request, claim: "the seal design remains unknown" }, { ...request, claim: "The seal design remains unknown." }, { ...request, document: "d2" }]) {
    const checked = checkNotesCoverageReply(JSON.stringify({ ...verdict, request: invalid }), reference)
    expect(checked.value).toBeUndefined()
    expect(checked.issues.join(" ")).toContain(invalid.document === "d2" ? "request.document" : "request.claim")
  }
})

test("coverage distinguishes the owned change explanation from a dependency's separately owned human action", () => {
  const input = { kind: "change" as const, text: "Explain the accepted seal clause change. The separate open-item section owns the custodian's action to supply the missing design; retain its unresolved qualification here." }
  const prompt = notesCoveragePrompt(input, "The revised seal clause is conditional on the design remaining subject to confirmation.")
  expect(prompt).toContain("A change explanation owns the change's actual effect")
  expect(prompt).toContain("Do not duplicate a separately owned human action")
  expect(prompt).toContain("never treat a missing dependency packet as evidence that the prose is false")
  const open = notesCoveragePrompt(assignment, proposed)
  expect(open).toContain("An open-item owner must explain the actual uncertainty, its practical consequence and the required human action")
})

test("coverage source contract distinguishes recorded verification status from current wording without hiding contradictions or duties", () => {
  const input = { ...assignment, text: "<unresolved_record>Verification of the release condition remains incomplete; the custodian must supply the missing inspection date.</unresolved_record>\n<current_accepted_evidence>The custodian may release the container only after inspection.</current_accepted_evidence>" }
  const prose = "The verification record leaves the inspection date unknown. Obtain the date before deciding whether the release condition is met."
  const prompt = notesCoveragePrompt(input, prose)
  expect(prompt).toContain(prose)
  expect(prompt).toContain("historical research or human-approval status needs its own research, decision or approval record")
  expect(prompt).toContain("precise attribution to a supplied unresolved record")
  expect(prompt).toContain("Current clauses do not establish what research was completed or whether human approval occurred")
  expect(prompt).toContain("not inferred legal validity, invalidity or approval")
  expect(prompt).toContain("Preserve the original assigned duty, practical consequence and required action")
  expect(prompt).toContain("cannot hide a material contradiction or drop an essential exhaustive requirement")
})

test("an inexact dependency claim is corrected inside the same verifier before returning needs-evidence", async () => {
  const directory = plannedMatter()
  const request = { document: "d1" as const, claim: "The seal design remains unknown", question: "What current qualification applies to the seal design?" }
  const needed = { status: "needs-evidence" as const, reason: "One current qualification is outside the packet.", request }
  const model = verifier([JSON.stringify({ ...needed, request: { ...request, claim: "the seal design remains unknown" } }), JSON.stringify(needed)])
  expect(await checkFocusedNotesCoverage(directory, readLeanState(directory)!, { sessions: model.port }, assignment, proposed, () => {})).toEqual(needed)
  expect(model.sessions).toHaveLength(1)
  expect(model.prompts).toHaveLength(2)
  expect(model.prompts.every(prompt => prompt.sessionID === "notes-coverage-session" && prompt.agent === NOTES_COVERAGE_AGENT)).toBe(true)
  expect(model.prompts[1]!.text).toContain("request.claim")
  expect(model.prompts[1]!.text).toContain("including capitalization and punctuation")
  expect(model.prompts[1]!.text).toContain('"status":"needs-evidence"')
  expect(model.prompts[1]!.text).toContain('"question"')
  expect(readLeanEvents(directory).filter(event => event.event === "checker").map(event => event.outcome)).toEqual([undefined, "needs-evidence"])
  expect(JSON.stringify(readLeanEvents(directory))).not.toContain("the seal design remains unknown")
})

test("an unplanned dependency document is corrected in the same verifier without changing the proposed prose", async () => {
  const directory = plannedMatter()
  const request = { document: "d1" as const, claim: "The seal design remains unknown", question: "What current condition applies?" }
  const needed = { status: "needs-evidence" as const, reason: "The relevant condition is not in this packet.", request }
  const model = verifier([JSON.stringify({ ...needed, request: { ...request, document: "d9" } }), JSON.stringify(needed)])
  expect(await checkFocusedNotesCoverage(directory, readLeanState(directory)!, { sessions: model.port }, assignment, proposed, () => {})).toEqual(needed)
  expect(model.sessions).toHaveLength(1)
  expect(model.prompts).toHaveLength(2)
  expect(model.prompts[1]!.text).toContain("request.document")
  expect(model.prompts[1]!.text).toContain("d9")
  expect(model.prompts[1]!.text).toContain("d1, d2")
  expect(model.prompts[0]!.text).toContain(proposed)
})

test("exhausted or missing-plan dependency requests never escape as a verdict", async () => {
  for (const { directory, document } of [{ directory: plannedMatter(), document: "d9" }, { directory: matter(), document: "d1" }]) {
    const model = verifier([JSON.stringify({ status: "needs-evidence", reason: "Need a current condition.", request: {
      document, claim: "The seal design remains unknown", question: "What current condition applies?",
    } })])
    const result = await checkFocusedNotesCoverage(directory, readLeanState(directory)!, { sessions: model.port }, assignment, proposed, () => {})
    expect(result.status).toBe("failed")
    expect(result.reason).toContain("request.document")
    expect(model.sessions).toHaveLength(1)
    expect(model.prompts).toHaveLength(LEAN_MAX_RETRIES + 1)
    expect(readLeanEvents(directory).filter(event => event.event === "checker").every(event => event.outcome === undefined)).toBe(true)
  }
})
