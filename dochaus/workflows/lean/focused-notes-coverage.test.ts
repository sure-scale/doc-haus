import { afterEach, expect, test } from "bun:test"
import { readLeanEvents } from "./events"
import { checkFocusedNotesCoverage, checkNotesCoverageReply, notesCoveragePrompt, NOTES_CHECK_PERMISSIONS, NOTES_COVERAGE_AGENT } from "./focused-notes-coverage"
import { LEAN_MAX_RETRIES, type StageSessionPort } from "./stages"
import { readLeanState } from "./state"
import { cleanupMatters, matter } from "./test-support"

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
