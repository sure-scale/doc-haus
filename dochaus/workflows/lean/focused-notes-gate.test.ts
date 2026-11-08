import { afterEach, expect, test } from "bun:test"
import { existsSync, readFileSync } from "node:fs"
import path from "node:path"
import { readParagraphs } from "./checks"
import { readLeanDocuments } from "./deliverables"
import { NOTES_COVERAGE_AGENT } from "./focused-notes-coverage"
import { enqueueLeanRequest } from "./requests"
import { readRepairRound, startRepairRound } from "./repairs"
import { runLean } from "./runner"
import { LEAN_MAX_RETRIES } from "./stages"
import { leanFile, readLeanJson, readLeanText, readQuestions, writeQuestions, type Question } from "./state"
import { AGENT, cleanupMatters, draftingNotes, happyPath, matter, registering, scriptedModel } from "./test-support"

afterEach(cleanupMatters)

const prose = "The accepted change supplies the thirty-day delivery period; payment remains due on delivery."
const claim = "payment remains due on delivery"
const payment = "Payment is due on delivery."
const choice: Question = { id: "q-layout", stage: "intake", status: "answered", kind: "judgment",
  question: "Which layout should the note use?", context: "Either layout is permitted.", answer: "Use prose paragraphs.",
  answer_kind: "option", option_id: "o1", options: [{ id: "o1", text: "Use prose paragraphs." }, { id: "o2", text: "Use a table." }], answered_at: "2026-10-01T10:00:00.000Z" }

for (const passes of [true, false]) test(`saved finalize failure resumes through actual gate only with a checked explanation (${passes ? "pass" : "control: no pass"})`, async () => {
  const directory = matter()
  writeQuestions(directory, [choice])
  // Exercise recovery with an already consumed ordinary repair allowance,
  // using the same product API as an admitted substantive repair.
  startRepairRound(directory, [], true)
  startRepairRound(directory, [], true)
  const registration = registering()
  let phase: "initial" | "resume" = "initial"
  const first = new Map<string, string>()
  const ownerWrites: string[] = [], ownerChecks: string[] = []
  let inspections = 0
  const model = scriptedModel({ ...happyPath,
    [AGENT.finalize]: (text, index, sessionID) => {
      if (!first.has(sessionID!)) first.set(sessionID!, text)
      const assignment = first.get(sessionID!)!
      if (assignment.includes("<evidence_request>")) {
        inspections++
        const target = readParagraphs(directory, "d1", "paragraphs-applied.json").find(row => row.text === payment)!
        expect(target).toBeDefined()
        return JSON.stringify({ targets: [{ document: "d1", pids: [target.pid] }], limitation: "" })
      }
      if (assignment.includes("<section_kind>\nchange\n")) {
        ownerWrites.push(text)
        return phase === "initial" ? "The change has been completed." : prose
      }
      return draftingNotes(assignment, index, sessionID)
    },
    [NOTES_COVERAGE_AGENT]: text => {
      const proposed = /<proposed_prose>\n([\s\S]*?)\n<\/proposed_prose>/.exec(text)?.[1]
      if (proposed === "The change has been completed.")
        return JSON.stringify({ status: "rewrite", reason: "The prose claims completion without explaining the changed delivery period or its practical effect." })
      if (proposed === prose) {
        ownerChecks.push(text)
        if (!text.includes("dependency_evidence") || !passes)
          return JSON.stringify({ status: "needs-evidence", reason: "Check the exact current payment wording before accepting this qualification.", request: { document: "d1", claim, question: "What current payment term qualifies this change explanation?" } })
        expect(text).toContain(payment)
        return JSON.stringify({ status: "pass", reason: "The unchanged prose explains the accepted delivery change and its payment qualification is supported by the inspected current body." })
      }
      return JSON.stringify({ status: "pass", reason: "The synthetic prose explains this separate assigned point." })
    },
  })
  const failed = await runLean(directory, { sessions: model.port, ...registration })
  expect(failed).toMatchObject({ status: "failed", stop_reason: expect.stringContaining("does not substantively explain") })
  expect(failed.stages.finalize?.status).toBe("failed")
  expect(failed.stages["d1.apply"]?.status).toBe("done")
  expect(model.count(NOTES_COVERAGE_AGENT)).toBeGreaterThan(0)
  expect(registration.names).toEqual([])
  expect(existsSync(leanFile(directory, "gate.json"))).toBe(false)
  const prefix = Object.fromEntries(Object.entries(failed.stages).filter(([, stage]) => stage?.status === "done")
    .map(([step, stage]) => [step, { input_sha: stage!.input_sha, finished_at: stage!.finished_at }]))
  const body = readFileSync(leanFile(directory, "d1/body-applied.docx"))
  const answers = readLeanText(directory, "questions.json")
  const repair = readRepairRound(directory)
  expect(repair?.substantive_attempts).toBe(2)
  const writtenBefore = ownerWrites.length
  phase = "resume"
  enqueueLeanRequest(directory, { kind: "continue" })
  const resumed = await runLean(directory, { sessions: model.port, ...registration })
  for (const [step, accepted] of Object.entries(prefix))
    expect(resumed.stages[step as keyof typeof resumed.stages]).toMatchObject(accepted)
  expect(readFileSync(leanFile(directory, "d1/body-applied.docx"))).toEqual(body)
  expect(readLeanText(directory, "questions.json")).toBe(answers)
  expect(readQuestions(directory)).toEqual([choice])
  expect(readRepairRound(directory)).toEqual(repair)
  expect(ownerWrites.length - writtenBefore).toBe(1)
  expect(ownerChecks.length).toBeGreaterThanOrEqual(2)
  expect(ownerChecks.every(text => text.includes(`<proposed_prose>\n${prose}\n</proposed_prose>`))).toBe(true)
  expect(ownerChecks[0]).not.toContain("dependency_evidence")
  expect(ownerChecks[1]).toContain("dependency_evidence")
  if (!passes) {
    expect(resumed.status).toBe("failed")
    expect(resumed.stop_reason).toContain("still needs current evidence")
    expect(inspections).toBe(LEAN_MAX_RETRIES)
    expect(registration.names).toEqual([])
    expect(existsSync(leanFile(directory, "gate.json"))).toBe(false)
    return
  }
  expect(resumed.status).toBe("completed")
  expect(inspections).toBe(1)
  for (const step of ["finalize", "checklist-check", "package-review", "gate"] as const)
    expect(resumed.stages[step]?.status).toBe("done")
  expect(readLeanJson(directory, "gate.json")).toMatchObject({ status: "pass", checked_at: expect.any(String) })
  expect(model.count(AGENT["package-review"])).toBe(1)
  expect(readLeanText(directory, "drafting-notes.md")).toContain(prose)
  expect(readFileSync(leanFile(directory, "drafting-notes.docx")).subarray(0, 2).toString()).toBe("PK")
  const delivered = readLeanDocuments(directory)!
  expect(delivered.status).toBe("completed")
  const notes = delivered.documents.find(row => row.kind === "drafting_notes")!
  expect(notes.deliverable).toBeTruthy()
  expect(existsSync(path.join(directory, notes.deliverable!))).toBe(true)
  expect(registration.names).toContain(notes.deliverable!)
}, 120_000)
