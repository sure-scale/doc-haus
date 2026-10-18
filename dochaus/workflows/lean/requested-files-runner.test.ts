import { afterEach, describe, expect, test } from "bun:test"
import { existsSync } from "node:fs"
import { combinedDelivery, memoNotesFile } from "./deliverables"
import { PLAN_FILES_REVISIONS } from "./draft-sections"
import { readLeanEvents } from "./events"
import type { GateResult } from "./gate"
import { parsePlan, readLeanPlan, writeLeanPlan } from "./plan"
import { enqueueLeanRequest } from "./requests"
import { LEAN_REQUESTED_FILES_SCHEMA, readRequestedFiles, requestedFilesItems, writeRequestedFiles, type RequestedFile } from "./requested-files"
import { runLean } from "./runner"
import { LEAN_CHILD_PERMISSIONS, LEAN_MAX_RETRIES, disclosableItems, outlineBasis } from "./stages"
import { leanFile, readLeanJson, readLeanText, readQuestions, writeLeanText } from "./state"
import { AGENT, cleanupMatters, draftingNotes, matter, noChanges, registering, scriptedModel, type Reply } from "./test-support"

// Requested files through the whole run (runLean): the outline's files job,
// its revision requests to the outline session, the plan-time question on a
// mismatch the revisions leave, the names the lawyer's words give the plan's
// files, and the disclosure of a file the run does not deliver as asked. The
// job's prompt, grounding and the derived items are in requested-files.test.ts.
// Synthetic parties and files only; only the model is scripted.

afterEach(cleanupMatters)

const SOURCE_TEXT = "Kestrel Works shall release the deposit once the Orchard Committee approves the revised schedule."
const REQUEST = "Prepare a release letter for Kestrel Works with a cover memo, all in one file named Kestrel Release Pack.docx."
const PACK = "with a cover memo, all in one file named Kestrel Release Pack.docx"
const REVISE = "Your reply did not pass the draft checks:"
const ANSWER = "Deliver them as planned."

const outline = (...lines: string[]) => ["Acting for: not stated", ...lines].join("\n")
const letter = (...lines: string[]) => ["# Release letter", ...lines, "## Release", "The release."].join("\n")
const BODY = outline("Memo: cover memo", letter())
const COMBINED = outline("Memo: cover memo", letter("Deliverable: combined"))

/** The files job's reply naming these files. */
const filesOf = (...files: RequestedFile[]): Reply => () => JSON.stringify({ stated: true, files })
const PACKED: RequestedFile = { holds: ["memo", "d1"], name: "Kestrel Release Pack.docx", quote: PACK }

/** The outline agent: `first` answers the outline prompt, `revised` any
 * revision request, and each part gets one plain paragraph. */
const outliner = (first: string, revised = first): Reply => text =>
  /^Draft part \d+ of \d+ of "/m.test(text) ? "The terms are recorded." : text.startsWith(REVISE) ? revised : first

/** A whole run whose later steps find nothing to change. */
const replies = (extra: Record<string, Reply>): Record<string, Reply> => ({
  [AGENT.review]: () => "No corrections are needed.",
  [AGENT.changes]: noChanges,
  [AGENT.finalize]: draftingNotes,
  ...extra,
})

function prepared(request = REQUEST) {
  const directory = matter(SOURCE_TEXT)
  writeLeanText(directory, "request.md", request)
  return directory
}

type Model = ReturnType<typeof scriptedModel>
const revisions = (model: Model) => model.prompts.filter(prompt => prompt.agent === AGENT.draft && prompt.text.startsWith(REVISE))
const outlinePrompt = (model: Model) => model.prompts.find(prompt => prompt.agent === AGENT.draft)!
const partPrompts = (model: Model) => model.prompts.filter(prompt => /^Draft part \d+ of /m.test(prompt.text))
const filesSessions = (model: Model) => model.sessions.filter(session => session.agent === AGENT.files)
const filesEvents = (directory: string) => readLeanEvents(directory).filter(row => row.event === "requested_files")
const gateOf = (directory: string) => readLeanJson(directory, "gate.json") as GateResult
/** The plan the outline gives as code saved it before the files job. */
function planned(outlineText: string, request: string, directory: string) {
  const parsed = parsePlan(outlineText, request).value!
  const other = matter(SOURCE_TEXT)
  writeLeanPlan(other, parsed.documents, outlineBasis(directory), parsed.acting_for,
    parsed.memo ? { name: parsed.memo, file: memoNotesFile(parsed.documents, parsed.memo) } : undefined)
  return readLeanText(other, "plan.json")
}

describe("requested files through the run", () => {
  test("the memo and one document asked for in one named file: the outline is revised to combined and the file takes the lawyer's name", async () => {
    const directory = prepared()
    const model = scriptedModel(replies({ [AGENT.draft]: outliner(BODY, COMBINED), [AGENT.files]: filesOf(PACKED) }))
    const state = await runLean(directory, { sessions: model.port })
    expect(state.status).toBe("completed")

    // One files job, in a fresh session of its own; one revision, in the outline's session, carrying its one issue.
    expect(filesSessions(model)).toEqual([expect.objectContaining({ title: "Lean drafting: draft: files", permissions: LEAN_CHILD_PERMISSIONS })])
    expect(model.count(AGENT.files)).toBe(1)
    expect(revisions(model)).toEqual([expect.objectContaining({ sessionID: outlinePrompt(model).sessionID })])
    expect(revisions(model)[0]!.text.split("\n").filter(line => line.startsWith("- "))).toEqual([
      `- The lawyer's words "${PACK}" ask for the memo and "Release letter" in one file. Write "Deliverable: combined" under "Release letter".`,
    ])

    const plan = readLeanPlan(directory)!
    expect(plan).toMatchObject({ basis: outlineBasis(directory), memo: "cover memo", notes_file: "Kestrel Release Pack - cover memo.docx" })
    expect(plan.documents).toEqual([expect.objectContaining({ id: "d1", title: "Release letter", deliverable: "combined", file: "Kestrel Release Pack.docx" })])
    expect(readRequestedFiles(directory)).toEqual({ schema: LEAN_REQUESTED_FILES_SCHEMA, basis: plan.basis!, stated: true, files: [PACKED] })
    expect(filesEvents(directory)).toEqual([expect.objectContaining({ stage: "draft", document: "d1", step: "files", outcome: "stated", issues: [],
      reason: '"Release letter": file "Kestrel Release Pack.docx"' })])

    // Delivered as asked: one combined file under the lawyer's name, nothing to disclose or ask.
    expect(combinedDelivery(directory)).toEqual({ document: "d1", name: "Kestrel Release Pack" })
    expect(requestedFilesItems(directory)).toEqual([])
    expect(readQuestions(directory)).toEqual([])
    expect(gateOf(directory)).toMatchObject({ status: "pass", disclosed: [], blocking: [] })
  }, 120_000)

  test("an outline still not combined after the revisions blocks the run on one question; answered, the plan is saved, nothing is asked again and the file is disclosed", async () => {
    const directory = prepared()
    const model = scriptedModel(replies({ [AGENT.draft]: outliner(BODY), [AGENT.files]: filesOf(PACKED) }))
    const state = await runLean(directory, { sessions: model.port })

    expect(state.status).toBe("blocked")
    expect(state.stages["d1.draft"]?.status).toBe("blocked")
    expect(revisions(model)).toHaveLength(PLAN_FILES_REVISIONS)
    // The revisions change neither the documents nor the memo, so the job is not asked again.
    expect(model.count(AGENT.files)).toBe(1)
    const asked = readQuestions(directory)
    expect(asked).toEqual([expect.objectContaining({ id: expect.stringMatching(/^requested-files-/), stage: "draft", document: "d1", status: "open" })])
    expect(asked[0]!.question).toContain(`Your words "${PACK}" ask for the memo and "Release letter" in one file named "Kestrel Release Pack.docx", ` +
      'but the document plan delivers the memo and "Release letter" as separate files.')
    expect(existsSync(leanFile(directory, "plan.json"))).toBe(false)
    expect(readRequestedFiles(directory)).toBeUndefined()
    expect(partPrompts(model)).toEqual([])

    // Run again without an answer: still blocked on the same question, with no model call.
    const idle = scriptedModel({})
    expect((await runLean(directory, { sessions: idle.port })).status).toBe("blocked")
    expect(idle.prompts).toEqual([])
    expect(readQuestions(directory)).toEqual(asked)

    enqueueLeanRequest(directory, { kind: "answer", question_id: asked[0]!.id, answer: ANSWER })
    // The job may quote the lawyer's words differently on the new outline.
    const requoted: RequestedFile = { ...PACKED, quote: "all in one file named Kestrel Release Pack.docx" }
    const answered = scriptedModel(replies({ [AGENT.draft]: outliner(BODY), [AGENT.files]: filesOf(requoted) }))
    expect((await runLean(directory, { sessions: answered.port })).status).toBe("completed")
    // The outline is drawn up again with the answer; the mismatch already asked
    // about is neither revised against the answer nor asked again.
    expect(outlinePrompt(answered).text).toContain(ANSWER)
    expect(revisions(answered)).toEqual([])
    expect(readQuestions(directory)).toEqual([expect.objectContaining({ id: asked[0]!.id, status: "answered", answer: ANSWER })])
    const plan = readLeanPlan(directory)!
    expect(plan).toMatchObject({ basis: outlineBasis(directory), notes_file: "Kestrel Release Pack - cover memo.docx" })
    expect(plan.documents).toEqual([expect.objectContaining({ id: "d1", file: "Kestrel Release Pack.docx" })])
    expect(plan.documents[0]).not.toHaveProperty("deliverable")
    expect(readRequestedFiles(directory)).toMatchObject({ basis: plan.basis, files: [requoted] })

    // The file not delivered as asked is disclosed in the drafting notes.
    const items = requestedFilesItems(directory)
    expect(items).toEqual([expect.objectContaining({ stage: "finalize", kind: "omission", ref: expect.objectContaining({ repair_target: "memo" }) })])
    expect(items[0]!.detail).toContain(`the lawyer's words "${requoted.quote}" [source: request] ask for the memo and "Release letter" in one file named "Kestrel Release Pack.docx"`)
    expect(disclosableItems(directory)).toContainEqual(items[0])
    expect(readLeanText(directory, "drafting-notes.md")).toContain(`[ref: ${items[0]!.id}]`)
    expect(gateOf(directory)).toMatchObject({ status: "pass", blocking: [] })
    expect(gateOf(directory).disclosed).toContainEqual(expect.objectContaining({ id: items[0]!.id, status: "disclosed" }))
  }, 120_000)

  test("no memo and one document: the files job is never asked and the plan is the outline's", async () => {
    const request = "Prepare a release letter for Kestrel Works, in one file named Kestrel Release.docx."
    const single = outline(letter("File: Kestrel Release.docx"))
    const directory = prepared(request)
    const model = scriptedModel(replies({ [AGENT.draft]: outliner(single), [AGENT.files]: () => { throw new Error("The files job was asked.") } }))
    expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")

    expect(model.count(AGENT.files)).toBe(0)
    expect(filesSessions(model)).toEqual([])
    expect(readLeanText(directory, "plan.json")).toBe(planned(single, request, directory))
    expect(readRequestedFiles(directory)).toBeUndefined()
    expect(filesEvents(directory)).toEqual([])
  }, 120_000)

  test("separate files asked for: no revision, and the memo-only file names the notes", async () => {
    const request = "Prepare a release letter for Kestrel Works with a cover memo. Deliver the release letter as Kestrel Release.docx and the cover memo separately as Kestrel Advice.docx."
    const directory = prepared(request)
    const deliverables = registering()
    const files: RequestedFile[] = [
      { holds: ["d1"], name: "Kestrel Release.docx", quote: "Deliver the release letter as Kestrel Release.docx" },
      { holds: ["memo"], name: "Kestrel Advice.docx", quote: "the cover memo separately as Kestrel Advice.docx" },
    ]
    const model = scriptedModel(replies({ [AGENT.draft]: outliner(BODY), [AGENT.files]: filesOf(...files) }))
    expect((await runLean(directory, { sessions: model.port, registerDeliverable: deliverables.registerDeliverable })).status).toBe("completed")

    expect(revisions(model)).toEqual([])
    const plan = readLeanPlan(directory)!
    expect(plan).toMatchObject({ memo: "cover memo", notes_file: "Kestrel Advice.docx" })
    expect(plan.documents).toEqual([expect.objectContaining({ id: "d1", file: "Kestrel Release.docx" })])
    expect(plan.documents[0]).not.toHaveProperty("deliverable")
    expect(readRequestedFiles(directory)).toEqual({ schema: LEAN_REQUESTED_FILES_SCHEMA, basis: plan.basis!, stated: true, files })
    expect(requestedFilesItems(directory)).toEqual([])
    expect(readQuestions(directory)).toEqual([])
    expect(deliverables.names).toEqual(expect.arrayContaining(["Kestrel Release.docx", "Kestrel Advice.docx"]))
  }, 120_000)

  test("a files reply that never passes its check leaves no record and asks nothing; the plan is saved as before", async () => {
    const directory = prepared()
    // A record left from an earlier outline of the same inputs is not this plan's.
    writeRequestedFiles(directory, { schema: LEAN_REQUESTED_FILES_SCHEMA, basis: outlineBasis(directory), stated: true, files: [PACKED] })
    const model = scriptedModel(replies({ [AGENT.draft]: outliner(BODY), [AGENT.files]: () => "The lawyer asks for one file." }))
    expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")

    // Every allowed attempt in the one files session, then no record, revision or question.
    expect(filesSessions(model)).toHaveLength(1)
    expect(model.count(AGENT.files)).toBe(LEAN_MAX_RETRIES + 1)
    expect(new Set(model.prompts.filter(prompt => prompt.agent === AGENT.files).map(prompt => prompt.sessionID)).size).toBe(1)
    expect(revisions(model)).toEqual([])
    expect(readRequestedFiles(directory)).toBeUndefined()
    expect(requestedFilesItems(directory)).toEqual([])
    expect(filesEvents(directory)).toEqual([])
    expect(readQuestions(directory)).toEqual([])
    expect(readLeanText(directory, "plan.json")).toBe(planned(BODY, REQUEST, directory))
  }, 120_000)
})
