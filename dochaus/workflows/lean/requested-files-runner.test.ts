import { afterEach, describe, expect, test } from "bun:test"
import { existsSync } from "node:fs"
import { combinedDelivery, memoNotesFile } from "./deliverables"
import { PLAN_FILES_REVISIONS } from "./draft-sections"
import { readLeanEvents } from "./events"
import type { GateResult } from "./gate"
import { parsePlan, readLeanPlan, writeLeanPlan } from "./plan"
import { PACKAGE_SCOPE_QUESTION } from "./package-scope"
import { enqueueLeanRequest } from "./requests"
import { FILES_REVISION_SCOPE, LEAN_REQUESTED_FILES_SCHEMA, readRequestedFiles, requestedFilesItems, writeRequestedFiles, type RequestedFile } from "./requested-files"
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
/** The planned documents and memo every outline here gives. */
const PLANNED = { titles: ["Release letter"], memo: "cover memo" }

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
const coverageSessions = (model: Model) => model.sessions.filter(session => session.agent === AGENT.coverage)
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
    // The revision changes only the delivery, so the coverage job is not asked again either.
    expect(coverageSessions(model)).toHaveLength(1)
    expect(revisions(model)).toEqual([expect.objectContaining({ sessionID: outlinePrompt(model).sessionID })])
    expect(revisions(model)[0]!.text.split("\n").filter(line => line.startsWith("- "))).toEqual([
      `- ${FILES_REVISION_SCOPE}`,
      `- The lawyer's words "${PACK}" ask for the memo and "Release letter" in one file. Write "Deliverable: combined" under "Release letter".`,
    ])

    const plan = readLeanPlan(directory)!
    expect(plan).toMatchObject({ basis: outlineBasis(directory), memo: "cover memo", notes_file: "Kestrel Release Pack - cover memo.docx" })
    expect(plan.documents).toEqual([expect.objectContaining({ id: "d1", title: "Release letter", deliverable: "combined", file: "Kestrel Release Pack.docx" })])
    expect(readRequestedFiles(directory)).toEqual({ schema: LEAN_REQUESTED_FILES_SCHEMA, basis: plan.basis!, planned: PLANNED, stated: true, files: [PACKED] })
    expect(filesEvents(directory)).toEqual([expect.objectContaining({ stage: "draft", document: "d1", step: "files", outcome: "stated", issues: [],
      reason: '"Release letter": file "Kestrel Release Pack.docx"' })])

    // Delivered as asked: one combined file under the lawyer's name, nothing to disclose or ask.
    expect(combinedDelivery(directory)).toEqual({ document: "d1", name: "Kestrel Release Pack" })
    expect(requestedFilesItems(directory)).toEqual([])
    expect(readQuestions(directory)).toEqual([])
    expect(gateOf(directory)).toMatchObject({ status: "pass", disclosed: [], blocking: [] })
  }, 120_000)

  test("a files revision that also renames a document is not adopted: it is spent, and the next revision starts from the earlier outline", async () => {
    const directory = prepared()
    const renamed = outline("Memo: cover memo", "# Release pack letter", "Deliverable: combined", "## Release", "The release.")
    // The first revision renames the document as it combines it; the second only combines it.
    const draft: Reply = text => !text.startsWith(REVISE) ? outliner(BODY)(text, 0) : revisions(model).length === 1 ? renamed : COMBINED
    const model = scriptedModel(replies({ [AGENT.draft]: draft, [AGENT.files]: filesOf(PACKED) }))
    expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")

    const issue = `- The lawyer's words "${PACK}" ask for the memo and "Release letter" in one file. Write "Deliverable: combined" under "Release letter".`
    expect(revisions(model).map(prompt => prompt.text.split("\n").filter(line => line.startsWith("- ")))).toEqual([[`- ${FILES_REVISION_SCOPE}`, issue], [`- ${FILES_REVISION_SCOPE}`, issue]])
    expect(readLeanPlan(directory)!.documents).toEqual([expect.objectContaining({ id: "d1", title: "Release letter", deliverable: "combined", file: "Kestrel Release Pack.docx" })])
    expect(coverageSessions(model)).toHaveLength(1)
    expect(filesSessions(model)).toHaveLength(1)
    expect(readRequestedFiles(directory)).toMatchObject({ planned: PLANNED, files: [PACKED] })
  }, 120_000)

  test("from a files revision only the delivery is adopted: a part it also changes is not, and neither job is asked again", async () => {
    const directory = prepared()
    const reworked = outline("Memo: cover memo", "# Release letter", "Deliverable: combined", "## Release terms", "The release terms.")
    const model = scriptedModel(replies({ [AGENT.draft]: outliner(BODY, reworked), [AGENT.files]: filesOf(PACKED) }))
    expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")

    expect(revisions(model)).toHaveLength(1)
    const plan = readLeanPlan(directory)!
    expect(plan.documents).toEqual([expect.objectContaining({ id: "d1", title: "Release letter", deliverable: "combined", file: "Kestrel Release Pack.docx" })])
    expect(plan.documents[0]!.parts.map(part => part.heading)).toEqual(["Release"])
    expect(coverageSessions(model)).toHaveLength(1)
    expect(filesSessions(model)).toHaveLength(1)
  }, 120_000)

  test("a files revision whose titles differ only in case and spacing is adopted, keeping the plan's spelling", async () => {
    const directory = prepared()
    const respaced = outline("Memo: cover memo", "# release  LETTER", "Deliverable: combined", "## Release", "The release.")
    const model = scriptedModel(replies({ [AGENT.draft]: outliner(BODY, respaced), [AGENT.files]: filesOf(PACKED) }))
    expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")

    expect(revisions(model)).toHaveLength(1)
    expect(readLeanPlan(directory)!.documents).toEqual([expect.objectContaining({ id: "d1", title: "Release letter", deliverable: "combined", file: "Kestrel Release Pack.docx" })])
    expect(readRequestedFiles(directory)).toMatchObject({ planned: PLANNED, files: [PACKED] })
  }, 120_000)

  test("a files revision's delivery is taken only for the documents its issues concern", async () => {
    const directory = prepared(`${REQUEST} Also prepare a deposit notice for the Orchard Committee.`)
    const notice = ["# Deposit notice", "Deliverable: redline", "## Deposit", "The deposit."]
    // The revision combines the letter and leaves out the notice's Deliverable line.
    const revised = outline("Memo: cover memo", letter("Deliverable: combined"), ...notice.filter(line => !line.startsWith("Deliverable")))
    const model = scriptedModel(replies({ [AGENT.draft]: outliner(outline("Memo: cover memo", letter(), ...notice), revised), [AGENT.files]: filesOf(PACKED) }))
    await runLean(directory, { sessions: model.port })

    expect(revisions(model)).toHaveLength(1)
    expect(readLeanPlan(directory)!.documents).toEqual([
      expect.objectContaining({ id: "d1", title: "Release letter", deliverable: "combined", file: "Kestrel Release Pack.docx" }),
      expect.objectContaining({ id: "d2", title: "Deposit notice", deliverable: "redline" }),
    ])
  }, 120_000)

  test("the memo asked for with a document planned as a redline goes to the lawyer with no revision", async () => {
    const directory = prepared()
    const model = scriptedModel(replies({ [AGENT.draft]: outliner(outline("Memo: cover memo", letter("Deliverable: redline")), COMBINED), [AGENT.files]: filesOf(PACKED) }))
    expect((await runLean(directory, { sessions: model.port })).status).toBe("blocked")

    expect(revisions(model)).toEqual([])
    const asked = readQuestions(directory)
    expect(asked).toEqual([expect.objectContaining({ id: expect.stringMatching(/^requested-files-/), stage: "draft", status: "open" })])
    expect(asked[0]!.question).toContain(`Your words "${PACK}" ask for the memo and "Release letter" in one file named "Kestrel Release Pack.docx", ` +
      'but the document plan delivers "Release letter" as a redline and the memo as a file of its own.')
    expect(existsSync(leanFile(directory, "plan.json"))).toBe(false)
  }, 120_000)

  test("a coverage scope question stops the step before the files job is asked", async () => {
    const request = "Prepare the papers so that the Orchard Committee approves the revised schedule, and a release letter for Kestrel Works with a cover memo."
    const directory = prepared(request)
    const act = { act: "approve the revised schedule", source: "request", quote: "the Orchard Committee approves the revised schedule", parties: ["the Orchard Committee"], document: "d1" }
    const misplaced = outline("Memo: cover memo", letter("Executed by: Kestrel Works"))
    const model = scriptedModel(replies({ [AGENT.draft]: outliner(misplaced), [AGENT.coverage]: () => JSON.stringify({ acts: [act] }) }))
    const state = await runLean(directory, { sessions: model.port })

    expect(state.status).toBe("blocked")
    expect(state.stages["d1.draft"]).toMatchObject({ status: "blocked", outputs: ["coverage.json"] })
    expect(readQuestions(directory)).toEqual([expect.objectContaining({ id: expect.stringMatching(new RegExp(`^${PACKAGE_SCOPE_QUESTION}`)), stage: "draft", status: "open" })])
    expect(filesSessions(model)).toEqual([])
    expect(model.count(AGENT.files)).toBe(0)
    expect(filesEvents(directory)).toEqual([])
    expect(existsSync(leanFile(directory, "plan.json"))).toBe(false)
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
    expect(readRequestedFiles(directory)).toMatchObject({ basis: plan.basis, planned: PLANNED, files: [requoted] })

    // The file not delivered as asked is disclosed in the drafting notes, citing the answer.
    const items = requestedFilesItems(directory)
    expect(items).toEqual([expect.objectContaining({ stage: "finalize", kind: "omission", ref: expect.objectContaining({ repair_target: "memo" }) })])
    expect(items[0]!.detail).toBe(`Requested files: the lawyer's words "${requoted.quote}" [source: request] ask for the memo and "Release letter" in one file named "Kestrel Release Pack.docx", ` +
      'but the run delivers the memo in "Kestrel Release Pack - cover memo.docx" and "Release letter" in "Kestrel Release Pack.docx". ' +
      `The lawyer answered the requested-files question ${asked[0]!.id}; the plan drawn up with that answer still delivers them so.`)
    expect(disclosableItems(directory)).toContainEqual(items[0])
    expect(readLeanText(directory, "drafting-notes.md")).toContain(`[ref: ${items[0]!.id}]`)
    expect(gateOf(directory)).toMatchObject({ status: "pass", blocking: [] })
    expect(gateOf(directory).disclosed).toContainEqual(expect.objectContaining({ id: items[0]!.id, status: "disclosed" }))

    // A later update that leaves these words as they were: the plan is drawn up
    // again, with no revision and no new question, and the item still cites the answer.
    enqueueLeanRequest(directory, { kind: "update", text: "Address the release letter to the Orchard Committee." })
    const updated = scriptedModel(replies({ [AGENT.draft]: outliner(BODY), [AGENT.files]: filesOf(requoted) }))
    expect((await runLean(directory, { sessions: updated.port })).status).toBe("completed")
    expect(outlinePrompt(updated).text).toContain("Address the release letter to the Orchard Committee.")
    expect(updated.count(AGENT.files)).toBe(1)
    expect(revisions(updated)).toEqual([])
    expect(readQuestions(directory)).toEqual([expect.objectContaining({ id: asked[0]!.id, status: "answered", answer: ANSWER })])
    expect(requestedFilesItems(directory).map(item => item.detail)).toEqual([items[0]!.detail])
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
    expect(readRequestedFiles(directory)).toEqual({ schema: LEAN_REQUESTED_FILES_SCHEMA, basis: plan.basis!, planned: PLANNED, stated: true, files })
    expect(requestedFilesItems(directory)).toEqual([])
    expect(readQuestions(directory)).toEqual([])
    expect(deliverables.names).toEqual(expect.arrayContaining(["Kestrel Release.docx", "Kestrel Advice.docx"]))
  }, 120_000)

  test("a files reply that never passes its check is recorded as unchecked and disclosed, and asks nothing; the plan is saved as before", async () => {
    const directory = prepared()
    // A record left from an earlier outline of the same inputs is not this plan's.
    writeRequestedFiles(directory, { schema: LEAN_REQUESTED_FILES_SCHEMA, basis: outlineBasis(directory), planned: PLANNED, stated: true, files: [PACKED] })
    const model = scriptedModel(replies({ [AGENT.draft]: outliner(BODY), [AGENT.files]: () => "The lawyer asks for one file." }))
    expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")

    // Every allowed attempt in the one files session, then no revision or question.
    expect(filesSessions(model)).toHaveLength(1)
    expect(model.count(AGENT.files)).toBe(LEAN_MAX_RETRIES + 1)
    expect(new Set(model.prompts.filter(prompt => prompt.agent === AGENT.files).map(prompt => prompt.sessionID)).size).toBe(1)
    expect(revisions(model)).toEqual([])
    const plan = readLeanPlan(directory)!
    expect(readRequestedFiles(directory)).toEqual({ schema: LEAN_REQUESTED_FILES_SCHEMA, basis: plan.basis!, planned: PLANNED, status: "unchecked", stated: false, files: [] })
    expect(filesEvents(directory)).toEqual([expect.objectContaining({ stage: "draft", document: "d1", step: "files", outcome: "unchecked",
      reason: "The files job's reply did not pass its check, so the files the lawyer asks to receive could not be confirmed; the plan delivers its files as outlined and the drafting notes say so." })])
    expect(filesEvents(directory)[0]).not.toHaveProperty("issues")
    expect(readQuestions(directory)).toEqual([])
    expect(readLeanText(directory, "plan.json")).toBe(planned(BODY, REQUEST, directory))

    // The lawyer is told the files could not be confirmed.
    const items = requestedFilesItems(directory)
    expect(items).toEqual([expect.objectContaining({ stage: "finalize", kind: "omission", ref: expect.objectContaining({ repair_target: "memo" }) })])
    expect(items[0]!.detail).toStartWith("Requested files could not be checked:")
    expect(readLeanText(directory, "drafting-notes.md")).toContain(`[ref: ${items[0]!.id}]`)
    expect(gateOf(directory)).toMatchObject({ status: "pass", blocking: [] })
    expect(gateOf(directory).disclosed).toContainEqual(expect.objectContaining({ id: items[0]!.id, status: "disclosed" }))
  }, 120_000)

  test("a name refused on one attempt and dropped on the next is still disclosed", async () => {
    const directory = prepared()
    const refused = "Kestrel Release Bundle.docx"
    const files: Reply = (_text, index) => JSON.stringify({ stated: true, files: [{ ...PACKED, name: index === 0 ? refused : null }] })
    const model = scriptedModel(replies({ [AGENT.draft]: outliner(COMBINED), [AGENT.files]: files }))
    expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")

    expect(model.count(AGENT.files)).toBe(2)
    expect(readRequestedFiles(directory)).toMatchObject({ stated: true, files: [{ ...PACKED, name: null, refused_name: refused }] })
    const items = requestedFilesItems(directory)
    expect(items).toHaveLength(1)
    expect(items[0]!.detail).toBe(`Requested files: the lawyer's words "${PACK}" [source: request] ask for the memo and "Release letter" in one file and name this file; ` +
      `the run could not confirm the name "${refused}" in those words and delivers the memo and "Release letter" in one file named "Release letter.docx". ` +
      "The lawyer should confirm how these files are to be delivered.")
    expect(gateOf(directory).disclosed).toContainEqual(expect.objectContaining({ id: items[0]!.id, status: "disclosed" }))
  }, 120_000)
})
