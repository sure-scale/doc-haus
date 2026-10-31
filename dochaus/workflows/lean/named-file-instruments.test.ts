import { afterEach, describe, expect, test } from "bun:test"
import { readFileSync } from "node:fs"
import { strFromU8, unzipSync } from "fflate"
import { renderMarkdownToDocx } from "../../lib/render"
import { presentDraft } from "./draft-presentation"
import { coverageScopeQuestion, planCoverageCurrent, separateFileChosen, type CoverageAct } from "./package-coverage"
import { readLeanPlan, type PlannedDocument } from "./plan"
import { foldRequestedFiles, type RequestedFile } from "./requested-files"
import { runLean } from "./runner"
import { leanFile, readQuestions, writeLeanText, type Question } from "./state"
import { AGENT, NO_FINDINGS, cleanupMatters, coverageJobs, draftingNotes, matter, noChanges, scriptedModel, type Reply } from "./test-support"

// A request that names one file, and a plan that needs a second instrument
// executed by other parties: the second instrument is drafted as a separate
// instrument inside the named file, under its own title on a new page, unless
// the lawyer chose a file of its own for it. Synthetic parties and files only;
// only the model is scripted.

afterEach(cleanupMatters)

const SOURCE_TEXT = "Kestrel Works shall release the deposit once the Orchard Committee approves the revised schedule."
const REQUEST = "Prepare the papers so that the Orchard Committee approves the revised schedule and Kestrel Works then releases the deposit, all in one file named Kestrel Release Pack.docx."
const PACK = "all in one file named Kestrel Release Pack.docx"
const APPROVE = "the Orchard Committee approves the revised schedule"
const PACKED: RequestedFile = { holds: ["d1", "d2"], name: "Kestrel Release Pack.docx", quote: PACK }

const TWO = "Acting for: not stated\n# Committee resolution\nForm: other\nExecuted by: the Orchard Committee\n## Terms\nThe approval.\n" +
  "# Release letter\nForm: letter\nExecuted by: Kestrel Works\n## Terms\nThe release."
const act: CoverageAct = { act: "approve the revised schedule", source: "request", quote: APPROVE, parties: ["the Orchard Committee"], document: "d1" }

const doc = (title: string, executed: string, extra: Partial<PlannedDocument> = {}): Omit<PlannedDocument, "id"> =>
  ({ title, purpose: `The ${title}.`, parts: [{ heading: "Terms", scope: "The terms." }], executed_by: [executed], ...extra })
const RESOLUTION = doc("Committee resolution", "the Orchard Committee", { form: "other" })
const LETTER = doc("Release letter", "Kestrel Works", { form: "letter" })

describe("one named file holding two planned documents", () => {
  test("is held as one document whose second instrument keeps its own title, parties and form", () => {
    const folded = foldRequestedFiles([RESOLUTION, LETTER], { stated: true, files: [PACKED] })
    expect(folded.documents).toHaveLength(1)
    expect(folded.documents[0]).toMatchObject({ title: "Committee resolution", executed_by: ["the Orchard Committee", "Kestrel Works"] })
    expect(folded.documents[0]!.parts).toEqual([
      { heading: "Terms", scope: "The terms." },
      { heading: "Terms (Release letter)", display_heading: "Terms", scope: "The terms.", instrument: { title: "Release letter", purpose: "The Release letter.", executed_by: ["Kestrel Works"], form: "letter" } },
    ])
    expect(folded.requested.files).toEqual([{ ...PACKED, holds: ["d1"] }])
    expect(folded.at).toEqual([0, 0])
  })

  test("is left as planned when no file is named, the lawyer chose a file of its own, or one is drafted against a template", () => {
    const unchanged = (documents: Omit<PlannedDocument, "id">[], ...args: [Parameters<typeof foldRequestedFiles>[1], boolean?]) =>
      expect(foldRequestedFiles(documents, ...args)).toMatchObject({ documents, applied: [] })
    unchanged([RESOLUTION, LETTER], { stated: false, files: [] })
    unchanged([RESOLUTION, LETTER], { stated: true, files: [{ ...PACKED, holds: ["d1"] }, { ...PACKED, holds: ["d2"], name: "Other.docx" }] })
    unchanged([RESOLUTION, LETTER], { stated: true, files: [PACKED] }, true)
    unchanged([RESOLUTION, { ...LETTER, template: "Letter template.docx" }], { stated: true, files: [PACKED] })
    unchanged([RESOLUTION, { ...LETTER, playbook_review: { sources: [{ source_id: "s1", quote: "Releases follow approval." }], reason: "The firm's release playbook." } }], { stated: true, files: [PACKED] })
  })

  test("folds when the outline records that no playbook applies", () => {
    const none = { sources: [], reason: "No playbook supplied." }
    expect(foldRequestedFiles([{ ...RESOLUTION, playbook_review: none }, { ...LETTER, playbook_review: none }], { stated: true, files: [PACKED] }).documents).toHaveLength(1)
  })
})

describe("the plan-scope question when the request names one file", () => {
  const named = [{ ...doc("Release letter", "Kestrel Works"), id: "d1" as const, file: "Kestrel Release Pack.docx" }]
  test("offers the named file first, then a file of its own; with no named file or a saved plan it offers no choices", () => {
    expect(coverageScopeQuestion([act], named)!.options?.map(option => option.text.split(":")[0])).toEqual(['In "Kestrel Release Pack.docx", the file the request names', "In a file of its own", "Outside this request."])
    expect(coverageScopeQuestion([act], [{ ...named[0]!, file: undefined }])!.options).toBeUndefined()
    expect(coverageScopeQuestion([act], named, [], { saved: true })!.options).toBeUndefined()
  })

  test("a file of its own, once answered, keeps the documents in separate files", () => {
    const asked = coverageScopeQuestion([act], named)!
    const answered = (option: string): Question[] => [{ ...asked, status: "answered", option_id: option, answer: "chosen" }]
    expect(separateFileChosen(answered("o2"))).toBe(true)
    expect(separateFileChosen(answered("o1"))).toBe(false)
    expect(separateFileChosen([asked])).toBe(false)
  })
})

test("a later title starts a new page; a document with one title gains no page break", async () => {
  const breaks = async (markdown: string) => {
    const { bytes } = presentDraft((await renderMarkdownToDocx(markdown)).bytes, [{ heading: "Terms" }])
    return strFromU8(unzipSync(bytes)["word/document.xml"]!).match(/<w:pageBreakBefore\/>/g)?.length ?? 0
  }
  expect(await breaks("# Committee resolution\n\n## Terms\n\nThe approval.\n")).toBe(0)
  expect(await breaks("# Committee resolution\n\n## Terms\n\nThe approval.\n\n# Release letter\n\n## Terms\n\nThe release.\n")).toBe(1)
})

test("through the run, the named file is one body holding both instruments, and the coverage check holds on it", async () => {
  const directory = matter(SOURCE_TEXT)
  writeLeanText(directory, "request.md", REQUEST)
  const outliner: Reply = text => /^Draft part \d+ of \d+ of "/m.test(text) ? "The terms are recorded." : TWO
  const model = scriptedModel({
    [AGENT.draft]: outliner,
    [AGENT.files]: () => JSON.stringify({ stated: true, files: [PACKED] }),
    ...coverageJobs(() => JSON.stringify({ acts: [act] })),
    [AGENT.review]: () => NO_FINDINGS,
    [AGENT.changes]: noChanges,
    [AGENT.finalize]: draftingNotes,
  })
  expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")

  const plan = readLeanPlan(directory)!
  expect(plan.documents).toEqual([expect.objectContaining({ id: "d1", title: "Committee resolution", file: "Kestrel Release Pack.docx", executed_by: ["the Orchard Committee", "Kestrel Works"] })])
  expect(plan.documents[0]!.parts.map(part => part.instrument?.title)).toEqual([undefined, "Release letter"])
  expect(planCoverageCurrent(directory)).toBe(true)
  expect(readQuestions(directory)).toEqual([])

  // The second instrument is told it is one, and the body prints its title on a new page.
  const prompts = model.prompts.filter(prompt => /^Draft part \d+ of /m.test(prompt.text))
  expect(prompts.map(prompt => prompt.text.includes('in "Release letter", a separate instrument in the same file'))).toEqual([false, true])
  const body = strFromU8(unzipSync(readFileSync(leanFile(directory, "d1/body.docx")))["word/document.xml"]!)
  expect(body).toContain("Release letter")
  expect(body.match(/<w:pageBreakBefore\/>/g)).toHaveLength(1)
}, 120_000)
