import { afterEach, describe, expect, test } from "bun:test"
import { changePlanPrompt } from "./change-pipeline"
import { runDraftStage } from "./draft-sections"
import { packageReviewPrompt } from "./package-review"
import { writeLeanPlan, type DocID, type PlannedDocument, type StepID } from "./plan"
import { stagePrompt } from "./prompts"
import { stageInputs } from "./stages"
import { readLeanState, writeLeanText } from "./state"
import { AGENT, DRAFT, cleanupMatters, drafter, matter, scriptedModel, type Reply } from "./test-support"

// A requested memo is written separately, as the run's drafting notes, once the
// documents are final; a combined file is built at delivery, the memo first and
// then the clean body. Every body stage is told so from two flags (memo
// requested; this document is the combined body), the scope guards fire on it,
// and content-keyed resume sees it. Without a memo nothing changes.

afterEach(cleanupMatters)

const SEPARATE = "The request asks for a memo to accompany the documents. The run writes that memo separately, as the drafting notes, after the documents are final; it is not part of this document's body."
const COMBINED = "The request asks for that memo and this document in one file. The code builds that file at delivery: the memo first, then this document's clean body."
const KEEP_OUT = "Do not report the memo as missing from this body, and do not insert memo, cover-note or advice text meant for the memo into this body."
const KEEP_OWN = "Content this document's own purpose or the request assigns to it still belongs in this body. Advice belongs in this body only when this document's purpose is advice, and a reference to the memo only when the request requires one."
const BLOCK = `<package_composition>\n${SEPARATE}\n${KEEP_OUT}\n${KEEP_OWN}\n</package_composition>`
const COMBINED_BLOCK = `<package_composition>\n${SEPARATE}\n${COMBINED}\n${KEEP_OUT}\n${KEEP_OWN}\n</package_composition>`
const REVIEW_MEMO = "- Review only this document. Content the request assigns to the accompanying memo is not missing from this one; do not report it here."
const CHANGES_MEMO = "- Amend only this document. Content the request assigns to the accompanying memo belongs there; leave out a finding that asks for it here."
const REVIEW_BOTH = "- Review only this document. Content the request assigns to another planned document or the accompanying memo is not missing from this one; do not report it here."
const CHANGES_BOTH = "- Amend only this document. Content the request assigns to another planned document or the accompanying memo belongs there; leave out a finding that asks for it here."
const REVIEW_PLANNED = "- Review only this document. Content the request assigns to another planned document is not missing from this one; do not report it here."
const CHANGES_PLANNED = "- Amend only this document. Content the request assigns to another planned document belongs in that document; leave out a finding that asks for it here."
const PACKAGE_LINE = "The requested memo is the run-wide drafting notes, delivered separately from the document bodies; where a document and the memo share one file, the code places the memo before that body at delivery. Memo, cover-note or advice text meant for the memo that appears inside a document body is a body finding against that document. Content a document's own purpose or the request assigns to that document still belongs in its body. Advice belongs in a body only when that document's purpose is advice, and a reference to the memo only when the request requires one."

const NOTE: Omit<PlannedDocument, "id"> = { title: "Supply note", purpose: "Operative terms", parts: [{ heading: "Delivery", scope: "Delivery terms" }] }
const SCHEDULE: Omit<PlannedDocument, "id"> = { title: "Delivery schedule", purpose: "Delivery dates", parts: [{ heading: "Dates", scope: "The delivery dates" }] }
const COMBINED_NOTE: Omit<PlannedDocument, "id"> = { ...NOTE, file: "supply-pack.docx", deliverable: "combined" }
const MEMO = { name: "cover memo", file: "supply-pack - cover memo.docx" }

const REVIEWS = ["review", "review-independent", "review-wording", "review-client-disadvantage"] as const
const BODY_STAGES = ["draft", ...REVIEWS, "change-plan", "change-prepare", "changes", "change-review"] as const

/** The prompt each body stage of this document is sent, with its context.
 * change-prepare and change-review open with the context-only stage prompt. */
function bodyPrompts(directory: string, document: DocID): Record<(typeof BODY_STAGES)[number], string> {
  return {
    draft: stagePrompt(directory, `${document}.draft`),
    ...Object.fromEntries(REVIEWS.map(stage => [stage, stagePrompt(directory, `${document}.${stage}`)])) as Record<(typeof REVIEWS)[number], string>,
    "change-plan": changePlanPrompt(directory, document),
    "change-prepare": stagePrompt(directory, `${document}.change-prepare`, { contextOnly: true }),
    changes: stagePrompt(directory, `${document}.changes`),
    "change-review": stagePrompt(directory, `${document}.change-review`, { contextOnly: true }),
  }
}
const occurrences = (text: string, part: string) => text.split(part).length - 1
const composition = (directory: string, step: StepID) => stageInputs(directory, step).filter(row => row.name === "package-composition")

describe("a requested memo in the body stages' context", () => {
  test("a memo with one body document: every body stage is told the memo is separate, and the scope guards fire", () => {
    const directory = matter()
    writeLeanPlan(directory, [NOTE], undefined, undefined, MEMO)
    for (const [stage, prompt] of Object.entries(bodyPrompts(directory, "d1"))) {
      expect([stage, occurrences(prompt, "<package_composition>")]).toEqual([stage, 1])
      expect(prompt).toContain(BLOCK)
      expect(prompt).not.toContain(COMBINED)
    }
    for (const stage of REVIEWS) expect(stagePrompt(directory, `d1.${stage}`)).toContain(REVIEW_MEMO)
    expect(stagePrompt(directory, "d1.changes")).toContain(CHANGES_MEMO)
  })

  test("only the combined document is told the code places the memo before its body", () => {
    const directory = matter()
    writeLeanPlan(directory, [SCHEDULE, COMBINED_NOTE], undefined, undefined, MEMO)
    for (const prompt of Object.values(bodyPrompts(directory, "d2"))) expect(prompt).toContain(COMBINED_BLOCK)
    for (const prompt of Object.values(bodyPrompts(directory, "d1"))) {
      expect(prompt).toContain(BLOCK)
      expect(prompt).not.toContain(COMBINED)
    }
    for (const document of ["d1", "d2"] as const) {
      for (const stage of REVIEWS) expect(stagePrompt(directory, `${document}.${stage}`)).toContain(REVIEW_BOTH)
      expect(stagePrompt(directory, `${document}.changes`)).toContain(CHANGES_BOTH)
    }
    // Without the memo a combined deliverable is no combined file.
    writeLeanPlan(directory, [SCHEDULE, COMBINED_NOTE])
    for (const prompt of Object.values(bodyPrompts(directory, "d2"))) expect(prompt).not.toContain("package_composition")
  })

  test("without a memo the block is absent and the guards keep their wording", () => {
    const directory = matter()
    writeLeanPlan(directory, [NOTE])
    for (const prompt of Object.values(bodyPrompts(directory, "d1"))) {
      expect(prompt).not.toContain("package_composition")
      expect(prompt).not.toContain("accompanying memo")
      expect(prompt).not.toContain("Review only this document.")
      expect(prompt).not.toContain("Amend only this document.")
    }
    writeLeanPlan(directory, [SCHEDULE, NOTE])
    for (const document of ["d1", "d2"] as const) {
      for (const stage of REVIEWS) expect(stagePrompt(directory, `${document}.${stage}`)).toContain(REVIEW_PLANNED)
      expect(stagePrompt(directory, `${document}.changes`)).toContain(CHANGES_PLANNED)
      expect(stagePrompt(directory, `${document}.review`)).not.toContain("package_composition")
    }
  })

  test("a no-memo prompt is the memo prompt without the block and its guard", () => {
    const directory = matter()
    writeLeanPlan(directory, [NOTE], undefined, undefined, MEMO)
    const withMemo = bodyPrompts(directory, "d1")
    writeLeanPlan(directory, [NOTE])
    const withoutMemo = bodyPrompts(directory, "d1")
    for (const stage of BODY_STAGES)
      expect([stage, withMemo[stage].replace(`\n\n${BLOCK}`, "").replace(`\n${REVIEW_MEMO}`, "").replace(`\n${CHANGES_MEMO}`, "")]).toEqual([stage, withoutMemo[stage]])
  })

  test("stage inputs carry both flags for every body stage, and only when a memo is requested", () => {
    const directory = matter()
    writeLeanPlan(directory, [SCHEDULE, COMBINED_NOTE], undefined, undefined, MEMO)
    const steps = (["d1", "d2"] as const).flatMap(document => BODY_STAGES.map(stage => `${document}.${stage}` as StepID))
    const withMemo = new Map(steps.map(step => [step, stageInputs(directory, step)]))
    for (const step of steps)
      expect([step, composition(directory, step)]).toEqual([step, [{ name: "package-composition", bytes: step.startsWith("d2.") ? COMBINED_BLOCK : BLOCK }]])
    for (const step of ["d1.apply", "d2.apply", "finalize", "package-review", "gate"] as StepID[]) expect([step, composition(directory, step)]).toEqual([step, []])

    // The combined flag is its own input: the same memo with d2 as a plain body rekeys d2's stages only.
    writeLeanPlan(directory, [SCHEDULE, NOTE], undefined, undefined, MEMO)
    for (const step of steps) expect([step, composition(directory, step)]).toEqual([step, [{ name: "package-composition", bytes: BLOCK }]])

    writeLeanPlan(directory, [SCHEDULE, COMBINED_NOTE])
    for (const step of steps)
      expect([step, stageInputs(directory, step)]).toEqual([step, withMemo.get(step)!.filter(row => row.name !== "package-composition")])
  })

  test("package review counts memo text inside a body as a body finding only when a memo is requested", () => {
    const directory = matter()
    writeLeanPlan(directory, [NOTE], undefined, undefined, MEMO)
    expect(occurrences(packageReviewPrompt(directory), PACKAGE_LINE)).toBe(1)
    writeLeanPlan(directory, [NOTE])
    expect(packageReviewPrompt(directory)).not.toContain("The requested memo is the run-wide drafting notes")
  })

  test("the first part of each drafted body carries the block, also in the outline's session", async () => {
    const directory = matter()
    writeLeanText(directory, "request.md", "Draft a delivery schedule, a supply note from the supply terms and a cover memo, and send the note and memo together as supply-pack.docx.")
    const draft = drafter("# Delivery schedule\n\n## Dates\n\nThe delivery dates.\n", DRAFT)
    const outlined: Reply = (prompt, index) => {
      const reply = draft(prompt, index)
      return reply.startsWith("Acting for:")
        ? reply.replace("Acting for: not stated", "Acting for: not stated\nMemo: cover memo").replace("# Supply note", "# Supply note\nFile: supply-pack.docx\nDeliverable: combined")
        : reply
    }
    const model = scriptedModel({ [AGENT.draft]: outlined })
    const deps = { sessions: model.port }
    expect((await runDraftStage(directory, readLeanState(directory)!, deps, () => {}, "d1", () => {})).status).toBe("done")
    expect((await runDraftStage(directory, readLeanState(directory)!, deps, () => {}, "d2", () => {})).status).toBe("done")
    const first = (title: string) => model.prompts.find(prompt => new RegExp(`^Draft part 1 of \\d+ of "${title}": `, "m").test(prompt.text))?.text ?? ""
    const schedule = first("Delivery schedule")
    expect(schedule).toContain(BLOCK)
    expect(schedule).not.toContain(COMBINED)
    expect(first("Supply note")).toContain(COMBINED_BLOCK)
  }, 60_000)
})
