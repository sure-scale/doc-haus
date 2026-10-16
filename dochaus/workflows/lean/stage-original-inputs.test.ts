import { afterEach, expect, test } from "bun:test"
import { paragraphSha256 } from "../../lib/drafting-amendment"
import { packageReviewPrompt } from "./package-review"
import { writeLeanPlan } from "./plan"
import { stagePrompt } from "./prompts"
import { appliedChangeEvidence } from "./review-evidence"
import { stageInputSha } from "./runner"
import { readLeanJson, writeLeanText } from "./state"
import { cleanupMatters, matter } from "./test-support"

afterEach(cleanupMatters)

function deletionRecords() {
  const directory = matter()
  const plan = writeLeanPlan(directory, [
    {
      title: "Access note",
      purpose: "Record access instructions.",
      parts: [{ heading: "Access", scope: "Record access instructions." }],
    },
    {
      title: "Storage note",
      purpose: "Record storage instructions.",
      parts: [{ heading: "Storage", scope: "Record storage instructions." }],
    },
  ])
  for (const document of plan.documents) {
    const text = `Original ${document.title} wording, subject to the stated qualification.`
    writeLeanText(
      directory,
      `${document.id}/paragraphs.json`,
      JSON.stringify([{ pid: "p111111111111", sha256: paragraphSha256(text), text }]),
    )
    writeLeanText(directory, `${document.id}/paragraphs-applied.json`, JSON.stringify([]))
    writeLeanText(
      directory,
      `${document.id}/changes.json`,
      JSON.stringify({
        amendments: [
          {
            id: "deletion",
            pid: "p111111111111",
            existing_sha256: paragraphSha256(text),
            operation: "delete",
            updated: "",
            reason: "Remove this paragraph while retaining its explanation.",
            sources: [],
          },
        ],
        questions: [],
      }),
    )
    writeLeanText(
      directory,
      `${document.id}/apply.json`,
      JSON.stringify({
        applied: [{ id: "deletion", pid: "p111111111111", mode: "code", result_pids: [] }],
        unapplied: [],
      }),
    )
  }
  return { directory, documents: plan.documents }
}

// Literal record dependency test, not a claim that edited records were applied
// or accepted. Keep the stored hash unchanged to isolate the consumed text field.
test.each(["finalize", "package-review"] as const)(
  "original deletion wording in each planned body rekeys %s",
  (step) => {
    const fixture = deletionRecords()
    const prompt = () =>
      step === "finalize" ? stagePrompt(fixture.directory, step) : packageReviewPrompt(fixture.directory)
    for (const document of fixture.documents) {
      const key = stageInputSha(fixture.directory, step)
      const paragraphs = readLeanJson(fixture.directory, `${document.id}/paragraphs.json`) as {
        pid: string
        sha256: string
        text: string
      }[]
      const original = paragraphs[0]!.text
      expect(appliedChangeEvidence(fixture.directory, document.id)[0]!.deleted_text).toBe(original)
      expect(prompt()).toContain(JSON.stringify(original))

      const changed = `Revised ${document.title} wording, preserving the original qualification.`
      writeLeanText(
        fixture.directory,
        `${document.id}/paragraphs.json`,
        JSON.stringify([{ ...paragraphs[0]!, text: changed }]),
      )

      expect(appliedChangeEvidence(fixture.directory, document.id)[0]!.deleted_text).toBe(changed)
      expect(prompt()).toContain(JSON.stringify(changed))
      expect(prompt()).not.toContain(JSON.stringify(original))
      expect(stageInputSha(fixture.directory, step)).not.toBe(key)
    }
  },
)

test("notes and package review still exclude their own saved outputs from their resume keys", () => {
  const fixture = deletionRecords()
  const notesKey = stageInputSha(fixture.directory, "finalize")
  const reviewKeyBeforeNotes = stageInputSha(fixture.directory, "package-review")
  writeLeanText(
    fixture.directory,
    "drafting-notes.md",
    "# Notes\n\nThe deletion remains explained with its qualification.\n",
  )
  expect(stageInputSha(fixture.directory, "finalize")).toBe(notesKey)
  expect(stageInputSha(fixture.directory, "package-review")).not.toBe(reviewKeyBeforeNotes)
  const reviewKey = stageInputSha(fixture.directory, "package-review")
  writeLeanText(
    fixture.directory,
    "package-review.json",
    JSON.stringify({
      status: "pass",
      findings: [],
      limitations: [],
      delivery: [],
    }),
  )
  expect(stageInputSha(fixture.directory, "package-review")).toBe(reviewKey)
  expect(stageInputSha(fixture.directory, "finalize")).toBe(notesKey)
})
