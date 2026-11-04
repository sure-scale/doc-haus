import { afterEach, describe, expect, test } from "bun:test"
import { strFromU8, unzipSync } from "fflate"
import { renderMarkdownToDocx } from "../../lib/render"
import { consistencyFindings, type Para } from "./consistency"
import { draftPresentation } from "./draft-sections"
import { presentDraft } from "./draft-presentation"
import { parsePlan } from "./plan"
import { stagePrompt } from "./prompts"
import { cleanupMatters, matter } from "./test-support"

// An attachment the text refers to is planned as a trailing part of its
// document, headed as the attachment is headed. The plan marks it, the Word
// body starts it on a new page, and the consistency check holds every form to
// the attachments its text refers to. Fixtures are neutral and invented.

afterEach(cleanupMatters)

const outline = (...parts: string[]) => ["Acting for: not stated", "# Board Consent", "What it is for.", "Form: other", ...parts].join("\n")
const doc = (lines: string[]): Para[] => lines.map((text, index) => ({ pid: `p${index + 1}`, text }))

describe("attachment parts", () => {
  test("parsePlan marks a part headed as an attachment, and its heading prints", () => {
    const parsed = parsePlan(outline("## Resolutions", "The resolutions adopted.", "## Exhibit A – Form of Notice", "The notice the resolutions approve."))
    expect(parsed.issues).toEqual([])
    expect(parsed.value!.documents[0]!.parts).toEqual([
      { heading: "Resolutions", scope: "The resolutions adopted.", display_heading: "" },
      { heading: "Exhibit A – Form of Notice", scope: "The notice the resolutions approve.", attachment: true },
    ])
  })

  test("a heading that only mentions an attachment, or names none, is left unflagged", () => {
    const parsed = parsePlan(outline("## Schedule of meetings", "When the board meets.", "## Exhibits referred to", "The approvals.", "## Signatures", "Who signs."))
    expect(parsed.issues).toEqual([])
    expect(parsed.value!.documents[0]!.parts.some(part => part.attachment)).toBe(false)
  })

  test("the Word body starts an attachment part on a new page, and no other part", async () => {
    const plan = { title: "Board Consent", form: "other" as const, parts: [
      { heading: "Resolutions", scope: "", display_heading: "" },
      { heading: "Exhibit A – Form of Notice", scope: "", attachment: true as const },
    ] }
    const presentation = draftPresentation(plan, ["## Resolutions\n\nThe board approves the notice in Exhibit A.", "## Exhibit A – Form of Notice\n\nNotice text."])!
    expect(presentation.parts.map(part => part.attachment)).toEqual([undefined, true])
    const rendered = await renderMarkdownToDocx(presentation.source)
    const body = strFromU8(unzipSync(presentDraft(rendered.bytes, presentation.parts).bytes)["word/document.xml"]!)
    const paragraphs = body.split("</w:p>")
    const breaks = paragraphs.filter(paragraph => paragraph.includes("<w:pageBreakBefore/>"))
    expect(breaks).toHaveLength(1)
    expect(breaks[0]).toContain("Exhibit A – Form of Notice")
  })

  test("a document with no clause numbers that refers to a missing exhibit is reported, whatever its form", () => {
    const text = doc(["Board Consent", "The board approves the form of notice attached as Exhibit A.", "Signed by each director."])
    for (const form of ["other", "agreement", "letter", "advice"])
      expect(consistencyFindings({ title: "Board Consent", form, final: text, baselines: [], others: [] }).map(row => row.detail)).toEqual([
        'The paragraph beginning "The board approves the form of notice attached as Exhibit A." refers to Exhibit A, but the final text has no Exhibit A.',
      ])
    const attached = doc(["Board Consent", "The board approves the form of notice attached as Exhibit A.", "Signed by each director.", "Exhibit A – Form of Notice", "Notice text."])
    expect(consistencyFindings({ title: "Board Consent", form: "other", final: attached, baselines: [], others: [] })).toEqual([])
    // Advice discussing another pack document's exhibit is not held to it.
    const advice = doc(["Advice Note", "Exhibit A sets out the notice the board approves."])
    const others = [{ title: "Board Consent", paras: attached, earlier: true }]
    expect(consistencyFindings({ title: "Advice Note", form: "advice", final: advice, baselines: [], others })).toEqual([])
  })

  test("the outline prompt plans an attachment as a trailing part, once, or as a separate instrument when another party executes it", () => {
    const prompt = stagePrompt(matter(), "d1.draft", { outline: true })
    expect(prompt).toContain('An attachment the document\'s text will refer to (a schedule, exhibit, annex, form of order or notice) is planned as a trailing "## " part of that document, after its other parts, labelled as the attachment is headed, its kind and number first;')
    expect(prompt).toContain('When the attachment is executed by a party who does not execute the document, plan it instead as its own "# " document, a separate instrument.')
    expect(prompt).toContain("An attachment a template part already carries is planned once, as that template part")
    expect(prompt).toContain("No document or part refers to an attachment the outline does not plan.")
  })
})
