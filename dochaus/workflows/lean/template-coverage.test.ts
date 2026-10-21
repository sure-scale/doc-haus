import { afterEach, describe, expect, test } from "bun:test"
import { readParagraphs } from "./checks"
import { readFileSync, writeFileSync } from "node:fs"
import { finalGate, type GateResult } from "./gate"
import { readLeanPlan } from "./plan"
import { runLean } from "./runner"
import { leanFile, readLeanJson, readOpenItems } from "./state"
import { plannedContract } from "./template-contract"
import { contractGaps, rowsBody } from "./template-coverage"
import { AGREEMENT, PARTS, PLAIN, TEMPLATE, TITLE, heading, templateMatter } from "./template-test-support"
import { AGENT, NO_FINDINGS, cleanupMatters, drafter, draftingNotes, noChanges, pidIn, sameSession, scriptedModel, type Reply } from "./test-support"

// Template coverage through the real coordinator. Only the model is scripted.

const LETTER = "# Covering letter\nForm: letter\n\n## Enclosure\n\nWe enclose the help desk agreement for signature.\n\n## Charges\n\nThe charges are as the agreement states.\n"
const PART_PROMPT = /^Draft part \d+ of \d+ of "[^"]*": .+$/m
const ACTING = "Acting for: not stated"
const FORM = "Form: agreement"
const REFERENCED_OUTLINE = [ACTING, TITLE, FORM, `Template: ${TEMPLATE}`, heading(PARTS.scope), heading(PARTS.charges), "The charges, with the charging schedule of Annex B.",
  heading(PARTS.credits), heading(PARTS.duration), heading(PARTS.annexA)].join("\n")
const REFERENCED = [TITLE, PARTS.scope, PARTS.charges, PARTS.credits, PARTS.duration, PARTS.annexA].join("\n\n")
const kinds = (directory: string, kind: string) => readOpenItems(directory).filter(item => item.kind === kind)

afterEach(cleanupMatters)

describe("template coverage in a lean run", () => {
  test("unheaded template parts still carry their text and later removal still fails coverage", async () => {
    const directory = templateMatter()
    const outline = [ACTING, TITLE, FORM, `Template: ${TEMPLATE}`, ...Object.values(PARTS).map(part => `${heading(part)}\nDisplay heading: none`)].join("\n")
    const model = scriptedModel(replies(templateDrafter([AGREEMENT], { outline: () => outline })))
    const state = await runLean(directory, { sessions: model.port })
    expect(state.status).toBe("completed")
    expect(kinds(directory, "omission")).toEqual([])
    const plan = readLeanPlan(directory)!.documents[0]!
    const rows = readParagraphs(directory, "d1")
    const contract = plannedContract(directory, "d1")!
    expect(contractGaps(contract, rowsBody(plan, rows))).toEqual([])
    const stripped = rows.filter(row => row.draft_part !== 0)
    expect(contractGaps(contract, rowsBody(plan, rows, stripped)).map(row => row.unit.label)).toContain(contract.top[0]!.label)
    expect(model.count(AGENT.draft)).toBe(7)
  }, 120_000)

  test("a draft carrying every part and annex of its template leaves no omission", async () => {
    const directory = templateMatter()
    const model = scriptedModel(replies(templateDrafter([AGREEMENT], { template: TEMPLATE })))
    const state = await runLean(directory, { sessions: model.port })

    expect(state.status).toBe("completed")
    expect(model.prompts.filter(row => row.agent === AGENT.draft)[0]!.text).toContain('add the line "Template: <exact matter document name>" under its title')
    expect(model.count(AGENT.draft)).toBe(7)
    const annexPart = model.prompts.find(prompt => prompt.text.includes("Draft part 5 of 6"))!.text
    expect(annexPart).toContain(`From the template "${TEMPLATE}", this part carries: "ANNEX A – HELP DESK SERVICES". Carry their text in full, adapted to this matter.`)
    expect([...kinds(directory, "omission"), ...kinds(directory, "template_removal")]).toEqual([])
    expect(readLeanJson(directory, "gate.json")).toMatchObject({ status: "pass", delivery: "execution_ready" })
  }, 120_000)

  test("an outline that does not account for an annex is asked again to place it or leave it out with a reason", async () => {
    const short = [ACTING, TITLE, FORM, ...Object.values(PARTS).filter(part => part !== PARTS.annexB).map(heading)].join("\n")
    const directory = templateMatter()
    const model = scriptedModel(replies(templateDrafter([AGREEMENT], { template: TEMPLATE, outline: index => index === 0 ? short : undefined })))
    const state = await runLean(directory, { sessions: model.port })

    expect(state.status).toBe("completed")
    expect(model.prompts.filter(row => row.agent === AGENT.draft)[1]!.text).toContain(`does not account for these parts of its template "${TEMPLATE}": "ANNEX B – CHARGING SCHEDULE".`)
    expect(model.prompts.filter(row => row.agent === AGENT.draft)[1]!.text).toContain('add a line "Left out: "<template part>" – <reason>" under the title.')
    expect(readFileSync(leanFile(directory, "d1/draft.md"), "utf8")).toContain("The monthly charge covers up to forty hours")
  }, 120_000)

  test("a templated outline with no Acting for line is asked for it together with the template parts it misses", async () => {
    const short = [TITLE, FORM, ...Object.values(PARTS).filter(part => part !== PARTS.annexB).map(heading)].join("\n")
    const directory = templateMatter()
    const model = scriptedModel(replies(templateDrafter([AGREEMENT], { template: TEMPLATE, outline: index => index === 0 ? short : undefined })))
    const state = await runLean(directory, { sessions: model.port })

    expect(state.status).toBe("completed")
    const retry = model.prompts.filter(row => row.agent === AGENT.draft)[1]!.text
    expect(retry).toContain('The outline has no "Acting for:" line.')
    expect(retry).toContain(`does not account for these parts of its template "${TEMPLATE}": "ANNEX B – CHARGING SCHEDULE".`)
    expect(readLeanPlan(directory)).toMatchObject({ documents: [{ template: TEMPLATE }], acting_for: { party: "", stated: false } })
  }, 120_000)

  test("a template part the rendered draft falls short on is a draft open item change-plan receives; the draft step asks for nothing more and rewrites nothing", async () => {
    const directory = templateMatter()
    const draft = templateDrafter([REFERENCED], { outline: () => REFERENCED_OUTLINE })
    // Were the draft step to ask for more after its parts, this reply would put the annex into the markdown.
    const offered: Reply = (prompt, index) => index === 0 || PART_PROMPT.test(prompt) ? draft(prompt, index) : PARTS.annexB
    const drafted: { markdown?: string; body?: Buffer } = {}
    const plan: Reply = () => {
      drafted.markdown = readFileSync(leanFile(directory, "d1/draft.md"), "utf8")
      drafted.body = readFileSync(leanFile(directory, "d1/body.docx"))
      return JSON.stringify({ shared_decisions: [], groups: [], dismissed: [], questions: [] })
    }
    const model = scriptedModel({ ...replies(offered), [AGENT["change-plan"]]: plan })
    const state = await runLean(directory, { sessions: model.port })

    expect(state.status).toBe("completed")
    expect(model.count(AGENT.draft)).toBe(1 + 5)
    // The outline's part labels are not printed.
    expect(drafted.markdown).toBe(`${REFERENCED.replace(/^## .+\n\n/gm, "")}\n`)
    expect(readFileSync(leanFile(directory, "d1/draft.md"), "utf8")).toBe(drafted.markdown!)
    expect(readFileSync(leanFile(directory, "d1/body.docx")).equals(drafted.body!)).toBe(true)
    const gaps = readOpenItems(directory).filter(item => item.stage === "draft" && item.kind === "omission")
    expect(gaps).toEqual([expect.objectContaining({ ref: { template_part: "ANNEX B – CHARGING SCHEDULE", source_id: TEMPLATE, document: "d1" } })])
    const asked = model.prompts.filter(prompt => prompt.agent === AGENT["change-plan"])
    expect(asked).toHaveLength(1)
    expect(asked[0]!.text).toMatch(new RegExp(`<required_repairs>[\\s\\S]*- ${gaps[0]!.id} \\(template part missing, cut short or without its heading in the draft\\): ${gaps[0]!.detail.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}[\\s\\S]*</required_repairs>`))
    // The part names no paragraph: the change cycle is told to add it from the template.
    expect(asked[0]!.text).toContain("For a template part missing or cut short in the draft, which names no paragraph, add the missing wording from the template: " +
      "split the paragraph the part belongs after, or replace a paragraph of the part with its full wording.")
  }, 120_000)

  test("an annex the change cycle does not add is kept as an omission the drafting notes disclose; the run does not fail", async () => {
    const directory = templateMatter()
    const model = scriptedModel(replies(templateDrafter([REFERENCED], { outline: () => REFERENCED_OUTLINE })))
    const state = await runLean(directory, { sessions: model.port })

    expect(state.status).toBe("completed")
    const omissions = kinds(directory, "omission").filter(item => item.stage === "apply")
    expect(omissions).toHaveLength(1)
    expect(omissions[0]).toMatchObject({ stage: "apply", ref: { template_part: "ANNEX B – CHARGING SCHEDULE", source_id: TEMPLATE, document: "d1" } })
    const gate = readLeanJson(directory, "gate.json") as GateResult
    expect(gate).toMatchObject({ status: "pass", delivery: "review_draft" })
    expect(gate.disclosed.map(item => item.id)).toContain(omissions[0]!.id)
  }, 120_000)

  test("a template part the outline leaves out with a reason is not retried or asked for, and the drafting notes disclose the removal", async () => {
    const outline = [ACTING, TITLE, FORM, `Template: ${TEMPLATE}`, 'Left out: "3. Service Credits" – the customer asked for no service credits.',
      ...Object.values(PARTS).filter(part => part !== PARTS.credits).map(heading)].join("\n")
    const draft = [TITLE, ...Object.values(PARTS).filter(part => part !== PARTS.credits)].join("\n\n")
    const directory = templateMatter()
    const model = scriptedModel(replies(templateDrafter([draft], { outline: () => outline })))
    const state = await runLean(directory, { sessions: model.port })

    expect(state.status).toBe("completed")
    expect(model.count(AGENT.draft)).toBe(1 + 5)
    expect(model.prompts.filter(prompt => PART_PROMPT.test(prompt.text)).some(prompt => prompt.text.includes("Service Credits"))).toBe(false)
    expect(readLeanPlan(directory)!.documents[0]!.left_out).toEqual([{ part: "3. Service Credits", reason: "the customer asked for no service credits." }])
    expect(kinds(directory, "omission")).toEqual([])
    const removals = kinds(directory, "template_removal")
    expect(removals.map(item => item.ref.template_part)).toEqual(["3 Service Credits"])
    expect(removals[0]!.detail).toContain("the customer asked for no service credits.")
    const gate = readLeanJson(directory, "gate.json") as GateResult
    expect(gate).toMatchObject({ status: "pass", delivery: "execution_ready", blocking: [] })
    expect(gate.items.find(item => item.id === removals[0]!.id)).toMatchObject({ status: "resolved" })

    writeFileSync(leanFile(directory, "drafting-notes.md"), "# Drafting notes\n\nNothing to report.\n")
    expect(finalGate(directory)).toMatchObject({ status: "blocked", delivery: null,
      blocking: [expect.objectContaining({ kind: "template_removal", resolution: "the drafting notes do not name this removal of a template part" })] })
  }, 120_000)

  test("a change list that cuts down a template part with a reason is not retried, and the drafting notes disclose the removal", async () => {
    const quote = { source_id: TEMPLATE, quote: "If the provider misses a response time in any month" }
    const changes: Reply = prompt => JSON.stringify({
      amendments: [{
        pid: pidIn(prompt, "If the provider misses"),
        updated: "The customer's remedies for a missed response time are those the law allows.",
        reason: "The customer asked for no service credits.",
        sources: [quote],
      }],
      questions: [],
    })
    const directory = templateMatter()
    const model = scriptedModel({ ...replies(templateDrafter([AGREEMENT], { template: TEMPLATE })), [AGENT.changes]: sameSession(changes) })
    const state = await runLean(directory, { sessions: model.port })

    expect(state.status).toBe("completed")
    const asked = model.prompts.filter(prompt => prompt.agent === AGENT.changes)
    expect(asked).toHaveLength(1)
    expect(asked[0]!.text).toContain(`Where an amendment takes out or shortens text the draft carries from its template "${TEMPLATE}", its reason says why.`)
    expect(kinds(directory, "omission")).toEqual([])
    const removals = kinds(directory, "template_removal")
    expect(removals.map(item => item.ref.template_part)).toEqual(["3 Service Credits"])
    expect(removals[0]!.detail).toContain("is cut down by amendment a1. Reason given: a1: The customer asked for no service credits.")
    const gate = readLeanJson(directory, "gate.json") as GateResult
    expect(gate).toMatchObject({ status: "pass", delivery: "execution_ready", blocking: [] })
    expect(gate.items.find(item => item.id === removals[0]!.id)).toMatchObject({ status: "resolved" })
  }, 120_000)

  test("a template with no headed parts is not checked", async () => {
    const outline = [ACTING, TITLE, FORM, `Template: ${PLAIN}`, heading(PARTS.scope), heading(PARTS.charges), heading(PARTS.duration)].join("\n")
    const directory = templateMatter()
    const draft = [TITLE, PARTS.scope, PARTS.charges, PARTS.duration].join("\n\n")
    const model = scriptedModel(replies(templateDrafter([draft], { outline: () => outline })))
    const state = await runLean(directory, { sessions: model.port })

    expect(state.status).toBe("completed")
    expect(model.count(AGENT.draft)).toBe(4)
    expect(model.prompts.some(prompt => /template "|From the template/.test(prompt.text))).toBe(false)
    expect(kinds(directory, "omission")).toEqual([])
  }, 120_000)

  test("a covering letter planned beside a templated document is not held to the template", async () => {
    const directory = templateMatter()
    const model = scriptedModel(replies(templateDrafter([LETTER, AGREEMENT], { template: TEMPLATE })))
    const state = await runLean(directory, { sessions: model.port })

    expect(state.status).toBe("completed")
    const letterParts = model.prompts.filter(prompt => prompt.text.includes('of "Covering letter":'))
    expect(letterParts).toHaveLength(2)
    expect(letterParts.some(prompt => prompt.text.includes("template"))).toBe(false)
    const agreement = model.prompts.find(prompt => prompt.text.includes('Draft part 1 of 6 of "Help desk agreement"'))!.text
    expect(agreement).toContain(`This document is drafted from the matter document "${TEMPLATE}" as its template.`)
    expect(agreement).toContain(`From the template "${TEMPLATE}", this part carries: "1 Scope of Work".`)
    expect(kinds(directory, "omission")).toEqual([])
    expect(readLeanJson(directory, "gate.json")).toMatchObject({ status: "pass" })
  }, 120_000)

  test("documents drafted from one template that share its parts between them are each held only to their own share", async () => {
    const annexes = ["# Help desk annexes", PARTS.annexA, PARTS.annexB].join("\n\n")
    const agreement = [TITLE, PARTS.scope, PARTS.charges, PARTS.credits, PARTS.duration].join("\n\n")
    const outline = [ACTING, TITLE, FORM, `Template: ${TEMPLATE}`, ...[PARTS.scope, PARTS.charges, PARTS.credits, PARTS.duration].map(heading),
      "# Help desk annexes", FORM, `Template: ${TEMPLATE}`, heading(PARTS.annexA), heading(PARTS.annexB)].join("\n")
    const directory = templateMatter()
    const model = scriptedModel(replies(templateDrafter([agreement, annexes], { outline: () => outline })))
    const state = await runLean(directory, { sessions: model.port })

    expect(state.status).toBe("completed")
    expect(model.count(AGENT.draft)).toBe(1 + 4 + 2)
    const annexPart = model.prompts.find(prompt => prompt.text.includes('Draft part 2 of 2 of "Help desk annexes"'))!.text
    expect(annexPart).toContain('this part carries: "ANNEX B – CHARGING SCHEDULE".')
    expect([...kinds(directory, "omission"), ...kinds(directory, "template_removal")]).toEqual([])
    expect(readLeanJson(directory, "gate.json")).toMatchObject({ status: "pass", delivery: "execution_ready" })
  }, 120_000)
})

/** The scripted drafter with the outline and its "Template:" line set by the test. */
function templateDrafter(documents: string[], options: { template?: string; outline?: (index: number) => string | undefined }): Reply {
  const base = drafter(...documents)
  const seen = { parts: false, outlines: 0 }
  return (prompt, index) => {
    if (PART_PROMPT.test(prompt)) {
      seen.parts = true
      return base(prompt, index)
    }
    if (seen.parts) return "No change."
    const outline = options.outline?.(seen.outlines++) ?? base(prompt, index)
    return options.template ? outline.replace(`${TITLE}\n`, `${TITLE}\nTemplate: ${options.template}\n`) : outline
  }
}

function replies(draft: Reply): Record<string, Reply> {
  return { [AGENT.draft]: draft, [AGENT.review]: () => NO_FINDINGS, [AGENT.changes]: noChanges, [AGENT.finalize]: draftingNotes }
}
