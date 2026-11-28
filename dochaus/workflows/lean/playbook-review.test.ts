import { afterEach, describe, expect, test } from "bun:test"
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { boundPlaybook, boundPlaybookContext } from "../../lib/playbook-source"
import { verifyQuote } from "../../lib/quote-verify"
import readDocument from "../../tool/read-document"
import { changePlanPrompt } from "./change-pipeline"
import { checkEvidence, evidencePrompt } from "./evidence"
import { packageReviewPrompt } from "./package-review"
import { leanSteps, parsePlan, plannedDocument, reviewGuidance, reviewStages, writeLeanPlan } from "./plan"
import { playbookSelectionIssues } from "./playbook-review"
import { stagePrompt } from "./prompts"
import { matterSourceIDs } from "./quotes"
import { researchPrompt } from "./research"
import { runLean, stageInputSha } from "./runner"
import { LEAN_CHILD_PERMISSIONS, researchBasis, stageInputs } from "./stages"
import { hashStageInputs, initLeanState, readLeanText, writeLeanText } from "./state"
import { AGENT, DRAFT, DRAFTED, SOURCE, SOURCE_TEXT, SUPPORT, cleanupMatters, drafter, happyPath, matter, pidOf, scriptedModel } from "./test-support"

describe("playbook source binding", () => {
  const roots: string[] = []
  const workspace = process.env.WORKSPACE_ROOT
  afterEach(() => {
    roots.splice(0).forEach(root => rmSync(root, { recursive: true, force: true }))
    if (workspace === undefined) delete process.env.WORKSPACE_ROOT
    else process.env.WORKSPACE_ROOT = workspace
  })

  function fixture() {
    const root = mkdtempSync(path.join(tmpdir(), "playbook-drafting-"))
    roots.push(root)
    process.env.WORKSPACE_ROOT = root
    const directory = path.join(root, "matter")
    const library = path.join(root, ".playbooks/playbook-workshop")
    mkdirSync(directory, { recursive: true })
    mkdirSync(library, { recursive: true })
    const text = "---\nname: playbook-workshop\ndescription: Workshop positions\n---\n\n## Cancellation\nPreferred: Cancellation notice is eight working days.\nFallback: Three working days only when a substitute is confirmed.\n"
    const file = path.join(library, "SKILL.md")
    writeFileSync(file, text)
    writeFileSync(path.join(directory, "matter.json"), JSON.stringify({ id: "matter", playbook: "playbook-workshop" }))
    initLeanState(directory, { request: "Propose changes for the organiser using the assigned guidance. Retain the agreed catering arrangement.", ownerSessionID: "ses_owner", model: { providerID: "test", modelID: "scripted" } })
    return { directory, text, file }
  }

  test("the selected library source is retrievable and quoted exactly without an indexed copy", async () => {
    const { directory, text } = fixture()
    expect(boundPlaybook(directory)?.text).toBe(text)
    expect(matterSourceIDs(directory)).toContain("playbook:playbook-workshop")
    const result = await readDocument.execute({ document: "playbook:playbook-workshop" }, { directory } as Parameters<typeof readDocument.execute>[1])
    expect(typeof result === "object" && result.output).toContain(text)
    expect(verifyQuote({ directory, document: "playbook:playbook-workshop", quote: "Three working days only when a substitute is confirmed." }).found).toBe(true)
    expect(verifyQuote({ directory, quote: "Three working days only when a substitute is confirmed." }).hits[0]?.doc).toBe("playbook:playbook-workshop")
    expect(verifyQuote({ directory, document: "playbook:playbook-other", quote: "Cancellation notice is eight working days." }).found).toBe(false)
    expect(verifyQuote({ directory, document: "playbook:playbook-workshop", quote: "Three working days are always acceptable." }).found).toBe(false)
    expect(checkEvidence(directory, "spine", { entries: [{ kind: "recommendation", detail: "Conditional notice fallback", evidence: [{ source_id: "playbook:playbook-workshop", quote: "Three working days only when a substitute is confirmed." }], limitations: [] }], limitations: [] })).toEqual([])
  })

  test("binding reaches early stages, document review and final package; changes invalidate dependent inputs", () => {
    const { directory, file } = fixture()
    for (const prompt of [evidencePrompt(directory, "discrepancy"), evidencePrompt(directory, "spine"), researchPrompt(directory, "intake", researchBasis(directory)), stagePrompt(directory, "d1.draft"), stagePrompt(directory, "d1.review-independent"), stagePrompt(directory, "d1.changes"), stagePrompt(directory, "finalize"), packageReviewPrompt(directory)]) expect(prompt).toContain("playbook:playbook-workshop")
    const before = hashStageInputs("d1.review-independent", "test", stageInputs(directory, "d1.review-independent"))
    writeFileSync(file, "The cancellation notice is now ten working days.")
    expect(hashStageInputs("d1.review-independent", "test", stageInputs(directory, "d1.review-independent"))).not.toBe(before)
    expect(verifyQuote({ directory, document: "playbook:playbook-workshop", quote: "Cancellation notice is eight working days." }).found).toBe(false)
    writeFileSync(path.join(directory, "matter.json"), JSON.stringify({ id: "matter" }))
    expect(boundPlaybookContext(directory)).toBe("")
    expect(matterSourceIDs(directory)).not.toContain("playbook:playbook-workshop")
  })

  test("missing and invalid bindings are disclosed without inferring policy or importing another library entry", () => {
    const { directory } = fixture()
    writeFileSync(path.join(directory, "matter.json"), JSON.stringify({ playbook: "playbook-missing" }))
    expect(boundPlaybookContext(directory)).toContain("unavailable")
    expect(matterSourceIDs(directory)).toEqual([])
    writeFileSync(path.join(directory, "matter.json"), JSON.stringify({ playbook: "../playbook-workshop" }))
    expect(boundPlaybook(directory)?.text).toBeUndefined()
  })
})

describe("playbook review stage", () => {
  afterEach(cleanupMatters)
  const selection = { sources: [{ source_id: SOURCE, quote: SOURCE_TEXT }], reason: "The supplied drafting guidance applies to this supply note." }
  const PLAYBOOK_FINDING = "Correction required at the delivery paragraph: the supplied terms require the stated delivery period."
  /** The review findings a change plan prompt lists for adjudication. */
  const listedFindings = (prompt: string) => JSON.parse(/<review_findings>\n([\s\S]*?)\n<\/review_findings>/.exec(prompt)?.[1] ?? "[]") as Record<string, unknown>[]
  const guidance = (prompt: string) => /<guidance>\n([\s\S]*?)\n<\/guidance>/.exec(prompt)?.[1] ?? ""
  const focus = (prompt: string) => /<review_focus>\n([\s\S]*?)\n<\/review_focus>/.exec(prompt)?.[1] ?? ""
  const doc = { title: "Supply note", purpose: "Record supplied terms", parts: [{ heading: "Delivery", scope: "Delivery" }] }

  test("planning routes the optional review per document and rejects invented evidence", () => {
    const directory = matter()
    expect(leanSteps(directory)).not.toContain("d1.review-playbook")
    const parsed = parsePlan(`Acting for: not stated\n# Supply note\nRecord terms\nForm: other\nPlaybook review: ${JSON.stringify(selection)}\n## Delivery\nDeliver goods\n# Factual inventory\nList sources only\nForm: other\nPlaybook review: {"sources":[],"reason":"A factual inventory imposes no supply terms."}\n## Sources\nList sources`)
    expect(parsed.issues).toEqual([])
    expect(playbookSelectionIssues(directory, parsed.value!.documents)).toEqual([])
    writeLeanPlan(directory, parsed.value!.documents)
    expect(leanSteps(directory)).toContain("d1.review-playbook")
    expect(leanSteps(directory)).not.toContain("d2.review-playbook")
    expect(reviewStages(directory, "d1")).toHaveLength(5)
    expect(reviewStages(directory, "d2")).toHaveLength(4)
    const changed = structuredClone(parsed.value!.documents)
    changed[0]!.playbook_review!.sources[0]!.quote = "No such guidance was supplied."
    expect(playbookSelectionIssues(directory, changed)).toHaveLength(1)
    changed[0]!.playbook_review!.sources[0]!.source_id = "unknown.docx"
    expect(playbookSelectionIssues(directory, changed)).toHaveLength(1)
    expect(parsePlan('Acting for: not stated\n# Note\nPlaybook review: invalid\n## Terms').value).toBeUndefined()
  })

  test("a document drafted from a template gets the guidance review with the template among its guidance", () => {
    const directory = matter()
    writeLeanPlan(directory, [{ ...doc, template: SOURCE }, { ...doc, title: "Cover note" }])
    expect(reviewStages(directory, "d1")).toContain("review-playbook")
    expect(leanSteps(directory)).toContain("d1.review-playbook")
    expect(reviewGuidance(directory, "d1")).toEqual([SOURCE])
    // Without a template or a selected playbook, a document keeps the ordinary reviews.
    expect(reviewStages(directory, "d2")).not.toContain("review-playbook")
    expect(leanSteps(directory)).not.toContain("d2.review-playbook")
    const prompt = stagePrompt(directory, "d1.review-playbook")
    expect(guidance(prompt)).toContain(`- ${SOURCE}`)
  })

  test("a revision with no selected playbook keeps the ordinary reviews: its original is the baseline, not guidance", () => {
    const directory = matter()
    writeLeanPlan(directory, [{ ...doc, mode: "revise", template: SOURCE, comparison: SOURCE }, { ...doc, title: "Revised note", mode: "revise", template: SOURCE, playbook_review: selection }])
    expect(reviewGuidance(directory, "d1")).toEqual([])
    expect(reviewStages(directory, "d1")).not.toContain("review-playbook")
    expect(leanSteps(directory)).not.toContain("d1.review-playbook")
    // A selected playbook still schedules it, without the original or the template focus.
    expect(reviewGuidance(directory, "d2")).toEqual([SOURCE])
  })

  test("the guidance list is the selected playbook sources plus the template, adding template focus only with a template", () => {
    const directory = matter()
    writeLeanPlan(directory, [{ ...doc, playbook_review: selection, template: "precedent.md" }, { ...doc, title: "Cover note", playbook_review: selection }])
    expect(reviewGuidance(directory, "d1")).toEqual([SOURCE, "precedent.md"])
    expect(reviewGuidance(directory, "d2")).toEqual([SOURCE])
    const templated = stagePrompt(directory, "d1.review-playbook")
    expect(guidance(templated)).toContain(`- ${SOURCE}`)
    expect(guidance(templated)).toContain("- precedent.md")
    const selected = stagePrompt(directory, "d2.review-playbook")
    expect(guidance(selected)).toContain(`- ${SOURCE}`)
    expect(guidance(selected)).not.toContain("precedent.md")
    expect(focus(templated).startsWith(focus(selected))).toBe(true)
    expect(focus(templated).length).toBeGreaterThan(focus(selected).length)
    expect(packageReviewPrompt(directory)).toContain(`<playbook_review document="d1">\nSelection: ${JSON.stringify(selection)}\nTemplate: precedent.md\n`)
  })

  test("a changed selection revisits consumers while an unassigned body retains the original reviews", () => {
    const directory = matter()
    writeLeanPlan(directory, [{ ...doc, playbook_review: selection }])
    const steps = ["d1.review-independent", "d1.review-playbook", "d1.change-plan", "d1.change-review", "package-review"] as const
    const before = steps.map(step => stageInputSha(directory, step))
    writeLeanPlan(directory, [{ ...doc, playbook_review: { sources: [], reason: "The lawyer restricted this body to an inventory." } }])
    steps.forEach((step, index) => expect(stageInputSha(directory, step)).not.toBe(before[index]))
    expect(reviewStages(directory, "d1")).not.toContain("review-playbook")
  })

  test("the native coordinator accepts the specialist before adjudication and final package review", async () => {
    const directory = matter()
    const drafting = drafter(DRAFT)
    const base = scriptedModel({ ...happyPath, [AGENT.draft]: (prompt, index) => {
      const reply = drafting(prompt, index)
      return reply.includes("Acting for:") ? reply.replace("# Supply note", `# Supply note\nPlaybook review: ${JSON.stringify(selection)}`) : reply
    } })
    let checked = false
    const sessions = { ...base.port, async prompt(input: Parameters<typeof base.port.prompt>[0]) {
      if (input.agent === AGENT["review-playbook"]) {
        expect(input.text).toContain(JSON.stringify(selection))
        checked = true
        return JSON.stringify({ coverage: "Playbook coverage: supplied terms require the stated delivery period.", findings: [{
          pid: pidOf(input.text, DRAFTED), origin: "request", balance: "neutral", finding: PLAYBOOK_FINDING, sources: SUPPORT,
        }] })
      }
      if (input.agent === AGENT["change-plan"]) {
        expect(checked).toBe(true)
        // The plan adjudicates the specialist's finding, listed with its review.
        expect(listedFindings(input.text)).toContainEqual(expect.objectContaining({ review: "review-playbook", finding: PLAYBOOK_FINDING }))
      }
      if (input.agent === AGENT["package-review"]) expect(input.text).toContain("Playbook coverage:")
      return base.port.prompt(input)
    } }
    const state = await runLean(directory, { sessions })
    expect(state.status).toBe("completed")
    expect(state.stages["d1.review-playbook"]).toMatchObject({ status: "done", outputs: ["d1/review-playbook.json"] })
    expect(base.sessions.find(session => session.agent === AGENT["review-playbook"])?.permissions).toEqual(LEAN_CHILD_PERMISSIONS)
    expect(plannedDocument(directory, "d1")?.playbook_review).toEqual(selection)
    expect(listedFindings(changePlanPrompt(directory, "d1"))).toContainEqual(expect.objectContaining({ review: "review-playbook", finding: PLAYBOOK_FINDING }))
    expect(packageReviewPrompt(directory)).toContain("Playbook coverage:")
    const before = stageInputSha(directory, "d1.change-plan")
    writeLeanText(directory, "d1/review-playbook.json", JSON.stringify({ coverage: "Revised coverage after a source-backed correction.", findings: [] }))
    expect(stageInputSha(directory, "d1.change-plan")).not.toBe(before)
  }, 120_000)

  test("an invented planning quotation never enters the best-outline fallback", async () => {
    const directory = matter()
    const invalid = { sources: [{ source_id: SOURCE, quote: "This policy passage was never supplied." }], reason: "Purported applicable guidance" }
    const model = scriptedModel({ ...happyPath, [AGENT.draft]: () => `Acting for: not stated\n# Supply note\nPlaybook review: ${JSON.stringify(invalid)}\n## Delivery\nState delivery terms` })
    const state = await runLean(directory, { sessions: model.port })
    expect(state.status).toBe("failed")
    expect(state.stop_reason).toContain("quote its scope or position exactly")
    expect(readLeanText(directory, "plan.json")).toBeUndefined()
    expect(readLeanText(directory, "d1/draft.md")).toBeUndefined()
    expect(model.count(AGENT["review-playbook"])).toBe(0)
  }, 120_000)
})
