import { afterEach, describe, expect, test } from "bun:test"
import { acceptChangePlan, unadjudicatedFindings, writerGroup, writerPlan, type ChangeGroup, type ChangePlan, type ReviewDisposition } from "./change-pipeline"
import { CHECKLIST_JOB_AGENTS } from "./checklist-job"
import type { ChecklistCheckerInput, ChecklistProducerInput } from "./checklist-contract"
import { mergeRepairAdjudications, type RepairAdjudication } from "./checklist-repair-adjudication"
import { writeLeanPlan } from "./plan"
import { enqueueLeanRequest } from "./requests"
import { runLean } from "./runner"
import { readLeanJson, readLeanText, writeLeanText, writeQuestions, type Question } from "./state"
import { AGENT, DRAFTED, SUPPORT, cleanupMatters, happyPath, listedFindingIDs, matter, pidOf, scriptedModel, type Reply } from "./test-support"

// How a change plan adjudicates review findings by id: the resolves a focused
// checklist merge keeps, what acceptance saves and asks, what amendment writers
// are shown, and when the runner may take the focused path at all.

afterEach(cleanupMatters)

const PID = "p111111111111", OTHER = "p222222222222"
const FINDING_ID = /\bf-[0-9a-f]{10}\b/

const group = (id: string, pids: string[], resolves?: string[]): ChangeGroup =>
  ({ id, pids, goal: `Goal ${id}.`, sources: SUPPORT, limitations: [], ...(resolves ? { resolves } : {}) })
const disposition = (detail: string, resolves?: string[]): ReviewDisposition =>
  ({ disposition: "deferred", kind: "information", detail, reason: "Still open.", documents: ["d1"], sources: [], limitations: ["Unverified."], ...(resolves ? { resolves } : {}) })
const question = (text: string): Question & { resolves?: string[] } =>
  ({ id: `q-${text.length}`, stage: "change-plan", question: text, context: "Why it matters.", pid: PID, status: "open", kind: "fact" })

describe("a focused checklist merge keeps the review findings a plan resolves", () => {
  const previous: ChangePlan = { shared_decisions: [], groups: [group("old", [OTHER], ["f-000000000a"])], dismissed: [], questions: [],
    findings: [disposition("Prior timing assessment.", ["f-000000000b"]), disposition("Unrelated advice.", ["f-000000000c"])] }
  const superseding = (patch: Partial<RepairAdjudication>): RepairAdjudication => ({ shared_decisions: [], groups: [], dismissed: [], findings: [], questions: [],
    superseded_decisions: [], superseded_findings: [{ index: 0, reason: "The source calls for advice.", sources: SUPPORT }], ...patch })

  test("a superseded disposition's resolves pass to the first group, else finding, else question, of the same result", () => {
    const toGroup = mergeRepairAdjudications(previous, [superseding({ groups: [group("timing", [PID]), group("later", ["p333333333333"])], findings: [disposition("Also deferred.")] })])
    expect(toGroup.groups.map(row => row.resolves)).toEqual([["f-000000000b"], undefined])
    expect(toGroup.findings).toEqual([previous.findings![1]!, disposition("Also deferred.")])

    const toFinding = mergeRepairAdjudications(previous, [superseding({ findings: [disposition("Deferred again."), disposition("Second.")], questions: [question("Is approval given?")] })])
    expect(toFinding.findings!.map(row => row.resolves)).toEqual([["f-000000000c"], ["f-000000000b"], undefined])
    expect(toFinding.questions.slice(-1)[0]!.resolves).toBeUndefined()

    const toQuestion = mergeRepairAdjudications(previous, [superseding({ questions: [question("Is approval given?")] })])
    expect(toQuestion.questions.map(row => row.resolves)).toEqual([["f-000000000b"]])
  })

  test("a merged overlapping group keeps the union of its parts' resolves", () => {
    const merged = mergeRepairAdjudications(previous, [
      superseding({ groups: [group("timing", [PID], ["f-000000000d"])] }),
      superseding({ superseded_findings: [], groups: [group("dependent", [PID, OTHER], ["f-000000000e", "f-000000000d"])] }),
    ])
    expect(merged.groups).toHaveLength(1)
    expect(merged.groups[0]!.resolves!.toSorted()).toEqual(["f-000000000b", "f-000000000d", "f-000000000e"])
  })
})

describe("acceptChangePlan", () => {
  const OWN = "f-00000000a1", SECOND = "f-00000000a2"
  /** One planned document with one paragraph and a saved review of `findings`. */
  function reviewed(...findings: string[]) {
    const directory = matter()
    writeLeanPlan(directory, [{ title: "Supply note", purpose: "Record the supply terms", file: "supply note.docx", parts: [{ heading: "Delivery", scope: "Delivery" }] }])
    writeLeanText(directory, "d1/paragraphs.json", JSON.stringify([{ pid: PID, sha256: "initial", text: DRAFTED }]))
    writeLeanText(directory, "d1/review.json", JSON.stringify({ findings: findings.map((id, index) => ({ id, document: "d1", origin: "drafting", finding: `Finding ${index + 1}.`, sources: [] })) }))
    return directory
  }
  const reply = (patch: Partial<Record<keyof ChangePlan, unknown>>) => JSON.stringify({ shared_decisions: [], groups: [], dismissed: [], findings: [], questions: [], ...patch })
  const ask = (text: string, resolves?: string[]) => ({ pid: PID, question: text, kind: "fact", context: "Why it matters.", ...(resolves ? { resolves } : {}) })

  test("a question's resolves are saved in the change plan only; the questions to ask carry none", () => {
    const directory = reviewed(OWN)
    const accepted = acceptChangePlan(directory, "d1", reply({ questions: [ask("Is approval given?", [OWN])] }))
    expect(accepted.issues).toEqual([])
    expect(accepted.questions).toHaveLength(1)
    expect(accepted.questions![0]).not.toHaveProperty("resolves")
    expect((readLeanJson(directory, "d1/change-plan.json") as ChangePlan).questions[0]!.resolves).toEqual([OWN])
  })

  test("an id that names no listed finding is dropped without a retry", () => {
    const directory = reviewed(OWN)
    expect(acceptChangePlan(directory, "d1", reply({ findings: [disposition("Kept open.", [OWN, "f-00000000ff"])] })).issues).toEqual([])
    expect((readLeanJson(directory, "d1/change-plan.json") as ChangePlan).findings![0]!.resolves).toEqual([OWN])
  })

  test("a listed finding left out is retried before the last attempt and saved on it", () => {
    const directory = reviewed(OWN, SECOND)
    const plan = reply({ findings: [disposition("Kept open.", [OWN])] })
    const retried = acceptChangePlan(directory, "d1", plan)
    expect(retried.issues).toHaveLength(1)
    expect(retried.issues[0]).toContain(`These review findings are not adjudicated: ${SECOND}.`)
    expect(retried.issues[0]).not.toContain("An answered question")
    expect(readLeanText(directory, "d1/change-plan.json")).toBeUndefined()
    expect(acceptChangePlan(directory, "d1", plan, { final: true }).issues).toEqual([])
    expect(readLeanText(directory, "d1/change-plan.json")).toBeDefined()
  })

  test("the answered-question sentence appears only when the answered question carried resolves", () => {
    const directory = reviewed(OWN, SECOND)
    writeQuestions(directory, [{ id: "q-answered", stage: "change-plan", document: "d1", pid: PID, question: "Is approval given?", context: "Why it matters.", status: "answered", answer: "Yes." }])
    const findings = [disposition("Kept open.", [OWN])]
    const carried = acceptChangePlan(directory, "d1", reply({ findings, questions: [ask("Is approval given?", [SECOND])] }))
    expect(carried.issues[0]).toContain(`not adjudicated: ${SECOND}.`)
    expect(carried.issues[0]).toContain("An answered question no longer adjudicates a finding")
    const bare = acceptChangePlan(directory, "d1", reply({ findings, questions: [ask("Is approval given?")] }))
    expect(bare.issues[0]).toContain(`not adjudicated: ${SECOND}.`)
    expect(bare.issues[0]).not.toContain("An answered question")
  })
})

describe("amendment writers are shown no finding ids", () => {
  test("writerPlan and writerGroup strip resolves and settled groups", () => {
    const plan: ChangePlan = { shared_decisions: ["Keep payment."], groups: [group("timing", [PID], ["f-000000000a"])], dismissed: [],
      findings: [disposition("Kept open.", ["f-000000000b"])], questions: [{ ...question("Is approval given?"), resolves: ["f-000000000c"] }], settled_groups: [group("old", [OTHER], ["f-000000000d"])] }
    const shown = JSON.stringify(writerPlan(plan))
    expect(shown).not.toContain("resolves")
    expect(shown).not.toContain("settled_groups")
    expect(shown).not.toMatch(FINDING_ID)
    expect(writerPlan(plan).groups).toEqual([group("timing", [PID])])
    expect(writerGroup(plan.groups[0]!)).toEqual(group("timing", [PID]))
  })

  test("in a run, only the change plan prompt lists finding ids; preparation, changes and change review see none", async () => {
    const directory = matter()
    // The plan amends for every listed finding in one group, so preparation runs.
    const plan: Reply = prompt => JSON.stringify({ shared_decisions: [], dismissed: [], findings: [], questions: [],
      groups: [{ id: "delivery", pids: [pidOf(prompt, DRAFTED)], goal: "State the delivery period.", sources: SUPPORT, limitations: [], resolves: listedFindingIDs(prompt) }] })
    const model = scriptedModel({ ...happyPath, [AGENT["change-plan"]]: plan })
    expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
    const texts = (agent: string) => model.prompts.filter(row => row.agent === agent).map(row => row.text)
    expect(texts(AGENT["change-plan"]).some(text => FINDING_ID.test(text))).toBe(true)
    for (const agent of [AGENT["change-prepare"], AGENT.changes, AGENT["change-review"]]) {
      expect([agent, texts(agent).length > 0]).toEqual([agent, true])
      expect([agent, texts(agent).filter(text => FINDING_ID.test(text))]).toEqual([agent, []])
    }
  }, 120_000)
})

describe("the focused checklist repair path", () => {
  const FOCUSED = "ONE supplied checklist assessment"
  // One request requirement the checker finds unsatisfied, so a checklist
  // repair round starts after the first pass. The focused adjudication defers
  // it; a full plan rejects every listed review finding.
  function replies(fail: { once: boolean }): Record<string, Reply> {
    return {
      ...happyPath,
      [AGENT["change-plan"]]: text => {
        if (text.includes(FOCUSED)) {
          if (fail.once) { fail.once = false; throw new Error("Scripted adjudication failure") }
          return JSON.stringify({ shared_decisions: [], groups: [], dismissed: [], questions: [], superseded_findings: [], superseded_decisions: [], findings: [
            { disposition: "deferred", kind: "information", detail: "The assessed requirement still needs a correction.", reason: "No supported correction.", documents: ["d1"], sources: [], limitations: ["No supported correction identified."] }] })
        }
        const resolves = listedFindingIDs(text)
        return JSON.stringify({ shared_decisions: [], groups: [], dismissed: [], questions: [], findings: resolves.length ? [
          { disposition: "rejected", kind: "suggestion", detail: "The scripted plan adjudicates the review findings.", reason: "The scripted changes stage decides the amendments.", documents: ["d1"], sources: [], limitations: ["Scripted fixture."], resolves }] : [] })
      },
      [CHECKLIST_JOB_AGENTS.extract]: text => {
        const input = JSON.parse(text.split("SOURCE_INPUT:\n\n")[1]!) as ChecklistProducerInput
        return JSON.stringify({ items: input.source.document === "request" ? [{ check: "The requested supply note is drafted.", basis: "directive", source_refs: input.passages, qualifications: [], target: { ...input.targets[0], treatment: "operative" } }] : [] })
      },
      [CHECKLIST_JOB_AGENTS.check]: text => {
        const input = JSON.parse(text.split("CHECK_INPUT:\n\n")[1]!) as ChecklistCheckerInput
        return JSON.stringify({ item_id: input.item_id, status: "unsatisfied", reason: "Focused result unsatisfied.", evidence: input.current_evidence,
          source_refs: input.item.source_refs, inspection_scope: "Current body paragraphs only.", limitation: "Comments not inspected." })
      },
    }
  }
  const focused = (model: ReturnType<typeof scriptedModel>) => model.prompts.filter(row => row.agent === AGENT["change-plan"] && row.text.includes(FOCUSED)).length

  test("a review finding the saved plan does not name sends the round's change plan to a full plan", async () => {
    const directory = matter()
    const fail = { once: true }
    const first = scriptedModel(replies(fail))
    expect((await runLean(directory, { sessions: first.port })).status).toBe("failed")
    // The saved plan named every finding, so the round took the focused path.
    expect(focused(first)).toBe(1)
    expect(unadjudicatedFindings(directory, "d1")).toEqual([])

    // A review record that changed since the plan: new finding text, new id.
    const NEW = "f-00000000b1"
    const review = readLeanJson(directory, "d1/review.json") as { findings: unknown[] }
    writeLeanText(directory, "d1/review.json", JSON.stringify({ findings: [...review.findings, { id: NEW, document: "d1", origin: "drafting", finding: "The payment paragraph needs a due date.", sources: [] }] }))
    expect(unadjudicatedFindings(directory, "d1").map(row => row.id)).toEqual([NEW])

    enqueueLeanRequest(directory, { kind: "continue" })
    const resumed = scriptedModel(replies(fail))
    expect((await runLean(directory, { sessions: resumed.port })).status).toBe("completed")
    expect(focused(resumed)).toBe(0)
    expect(resumed.prompts.some(row => row.agent === AGENT["change-plan"] && listedFindingIDs(row.text).includes(NEW))).toBe(true)
    expect(unadjudicatedFindings(directory, "d1")).toEqual([])
  }, 180_000)
})
