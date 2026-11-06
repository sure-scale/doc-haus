import { afterEach, expect, test } from "bun:test"
import { assembleDiscrepancy, runDiscrepancyStage, UNVERIFIED_GAP } from "./discrepancy"
import { evidenceOutputValid, evidencePrompt } from "./evidence"
import { writeInstructionsFile } from "./instructions"
import { withWorkflowLimit } from "./model-concurrency-test-support"
import { researchPrompt } from "./research"
import { researchBasis, stageInputs, type StageSessionPort } from "./stages"
import { readLeanJson, readLeanState, writeLeanText } from "./state"
import { AGENT, cleanupMatters, matter, scriptedModel, SOURCE, SOURCE_TEXT, type Reply } from "./test-support"

afterEach(cleanupMatters)

// The discrepancy stage's jobs, told apart by their prompts; a retry in a
// session answers as the job that session asked.
const ENTRY = { detail: "Delivery term appears in the source", evidence: [{ source_id: SOURCE, quote: SOURCE_TEXT }], limitations: [], material: false, status: "benign", resolution: "Delivery and payment are distinct events." }
const GAP = { detail: "The supply terms state no governing law; the note needs one.", requirement: { source_id: "request", quote: "Draft a short supply note" },
  evidence: [{ source_id: SOURCE, quote: "Payment is due on delivery." }], limitations: [], material: true, lawyer: true }
// A gap the source in fact supplies: the delivery period.
const PERIOD = { ...GAP, detail: "The supply terms give no delivery period; the note needs one.", evidence: [], lawyer: false }
const RESTATED = { ...GAP, detail: "The delivery term restated as a gap.", evidence: ENTRY.evidence }
const SUPPLIES = { supplied: true, evidence: [{ source_id: SOURCE, quote: "within thirty days of the order date" }], limitations: [] }
const SILENT = { supplied: false, evidence: [], limitations: [] }

type Job = "differences" | "gaps" | "verify"
const kind = (prompt: string): Job => prompt.includes("<gap>\n") ? "verify" : prompt.includes("<output_contract>\nReturn JSON {\"gaps\"") ? "gaps" : "differences"

/** Scripted jobs: `verify` is given the gap's detail as its prompt shows it. */
function jobs(replies: { differences?: (attempt: number) => object; gaps?: (attempt: number) => object; verify?: (detail: string, attempt: number) => object } = {}) {
  const asked = new Map<string | undefined, { job: Job; detail: string }>()
  const attempts = { differences: 0, gaps: 0, verify: 0 }
  const perSession = new Map<string | undefined, number>()
  const reply: Reply = (prompt, _index, sessionID) => {
    if (!asked.has(sessionID)) asked.set(sessionID, { job: kind(prompt), detail: /<gap>\nA point another job found no matter source settles:\n(.*)\n/.exec(prompt)?.[1] ?? "" })
    const { job, detail } = asked.get(sessionID)!
    attempts[job]++
    const attempt = perSession.get(sessionID) ?? 0
    perSession.set(sessionID, attempt + 1)
    if (job === "differences") return JSON.stringify(replies.differences?.(attempt) ?? { entries: [ENTRY], limitations: ["Differences limit."] })
    if (job === "gaps") return JSON.stringify(replies.gaps?.(attempt) ?? { gaps: [GAP], limitations: ["Gaps limit."] })
    return JSON.stringify(replies.verify?.(detail, attempt) ?? SILENT)
  }
  return { reply, attempts, sessions: () => [...asked.values()].map(row => row.job) }
}

async function run(directory: string, reply: Reply) {
  const model = scriptedModel({ [AGENT.discrepancy]: reply })
  const result = await runDiscrepancyStage(directory, readLeanState(directory)!, { sessions: model.port }, () => {})
  return { result, model }
}

test("the differences and gaps jobs run in parallel, and the gaps' verify jobs run in parallel within the workflow limit", async () => withWorkflowLimit(3, async () => {
  const directory = matter()
  const gaps = ["governing law", "notice address", "signatory", "currency", "term"].map(point => ({ ...GAP, detail: `The supply terms give no ${point}; the note needs one.` }))
  const events: string[] = []
  const active = { verify: 0, most: 0 }
  const texts = new Map<string, string>()
  let created = 0
  const sessions: StageSessionPort = {
    async createSession() { return `ses_${++created}` },
    async prompt(input) {
      const job = kind(texts.get(input.sessionID) ?? input.text)
      texts.set(input.sessionID, texts.get(input.sessionID) ?? input.text)
      events.push(`${job} asked`)
      if (job === "verify") active.most = Math.max(active.most, ++active.verify)
      await new Promise(resolve => setTimeout(resolve, 5))
      if (job === "verify") active.verify--
      events.push(`${job} answered`)
      return JSON.stringify(job === "differences" ? { entries: [ENTRY], limitations: [] } : job === "gaps" ? { gaps, limitations: [] } : SILENT)
    },
  }
  const result = await runDiscrepancyStage(directory, readLeanState(directory)!, { sessions }, () => {})
  expect(result.status).toBe("done")
  // Both are asked before either answers; the gaps job is not shown the differences.
  expect(events.slice(0, 2).toSorted()).toEqual(["differences asked", "gaps asked"])
  expect([...texts.values()].find(text => kind(text) === "gaps")).not.toContain(ENTRY.detail)
  expect(events.filter(row => row === "verify asked")).toHaveLength(5)
  expect(active.most).toBe(3)
  expect(readLeanJson(directory, "discrepancy.json")).toEqual({ entries: [ENTRY], gaps, limitations: [] })
}))

test("a gap a source supplies is recorded under supplied with its words; the others keep any passage that speaks to them, and intake reads no supplied gap", async () => {
  const directory = matter()
  const speaks = { source_id: SOURCE, quote: "The supplier shall deliver the goods" }
  const scripted = jobs({ gaps: () => ({ gaps: [GAP, PERIOD], limitations: ["Gaps limit."] }), verify: detail => detail === PERIOD.detail ? SUPPLIES : { ...SILENT, evidence: [speaks] } })
  const { result, model } = await run(directory, scripted.reply)
  expect(result).toEqual({ status: "done", outputs: ["discrepancy.json"], openItems: [] })
  expect(scripted.sessions().toSorted()).toEqual(["differences", "gaps", "verify", "verify"])
  const verify = model.prompts.find(prompt => prompt.text.includes(PERIOD.detail) && kind(prompt.text) === "verify")!.text
  expect([...verify.matchAll(/^<(\w+)>$/gm)].map(match => match[1])).toEqual(["gap", "sources", "question", "output_contract"])
  expect(verify).toContain(`"${PERIOD.requirement.quote}"`)
  expect(verify).toContain(`- ${SOURCE}`)
  const record = { entries: [ENTRY], gaps: [{ ...GAP, evidence: [...GAP.evidence, speaks] }], supplied: [{ ...PERIOD, evidence: SUPPLIES.evidence }], limitations: ["Differences limit.", "Gaps limit."] }
  expect(readLeanJson(directory, "discrepancy.json")).toEqual(record)
  expect(Object.keys(readLeanJson(directory, "discrepancy.json") as object)).toEqual(["entries", "gaps", "supplied", "limitations"])
  expect(evidenceOutputValid(directory, "discrepancy")).toBe(true)

  // Intake asks or discloses the gaps alone; the spine reads the whole record.
  const intake = researchPrompt(directory, "intake", researchBasis(directory))
  expect(intake).toContain(GAP.detail)
  expect(intake).not.toContain(PERIOD.detail)
  expect(intake).not.toContain('"supplied"')
  const keyed = stageInputs(directory, "intake").find(part => part.name === "discrepancy.json")!
  expect(JSON.parse(String(keyed.bytes))).toEqual({ entries: record.entries, gaps: record.gaps, limitations: record.limitations })
  expect(evidencePrompt(directory, "spine")).toContain(PERIOD.detail)
})

test("a verify reply whose quote is not found retries in its own session; one that never passes keeps its gap with a limitation", async () => {
  const directory = matter()
  const invented = { supplied: true, evidence: [{ source_id: SOURCE, quote: "Governed by the law of nowhere." }], limitations: [] }
  const scripted = jobs({ gaps: () => ({ gaps: [GAP, PERIOD], limitations: [] }), verify: (detail, attempt) => detail === PERIOD.detail && attempt ? SUPPLIES : invented })
  const { result, model } = await run(directory, scripted.reply)
  expect(result.status).toBe("done")
  // GAP's job: the first ask and three retries; PERIOD's passes on its retry.
  expect(scripted.attempts).toEqual({ differences: 1, gaps: 1, verify: 6 })
  const retries = model.prompts.filter(prompt => prompt.text.startsWith("Your reply did not pass the discrepancy checks:"))
  expect(retries).toHaveLength(4)
  expect(retries[0]!.text).toContain('evidence[0] quotes "Governed by the law of nowhere." as from "supply-terms.docx", but those words are not in that document.')
  expect(new Set(retries.map(prompt => prompt.sessionID)).size).toBe(2)
  expect(readLeanJson(directory, "discrepancy.json")).toEqual({
    entries: [ENTRY], gaps: [{ ...GAP, limitations: [UNVERIFIED_GAP] }], supplied: [{ ...PERIOD, evidence: SUPPLIES.evidence }], limitations: ["Differences limit."],
  })
  expect(evidenceOutputValid(directory, "discrepancy")).toBe(true)
})

test("code leaves out a gap whose checked quote restates a difference, whether the gaps job or its verify job quoted it", async () => {
  const directory = matter()
  const echoed = { ...PERIOD, detail: "The delivery term is found again by the verifier." }
  const scripted = jobs({ gaps: () => ({ gaps: [RESTATED, GAP, echoed], limitations: [] }), verify: detail => detail === echoed.detail ? { ...SILENT, evidence: ENTRY.evidence } : SILENT })
  expect((await run(directory, scripted.reply)).result.status).toBe("done")
  // Accepted at once: restatement is not the gaps job's check.
  expect(scripted.attempts).toEqual({ differences: 1, gaps: 1, verify: 3 })
  expect(readLeanJson(directory, "discrepancy.json")).toEqual({ entries: [ENTRY], gaps: [GAP], limitations: ["Differences limit."] })
  expect(readLeanJson(directory, "discrepancy-gaps.json")).toMatchObject({ gaps: [RESTATED, GAP, echoed] })
  expect(assembleDiscrepancy({ entries: [], limitations: [] }, { gaps: [RESTATED], limitations: [] }, [SILENT]).gaps).toEqual([RESTATED])
})

test("a gaps reply that fails its check retries in its own session without asking another job again", async () => {
  const directory = matter()
  const scripted = jobs({ gaps: attempt => ({ gaps: [attempt ? GAP : { ...GAP, requirement: { source_id: "request", quote: "Draft a long supply agreement" } }], limitations: [] }) })
  const { result, model } = await run(directory, scripted.reply)
  expect(result.status).toBe("done")
  expect(scripted.attempts).toEqual({ differences: 1, gaps: 2, verify: 1 })
  expect(model.sessions).toHaveLength(3)
  const retry = model.prompts.find(prompt => prompt.text.startsWith("Your reply did not pass"))!
  expect(retry.text).toContain('gaps[0].requirement quotes "Draft a long supply agreement" as from "request"')
  expect(kind(model.prompts.find(prompt => prompt.sessionID === retry.sessionID)!.text)).toBe("gaps")
})

test("a resumed run asks only the jobs whose answer is missing; a new instruction list re-asks only the gaps job", async () => {
  const directory = matter()
  const failing = await run(directory, jobs({ differences: () => ({ entries: [{ ...ENTRY, status: "maybe" }], limitations: [] }), gaps: () => ({ gaps: [GAP, PERIOD], limitations: [] }) }).reply)
  expect(failing.result).toEqual({ status: "failed", reason: expect.stringMatching(/^discrepancy: differences checker failed after 3 retries: entries\[0\] needs status/) })
  expect(readLeanJson(directory, "discrepancy.json")).toBeUndefined()

  // The gaps and both verify answers were kept; a verify job that errors stops the stage.
  const stopped = jobs({ differences: () => ({ entries: [ENTRY], limitations: [] }) })
  writeLeanText(directory, "discrepancy-verify.json", JSON.stringify({ records: (readLeanJson(directory, "discrepancy-verify.json") as { records: object[] }).records.slice(0, 1) }))
  const erring = scriptedModel({ [AGENT.discrepancy]: (prompt, index, session) => {
    if (kind(prompt) === "verify") throw new Error("The engine could not complete the turn.")
    return stopped.reply(prompt, index, session)
  } })
  await expect(runDiscrepancyStage(directory, readLeanState(directory)!, { sessions: erring.port }, () => {})).rejects.toThrow("The engine could not complete the turn.")
  expect(stopped.attempts).toEqual({ differences: 1, gaps: 0, verify: 0 })
  expect(erring.prompts.map(prompt => kind(prompt.text)).toSorted()).toEqual(["differences", "verify"])

  const resumed = jobs()
  expect((await run(directory, resumed.reply)).result.status).toBe("done")
  expect(resumed.attempts).toEqual({ differences: 0, gaps: 0, verify: 1 })
  expect(readLeanJson(directory, "discrepancy.json")).toEqual({ entries: [ENTRY], gaps: [GAP, PERIOD], limitations: [] })

  // The same gaps from a new instruction list keep their verify answers.
  writeLeanText(directory, "instructions.json", JSON.stringify({ items: [{ id: "i1", text: "Keep the note to one page.", source_id: "request", quote: "Draft a short supply note" }], documents: [] }))
  writeInstructionsFile(directory)
  const instructed = jobs({ gaps: () => ({ gaps: [GAP, PERIOD], limitations: [] }) })
  const again = await run(directory, instructed.reply)
  expect(again.result.status).toBe("done")
  expect(instructed.attempts).toEqual({ differences: 0, gaps: 1, verify: 0 })
  expect(again.model.prompts[0]!.text).toContain("Keep the note to one page.")

  // A later instruction re-keys every job.
  writeLeanText(directory, "updated-notes.md", "Use the earlier delivery term.")
  const updated = jobs({ gaps: () => ({ gaps: [GAP, PERIOD], limitations: [] }) })
  expect((await run(directory, updated.reply)).result.status).toBe("done")
  expect(updated.attempts).toEqual({ differences: 1, gaps: 1, verify: 2 })
})
