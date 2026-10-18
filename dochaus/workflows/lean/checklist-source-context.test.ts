import { afterEach, expect, test } from "bun:test"
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { checklistProducerInputs } from "./checklist-source-context"
import { runChecklistExtractionFlow } from "./checklist-flow"
import type { StageSessionPort } from "./stages"
import { checklistProducerInputSchema } from "./checklist-contract"

const directories: string[] = []
const workspace = process.env.WORKSPACE_ROOT
afterEach(() => {
  directories.splice(0).forEach(directory => rmSync(directory, { recursive: true, force: true }))
  if (workspace === undefined) delete process.env.WORKSPACE_ROOT
  else process.env.WORKSPACE_ROOT = workspace
})

function prepare() {
  const directory = mkdtempSync(path.join(tmpdir(), "checklist-source-context-"))
  directories.push(directory)
  const put = (relative: string, text: string) => {
    mkdirSync(path.dirname(path.join(directory, relative)), { recursive: true })
    writeFileSync(path.join(directory, relative), text)
  }
  put(".dochaus/lean/plan.json", JSON.stringify({ schema: "dochaus.lean.plan/1", documents: [
    { id: "d1", title: "Agreement", purpose: "Record agreed terms", parts: [{ heading: "Terms", scope: "Terms" }], file: "agreement.docx" },
    { id: "d2", title: "Advice", purpose: "Explain decisions", parts: [{ heading: "Advice", scope: "Decisions" }], file: "advice.docx" },
  ] }))
  put(".dochaus/lean/request.md", "Review the supplied documents.\nKeep unconfirmed proposals open.\n")
  put("source-index.json", JSON.stringify({ source_ledger: "source-ledger.json", sources: [
    { document: "instructions.txt", extracted_text: "sources/instructions.txt", characters: 41 },
    { document: "draft.docx", extracted_text: "sources/draft.txt", characters: 38 },
  ] }))
  put("sources/instructions.txt", "  Proposed arrangement — subject to confirmation.\n\n")
  put("sources/draft.txt", "# Terms\n\nOriginal draft wording.\n")
  return { directory, put }
}

test("one job per original source plus request uses exact source bytes and planned destinations", () => {
  const { directory } = prepare()
  const jobs = checklistProducerInputs(directory)
  expect(jobs.map(job => job.key)).toEqual(["source:instructions.txt", "source:draft.docx", "request"])
  expect(jobs[0].input.passages[0].quote).toBe(readFileSync(path.join(directory, "sources/instructions.txt"), "utf8"))
  expect(jobs[1].input.passages[0].quote).toBe(readFileSync(path.join(directory, "sources/draft.txt"), "utf8"))
  expect(jobs[2].input.passages[0].quote).toBe(readFileSync(path.join(directory, ".dochaus/lean/request.md"), "utf8"))
  expect(jobs[0].input.targets.slice(0, 2)).toEqual([{ document: "d1", role: "Agreement: Record agreed terms" }, { document: "d2", role: "Advice: Explain decisions" }])
  expect(jobs[0].input.targets[2].document).toBe("notes")
  for (const job of jobs) expect(checklistProducerInputSchema.safeParse(job.input).success).toBe(true)
  expect(JSON.stringify(jobs)).not.toContain('"pid"')
})

test("included lawyer updates and actual answers stay separate from original text and model questions", () => {
  const { directory, put } = prepare()
  put(".dochaus/lean/questions.json", JSON.stringify({ questions: [
    { id: "q1", question: "MODEL QUESTION MUST NOT BECOME SOURCE", status: "answered", answer: "Keep this proposal unapproved pending confirmation." },
    { id: "q2", question: "OPEN MODEL QUESTION", status: "open", answer: "UNSETTLED TEXT" },
  ] }))
  put(".dochaus/lean/requests/done/update1.json", JSON.stringify({ kind: "update", id: "update1", text: "Use the current requested document names.", outcome: { status: "included" } }))
  put(".dochaus/lean/requests/queued.json", JSON.stringify({ kind: "update", id: "queued", text: "QUEUED TEXT" }))
  const jobs = checklistProducerInputs(directory)
  const request = jobs.at(-1)!.input
  expect(request.passages).toHaveLength(1)
  expect(request.instructions.map(ref => ref.quote)).toEqual(["Use the current requested document names.", "Keep this proposal unapproved pending confirmation."])
  expect(jobs[0].input.instructions.map(ref => ref.quote)).toEqual([request.passages[0].quote, ...request.instructions.map(ref => ref.quote)])
  expect(JSON.stringify(jobs)).not.toContain("MODEL QUESTION")
  expect(JSON.stringify(jobs)).not.toContain("UNSETTLED TEXT")
  expect(JSON.stringify(jobs)).not.toContain("QUEUED TEXT")
})

test("each answer keeps an id-keyed locator in a stable order however questions.json is ordered", () => {
  const { directory, put } = prepare()
  const answered = (id: string) => ({ id, question: `Question ${id}?`, status: "answered", answer: `Answer ${id}.` })
  const open = { id: "q0", question: "Open question?", status: "open" }
  put(".dochaus/lean/questions.json", JSON.stringify({ questions: [open, answered("q2"), answered("q1")] }))
  const before = checklistProducerInputs(directory)
  expect(before.at(-1)!.input.instructions).toEqual([
    { document: "request", locator: ".dochaus/lean/questions.json answer (q1)", quote: "Answer q1." },
    { document: "request", locator: ".dochaus/lean/questions.json answer (q2)", quote: "Answer q2." },
  ])
  // Answered questions first, as a step's question rewrite leaves them.
  put(".dochaus/lean/questions.json", JSON.stringify({ questions: [answered("q1"), answered("q2"), open] }))
  expect(checklistProducerInputs(directory)).toEqual(before)
})

test("generated artifacts and drafting assumptions never become producer evidence", () => {
  const { directory, put } = prepare()
  const baseline = checklistProducerInputs(directory)
  put(".dochaus/lean/d1/body.docx", "GENERATED OUTPUT")
  put(".dochaus/lean/drafting-notes.md", "GENERATED ADVICE")
  put(".dochaus/lean/spine.json", JSON.stringify({ assumptions: ["MODEL ASSUMPTION"] }))
  put(".dochaus/lean/updated-notes.md", "MODEL REWRITE")
  expect(checklistProducerInputs(directory)).toEqual(baseline)
})

test("missing plan, canonical request or indexed source text produces explicit preparation failure", () => {
  const { directory, put } = prepare()
  const plan = readFileSync(path.join(directory, ".dochaus/lean/plan.json"), "utf8")
  rmSync(path.join(directory, ".dochaus/lean/plan.json"))
  expect(() => checklistProducerInputs(directory)).toThrow("document plan")
  put(".dochaus/lean/plan.json", plan)
  put(".dochaus/lean/request.md", "  ")
  expect(() => checklistProducerInputs(directory)).toThrow("canonical lawyer request")
  put(".dochaus/lean/request.md", "Review the supplied sources.")
  put("sources/instructions.txt", "")
  expect(() => checklistProducerInputs(directory)).toThrow("instructions.txt has no extracted text")
})

test("request-only matter retains the real workflow notes destination without inventing sources", () => {
  const { directory } = prepare()
  rmSync(path.join(directory, "source-index.json"))
  const jobs = checklistProducerInputs(directory)
  expect(jobs.map(job => job.key)).toEqual(["request"])
  expect(jobs[0].input.targets.map(target => target.document)).toEqual(["d1", "d2", "notes"])
})


test("selected library playbook adds exact source job and library changes invalidate cached extraction", async () => {
  const { directory, put } = prepare()
  process.env.WORKSPACE_ROOT = directory
  const unbound = checklistProducerInputs(directory)
  const text = "# Firm guidance\nPreserve conditional fallback positions.\n"
  put(".playbooks/playbook-local-guidance/SKILL.md", text)
  expect(checklistProducerInputs(directory)).toEqual(unbound)
  put("matter.json", JSON.stringify({ playbook: "playbook-local-guidance" }))
  const bound = checklistProducerInputs(directory)
  const guidance = bound.find(job => job.key === "source:playbook:playbook-local-guidance")!
  expect(bound).toHaveLength(unbound.length + 1)
  expect(guidance.input.source.document).toBe("playbook:playbook-local-guidance")
  expect(guidance.input.passages).toEqual([{ document: "playbook:playbook-local-guidance", locator: "read-document playbook:playbook-local-guidance lines 1-3", quote: text }])
  expect(guidance.input.instructions).toEqual(bound[0].input.instructions)
  expect(guidance.input.targets).toEqual(bound[0].input.targets)
  expect(readFileSync(path.join(directory, "source-index.json"), "utf8")).not.toContain("playbook:")
  const calls: string[] = []
  const sessions: StageSessionPort = {
    async createSession() { return `session-${calls.length}` },
    async prompt(input) { calls.push(input.text); return JSON.stringify({ items: [] }) },
  }
  const first = await runChecklistExtractionFlow(directory, { sessions, model: null })
  expect(first.status).toBe("succeeded")
  expect(first.sources.some(row => row.key === guidance.key)).toBe(true)
  expect(calls).toHaveLength(4)
  await runChecklistExtractionFlow(directory, { sessions, model: null })
  expect(calls).toHaveLength(4)
  put(".playbooks/playbook-local-guidance/SKILL.md", `${text}Keep exceptions explicit.\n`)
  const changed = await runChecklistExtractionFlow(directory, { sessions, model: null })
  expect(calls).toHaveLength(8)
  expect(changed.sources.find(row => row.key === guidance.key)!.input_sha).not.toBe(first.sources.find(row => row.key === guidance.key)!.input_sha)
  put("matter.json", JSON.stringify({ playbook: "playbook-other-guidance" }))
  put(".playbooks/playbook-other-guidance/SKILL.md", text)
  const rebound = await runChecklistExtractionFlow(directory, { sessions, model: null })
  expect(rebound.sources.some(row => row.key === "source:playbook:playbook-other-guidance")).toBe(true)
  expect(rebound.sources.some(row => row.key === guidance.key)).toBe(false)
})

test("unavailable, invalid or empty selected playbooks surface explicit source coverage failure", () => {
  const { directory, put } = prepare()
  process.env.WORKSPACE_ROOT = directory
  expect(checklistProducerInputs(directory)).toHaveLength(3)
  put("matter.json", JSON.stringify({ playbook: "playbook-local-guidance" }))
  expect(() => checklistProducerInputs(directory)).toThrow("selected playbook playbook:playbook-local-guidance")
  put(".playbooks/playbook-local-guidance/SKILL.md", "  ")
  expect(() => checklistProducerInputs(directory)).toThrow("selected playbook has no text")
  put("matter.json", JSON.stringify({ playbook: "../other-guidance" }))
  expect(() => checklistProducerInputs(directory)).toThrow("Invalid bound playbook name")
})
