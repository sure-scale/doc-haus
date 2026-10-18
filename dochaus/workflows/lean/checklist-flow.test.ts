import { afterEach, expect, test } from "bun:test"
import { createHash } from "node:crypto"
import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { combineMemoAndBody } from "../../lib/combined-docx"
import { renderMarkdownToDocx } from "../../lib/render"
import { runChecklistCheckFlow, runChecklistExtractionFlow } from "./checklist-flow"
import { CHECKLIST_JOB_AGENTS } from "./checklist-job"
import { writeLeanPlan } from "./plan"
import { memoNotesFile } from "./deliverables"
import { leanFile, readLeanJson, shortHash, writeLeanBytes, writeLeanText } from "./state"
import type { StageSessionPort } from "./stages"
import type { ChecklistCheckerInput, ChecklistItem, ChecklistProducerInput } from "./checklist-contract"
import { buildChecklistDocumentContext } from "./checklist-document-context"
import { checklistProducerInputs } from "./checklist-source-context"
import { matterSourceBasis } from "./quotes"

const dirs: string[] = []
afterEach(() => dirs.splice(0).forEach(directory => rmSync(directory, { recursive: true, force: true })))
function matter(sources = 0) {
  const directory = mkdtempSync(path.join(tmpdir(), "checklist-flow-")); dirs.push(directory)
  writeLeanPlan(directory, [{ title: "Record", purpose: "Record instructions", file: "record.docx", parts: [{ heading: "Record", scope: "Record" }] }])
  writeLeanText(directory, "request.md", "Keep the proposed arrangement subject to confirmation.\n")
  mkdirSync(path.join(directory, "sources"), { recursive: true })
  const rows = Array.from({ length: sources }, (_, i) => ({ document: `source-${i}.txt`, extracted_text: `sources/source-${i}.txt` }))
  for (const row of rows) writeFileSync(path.join(directory, row.extracted_text), "An original source observation.\n")
  writeFileSync(path.join(directory, "source-index.json"), JSON.stringify({ sources: rows }))
  return directory
}
function port(directory: string, options: { failSource?: string; failCheck?: boolean; finding?: "open" | "indeterminate" | "unsatisfied" | "satisfied"; required_action?: "propose"; action_status?: "fulfilled" | "not_fulfilled" | "indeterminate"; delay?: number } = {}) {
  const counts = { prompts: 0, active: 0, peak: 0, created: 0, artifactsSeen: true, incrementalSeen: false }
  const sessions: StageSessionPort = {
    async createSession() {
      counts.created++
      if (counts.created > 4) counts.incrementalSeen ||= ((readLeanJson(directory, "checklist/extraction.json") as { sources?: unknown[] } | undefined)?.sources?.length ?? 0) > 0
      return `session-${counts.created}`
    },
    async prompt(input) {
      counts.prompts++; counts.active++; counts.peak = Math.max(counts.peak, counts.active)
      try {
        await Bun.sleep(options.delay ?? 1)
        if (input.agent === CHECKLIST_JOB_AGENTS.extract) {
          const packet: ChecklistProducerInput = JSON.parse(input.text.split("SOURCE_INPUT:\n\n")[1])
          if (packet.source.document === options.failSource) throw new Error("source worker failed")
          const ref = packet.source.document === "request" ? packet.passages[0] : packet.instructions[0]
          return JSON.stringify({ items: [{ ...(options.required_action ? { required_action: options.required_action } : {}), check: "Check the proposed arrangement is kept open.", basis: "proposal", source_refs: [ref], qualifications: ["Requires confirmation."], target: { ...packet.targets[0], treatment: "open_issue" } }] })
        }
        if (options.failCheck) throw new Error("checker worker failed")
        const packet: ChecklistCheckerInput = JSON.parse(input.text.split("CHECK_INPUT:\n\n")[1])
        counts.artifactsSeen &&= existsSync(leanFile(directory, "checklist-context/d1.json"))
        return JSON.stringify({ item_id: packet.item_id, ...(packet.item.required_action ? { action_status: options.action_status, open_decisions: ["Approval remains open"] } : {}), status: options.finding ?? "open", reason: "The arrangement remains unresolved.", evidence: packet.current_evidence, source_refs: packet.item.source_refs, inspection_scope: "Supplied current target and source.", limitation: "Confirmation remains outstanding." })
      } finally { counts.active-- }
    },
  }
  return { sessions, counts, model: null }
}

test("source completion is incremental, capped at four, exact items and IDs survive cached rerun", async () => {
  const directory = matter(6), deps = port(directory, { delay: 4 })
  const result = await runChecklistExtractionFlow(directory, deps)
  expect(result.status).toBe("succeeded")
  expect(result.sources).toHaveLength(7)
  expect(deps.counts.peak).toBe(4)
  expect(deps.counts.incrementalSeen).toBe(true)
  expect(result.items.every(row => row.item.basis === "proposal")).toBe(true)
  const saved = readLeanJson(directory, "checklist/extraction.json")
  expect(saved).toMatchObject({ expected_sources: result.sources.map(row => row.key) })
  const again = await runChecklistExtractionFlow(directory, deps)
  expect(again.items).toEqual(result.items)
  expect(deps.counts.created).toBe(7)
  writeFileSync(path.join(directory, "sources/source-0.txt"), "Changed original source context.\n")
  await runChecklistExtractionFlow(directory, deps)
  expect(deps.counts.created).toBeGreaterThan(7)
})

test("failed source remains explicit and consumer cannot claim complete coverage", async () => {
  const directory = matter(1), deps = port(directory, { failSource: "source-0.txt" })
  const extracted = await runChecklistExtractionFlow(directory, deps)
  expect(extracted.status).toBe("incomplete")
  expect(extracted.sources.find(row => row.key === "source:source-0.txt")).toMatchObject({ status: "failed", issues: ["source worker failed"], items: [] })
  const checked = await runChecklistCheckFlow(directory, deps)
  expect(checked.extraction_complete).toBe(false)
  expect(checked.status).toBe("incomplete")
  expect(checked.checks).toHaveLength(1)
})

test("current lookup is saved before checks; indeterminate is retained, unchanged context caches, Word changes invalidate", async () => {
  const directory = matter(4), deps = port(directory, { finding: "indeterminate", delay: 4 })
  const extracted = await runChecklistExtractionFlow(directory, deps)
  await writeLeanBytes(directory, "d1/body.docx", (await renderMarkdownToDocx("# Record\n\nThe proposed arrangement remains open.")).bytes)
  const checked = await runChecklistCheckFlow(directory, deps)
  expect(checked.status).toBe("succeeded")
  expect(checked.checks).toHaveLength(extracted.items.length)
  expect(checked.checks.every(row => row.finding?.status === "indeterminate")).toBe(true)
  expect(deps.counts.artifactsSeen).toBe(true)
  expect(deps.counts.peak).toBe(4)
  const before = deps.counts.created
  expect((await runChecklistCheckFlow(directory, deps)).checks).toEqual(checked.checks)
  expect(deps.counts.created).toBe(before)
  writeLeanBytes(directory, "d1/body-applied.docx", (await renderMarkdownToDocx("# Record\n\nThe proposed arrangement has changed but is still open.")).bytes)
  const changed = await runChecklistCheckFlow(directory, deps)
  expect(deps.counts.created).toBe(before + extracted.items.length)
  expect(changed.checks[0].input_sha).not.toBe(checked.checks[0].input_sha)
  expect(JSON.stringify(extracted.items)).toBe(JSON.stringify((readLeanJson(directory, "checklist/extraction.json") as { sources: typeof extracted.sources }).sources.flatMap(row => row.items)))
  rmSync(leanFile(directory, "d1/body.docx")); rmSync(leanFile(directory, "d1/body-applied.docx"))
  await runChecklistCheckFlow(directory, deps)
  expect(readLeanJson(directory, "checklist-context/d1.json")).toMatchObject({ status: "unavailable", paragraphs: [] })
})

test("source update prevents checking stale extracted requirements", async () => {
  const directory = matter(), deps = port(directory)
  await runChecklistExtractionFlow(directory, deps)
  const before = deps.counts.created
  writeLeanText(directory, "request.md", "The request now has a different condition.")
  const checked = await runChecklistCheckFlow(directory, deps)
  expect(checked.status).toBe("incomplete")
  expect(checked.extraction_complete).toBe(false)
  expect(checked.checks).toEqual([])
  expect(deps.counts.created).toBe(before)
})

test("failed check is persisted and only an explicit subsequent flow run retries it", async () => {
  const directory = matter(), deps = port(directory, { failCheck: true })
  await runChecklistExtractionFlow(directory, deps)
  const checked = await runChecklistCheckFlow(directory, deps)
  expect(checked.status).toBe("incomplete")
  expect(checked.checks[0]).toMatchObject({ status: "failed", issues: ["checker worker failed"] })
  expect(deps.counts.created).toBe(2)
  expect(readLeanJson(directory, "checklist/checks.json")).toMatchObject({ checks: checked.checks })
  const retry = port(directory)
  const recovered = await runChecklistCheckFlow(directory, retry)
  expect(recovered.status).toBe("succeeded")
  expect(recovered.checks[0].finding?.status).toBe("open")
  expect(retry.counts.created).toBe(1)
})


test("adverse findings need explicit reinspection, preserve original, and consume one allowance per actual evidence", async () => {
  const directory = matter()
  const options = { finding: "indeterminate" as "indeterminate" | "unsatisfied" | "satisfied" }
  const deps = port(directory, options)
  await runChecklistExtractionFlow(directory, deps)
  writeLeanBytes(directory, "d1/body.docx", (await renderMarkdownToDocx("# Record\n\nThe proposed arrangement remains open.")).bytes)
  const initial = await runChecklistCheckFlow(directory, deps)
  const before = deps.counts.created
  expect((await runChecklistCheckFlow(directory, deps)).checks).toEqual(initial.checks)
  expect(deps.counts.created).toBe(before)
  options.finding = "unsatisfied"
  const reviewed = await runChecklistCheckFlow(directory, deps, { reinspectAdverse: true })
  expect(deps.counts.created).toBe(before + 1)
  expect(reviewed.checks[0].finding?.status).toBe("unsatisfied")
  expect(reviewed.checks[0].reinspection?.original_finding).toEqual(initial.checks[0].finding)
  await runChecklistCheckFlow(directory, deps, { reinspectAdverse: true })
  expect(deps.counts.created).toBe(before + 1)
  // Model/session/prompt cache changes must not replenish substantive review.
  const changedModel = { ...deps, model: { providerID: "test", modelID: "different-model" } }
  await runChecklistExtractionFlow(directory, changedModel)
  await runChecklistCheckFlow(directory, changedModel, { reinspectAdverse: true })
  expect(deps.counts.created).toBe(before + 2)
  await runChecklistCheckFlow(directory, changedModel, { reinspectAdverse: true })
  expect(deps.counts.created).toBe(before + 2)
  // A regenerated paraphrase/session is not new canonical source or Word evidence.
  const extraction = readLeanJson(directory, "checklist/extraction.json") as { sources: { items: { item: { check: string } }[] }[] }
  extraction.sources[0].items[0].item.check = "Rephrased check of the same proposed arrangement."
  writeLeanText(directory, "checklist/extraction.json", JSON.stringify(extraction))
  await runChecklistCheckFlow(directory, changedModel, { reinspectAdverse: true })
  const afterParaphrase = deps.counts.created
  await runChecklistCheckFlow(directory, changedModel, { reinspectAdverse: true })
  expect(deps.counts.created).toBe(afterParaphrase)
  writeLeanBytes(directory, "d1/body.docx", (await renderMarkdownToDocx("# Record\n\nCurrent text has actually changed.")).bytes)
  await runChecklistCheckFlow(directory, changedModel)
  options.finding = "satisfied"
  const revised = await runChecklistCheckFlow(directory, changedModel, { reinspectAdverse: true })
  expect(revised.checks[0].finding?.status).toBe("satisfied")
  expect(revised.checks[0].reinspection?.canonical_input_sha).not.toBe(reviewed.checks[0].reinspection?.canonical_input_sha)
})

/** Evidence identity as computed before annotation properties carried text
 * measures (base 9d4a29764): artifact identity and SHA-256 only. */
function preMeasureCanonicalSha(directory: string, item_id: string, item: ChecklistItem) {
  const current = checklistProducerInputs(directory)
  const request = current.find(job => job.key === "request")!.input
  const settled = [...request.passages, ...request.instructions]
  const context = buildChecklistDocumentContext(directory, { item_id, item, settled_instructions: settled })
  return shortHash(JSON.stringify([matterSourceBasis(directory), item_id,
    current.map(job => job.input.passages.map(ref => ({ document: ref.document, quote: ref.quote }))),
    settled.map(ref => ({ document: ref.document, quote: ref.quote })), context.documents,
    context.artifacts.map(artifact => ({ id: artifact.value.id, document: artifact.value.document, aliases: artifact.value.aliases,
      artifact: artifact.value.artifact, sha256: artifact.value.sha256,
      annotations: artifact.value.annotations?.map(annotation => ({ artifact: annotation.artifact, sha256: annotation.sha256 })) })),
    ...(context.workflow.length ? [context.workflow] : [])]), 40)
}

test("a reinspection consumed before text measures existed is not reused once exact measures are observable", async () => {
  const directory = matter()
  const options = { finding: "indeterminate" as "indeterminate" | "satisfied" }
  const deps = port(directory, options)
  const extracted = await runChecklistExtractionFlow(directory, deps)
  writeLeanBytes(directory, "d1/body.docx", (await renderMarkdownToDocx("# Record\n\nThe proposed arrangement remains open.")).bytes)
  await runChecklistCheckFlow(directory, deps)
  const reviewed = await runChecklistCheckFlow(directory, deps, { reinspectAdverse: true })
  const consumed = reviewed.checks[0]!
  expect(consumed.reinspection?.status).toBe("succeeded")
  // Identical inputs keep the same identity, so the consumed allowance is not renewed.
  const before = deps.counts.created
  const repeated = await runChecklistCheckFlow(directory, deps, { reinspectAdverse: true })
  expect(deps.counts.created).toBe(before)
  expect(repeated.checks[0]!.reinspection!.canonical_input_sha).toBe(consumed.reinspection!.canonical_input_sha)
  const persist = (canonical_input_sha: string) => writeLeanText(directory, "checklist/checks.json", JSON.stringify({ extraction_complete: true, expected_items: [consumed.item_id],
    checks: [{ ...consumed, input_sha: "pre-measure-prompt", reinspection: { ...consumed.reinspection!, canonical_input_sha } }] }))
  // A changed prompt under the same evidence identity reuses the consumed finding without a call.
  persist(consumed.reinspection!.canonical_input_sha)
  expect((await runChecklistCheckFlow(directory, deps, { reinspectAdverse: true })).checks[0]!.finding).toEqual(consumed.finding)
  expect(deps.counts.created).toBe(before)
  // The same record consumed under the pre-measure identity admits exactly one fresh check.
  const legacy = preMeasureCanonicalSha(directory, consumed.item_id, extracted.items[0]!.item)
  expect(legacy).not.toBe(consumed.reinspection!.canonical_input_sha)
  persist(legacy)
  options.finding = "satisfied"
  const refreshed = await runChecklistCheckFlow(directory, deps, { reinspectAdverse: true })
  expect(deps.counts.created).toBe(before + 1)
  expect(refreshed.checks[0]!.finding?.status).toBe("satisfied")
  await runChecklistCheckFlow(directory, deps, { reinspectAdverse: true })
  expect(deps.counts.created).toBe(before + 1)
})

test("a consumed reinspection admits one fresh check once the plan delivers its target in a combined file", async () => {
  const directory = matter()
  const record = { title: "Record", purpose: "Record instructions", file: "record.docx", parts: [{ heading: "Record", scope: "Record" }] }
  const memo = { name: "cover memo", file: memoNotesFile([record], "cover memo") }
  writeLeanPlan(directory, [record], undefined, undefined, memo)
  const options = { finding: "indeterminate" as "indeterminate" | "satisfied" }
  const deps = port(directory, options)
  const checked: ChecklistCheckerInput[] = []
  const prompt = deps.sessions.prompt
  deps.sessions.prompt = async input => {
    if (input.agent === CHECKLIST_JOB_AGENTS.check) checked.push(JSON.parse(input.text.split("CHECK_INPUT:\n\n")[1]))
    return prompt(input)
  }
  await runChecklistExtractionFlow(directory, deps)
  writeLeanBytes(directory, "d1/body.docx", (await renderMarkdownToDocx("# Record\n\nThe proposed arrangement remains open.")).bytes)
  await runChecklistCheckFlow(directory, deps)
  expect((await runChecklistCheckFlow(directory, deps, { reinspectAdverse: true })).checks[0]!.reinspection?.status).toBe("succeeded")
  const before = deps.counts.created
  // Delivered on its own: passages and the record's own files, no combined file.
  expect(checked.flatMap(input => input.current_evidence).every(row => row.kind === "body" || (row.locator.startsWith("delivery_artifacts") && JSON.parse(row.text).kind !== "combined"))).toBe(true)
  // Only the record's delivery changes: it is now built into one file with the same memo.
  writeLeanPlan(directory, [{ ...record, deliverable: "combined" }], undefined, undefined, memo)
  options.finding = "satisfied"
  const refreshed = await runChecklistCheckFlow(directory, deps, { reinspectAdverse: true })
  expect(deps.counts.created).toBe(before + 1)
  expect(refreshed.checks[0]!.finding?.status).toBe("satisfied")
  expect(JSON.parse(checked.at(-1)!.current_evidence.at(-1)!.text)).toMatchObject({ kind: "combined", download_filename: "record.docx",
    contains: [{ order: 1, planned_filename: "record - cover memo.docx" }, { order: 2, planned_filename: "record - clean.docx" }] })
  await runChecklistCheckFlow(directory, deps, { reinspectAdverse: true })
  expect(deps.counts.created).toBe(before + 1)
})

test("a consumed reinspection admits one fresh check when only the combined file built from the target's parts changes", async () => {
  const directory = matter()
  const record = { title: "Record", purpose: "Record instructions", file: "record.docx", deliverable: "combined" as const, parts: [{ heading: "Record", scope: "Record" }] }
  writeLeanPlan(directory, [record], undefined, undefined, { name: "cover memo", file: memoNotesFile([record], "cover memo") })
  const options = { finding: "indeterminate" as "indeterminate" | "satisfied" }
  const deps = port(directory, options)
  const checked: ChecklistCheckerInput[] = []
  const prompt = deps.sessions.prompt
  deps.sessions.prompt = async input => {
    if (input.agent === CHECKLIST_JOB_AGENTS.check) checked.push(JSON.parse(input.text.split("CHECK_INPUT:\n\n")[1]))
    return prompt(input)
  }
  const docx = async (markdown: string) => (await renderMarkdownToDocx(markdown)).bytes
  await runChecklistExtractionFlow(directory, deps)
  writeLeanBytes(directory, "d1/body.docx", await docx("# Record\n\nThe proposed arrangement remains open."))
  const memo = await docx("# Cover memo\n\nThe proposed arrangement is summarised.")
  writeLeanBytes(directory, "drafting-notes.docx", memo)
  // Not an accepted apply result, so the checked passages stay those of the draft.
  writeLeanBytes(directory, "d1/body-applied.docx", await docx("# Record\n\nThe proposed arrangement remains open."))
  await runChecklistCheckFlow(directory, deps)
  expect((await runChecklistCheckFlow(directory, deps, { reinspectAdverse: true })).checks[0]!.reinspection?.status).toBe("succeeded")
  const before = deps.counts.created
  const passages = checked.at(-1)!.current_evidence.filter(row => row.kind === "body")
  expect(passages.length).toBeGreaterThan(0)
  // Only the combined file's body part changes; every passage the checker saw stays the same.
  const body = await docx("# Record\n\nThe proposed arrangement remains open, subject to confirmation.")
  writeLeanBytes(directory, "d1/body-applied.docx", body)
  options.finding = "satisfied"
  const refreshed = await runChecklistCheckFlow(directory, deps, { reinspectAdverse: true })
  expect(deps.counts.created).toBe(before + 1)
  expect(refreshed.checks[0]!.finding?.status).toBe("satisfied")
  const latest = checked.at(-1)!
  expect(latest.current_evidence.filter(row => row.kind === "body")).toEqual(passages)
  expect(JSON.parse(latest.current_evidence.at(-1)!.text).built_at_check).toEqual({
    sha256: createHash("sha256").update(combineMemoAndBody({ memo, body })).digest("hex"), nonempty_paragraphs: 4, paragraphs_equal_memo_then_body: true })
  await runChecklistCheckFlow(directory, deps, { reinspectAdverse: true })
  expect(deps.counts.created).toBe(before + 1)
})

test("failed reinspection preserves original adverse finding and does not retry provider on Continue", async () => {
  const directory = matter()
  const options = { finding: "indeterminate" as const, failCheck: false }
  const deps = port(directory, options)
  await runChecklistExtractionFlow(directory, deps)
  const initial = await runChecklistCheckFlow(directory, deps)
  options.failCheck = true
  const reviewed = await runChecklistCheckFlow(directory, deps, { reinspectAdverse: true })
  expect(reviewed.checks[0].finding).toEqual(initial.checks[0].finding)
  expect(reviewed.checks[0].reinspection).toMatchObject({ status: "failed", issues: ["checker worker failed"] })
  const calls = deps.counts.created
  await runChecklistCheckFlow(directory, deps, { reinspectAdverse: true })
  expect(deps.counts.created).toBe(calls)
})

test("a reinspection a pause interrupted spends nothing and runs again; its recorded outcome is never repeated", async () => {
  const directory = matter()
  const options = { finding: "indeterminate" as "indeterminate" | "unsatisfied" }
  const deps = port(directory, options)
  await runChecklistExtractionFlow(directory, deps)
  const initial = await runChecklistCheckFlow(directory, deps)
  const pause = new AbortController()
  const asked = Promise.withResolvers<void>()
  // As the runner's guard does, a pause rejects the pending prompt at once.
  const held = { ...deps, sessions: { ...deps.sessions, prompt: () => {
    asked.resolve()
    return new Promise<string>((_, reject) => pause.signal.addEventListener("abort", () => reject(new Error("Paused at your request.")), { once: true }))
  } } }
  const interrupted = runChecklistCheckFlow(directory, held, { reinspectAdverse: true, pause: pause.signal })
  await asked.promise
  pause.abort()
  const paused = (await interrupted).checks[0]!
  expect(paused.finding).toEqual(initial.checks[0]!.finding)
  expect(paused.reinspection).toMatchObject({ status: "failed", original_finding: initial.checks[0]!.finding })
  expect(paused.reinspection?.sessionID).toBeUndefined()
  // The Continue after a pause does not request reinspection; the unspent admission resumes.
  options.finding = "unsatisfied"
  const before = deps.counts.created
  const resumed = await runChecklistCheckFlow(directory, deps)
  expect(deps.counts.created).toBe(before + 1)
  expect(resumed.checks[0]!.finding?.status).toBe("unsatisfied")
  expect(resumed.checks[0]!.reinspection).toMatchObject({ status: "succeeded", original_finding: initial.checks[0]!.finding })
  await runChecklistCheckFlow(directory, deps, { reinspectAdverse: true })
  await runChecklistCheckFlow(directory, deps)
  expect(deps.counts.created).toBe(before + 1)
})

test("refining a selected parent preserves original finding and exact extraction while checking children", async () => {
  const { runChecklistRefinementFlow, reconcileChecklistParentCoverage } = await import("./checklist-flow")
  const directory = matter(), base = port(directory)
  const extracted = await runChecklistExtractionFlow(directory, base)
  await writeLeanBytes(directory, "d1/body.docx", (await renderMarkdownToDocx("The proposed arrangement remains subject to confirmation.")).bytes)
  const initial = await runChecklistCheckFlow(directory, base)
  const original = readLeanJson(directory, "checklist/extraction.json")
  const parent = extracted.items[0]!
  const sessions: StageSessionPort = {
    createSession: base.sessions.createSession.bind(base.sessions),
    async prompt(input) {
      if (input.text.includes("REFINEMENT_INPUT:\n\n")) {
        const packet = JSON.parse(input.text.split("REFINEMENT_INPUT:\n\n")[1]!)
        return JSON.stringify({ parent_id: packet.parent_id, unit_id: packet.source_unit.id, disposition: "refined", reason: "The source preserves a proposal.", items: [{ ...parent.item, required_action: "propose" }], source_refs: parent.item.source_refs, limitations: [] })
      }
      const reply = JSON.parse(await base.sessions.prompt(input))
      return JSON.stringify({ ...reply, action_status: "fulfilled", open_decisions: ["Confirmation"] })
    },
  }
  await runChecklistRefinementFlow(directory, { sessions, model: null }, { parent_id: parent.id, scope: "cited_sources", units: [{ id: "source-a", role: "Instruction", passages: parent.item.source_refs }], instructions: parent.item.source_refs, targets: [{ document: "d1", role: "Record" }], retrieval: "Use exact source." })
  const checked = await runChecklistCheckFlow(directory, { sessions, model: null })
  expect(checked.checks).toHaveLength(1)
  expect(checked.checks[0]!.item_id).not.toBe(parent.id)
  expect((readLeanJson(directory, "checklist/checks.json") as { checks: unknown[] }).checks).toContainEqual(initial.checks[0])
  expect(readLeanJson(directory, "checklist/extraction.json")).toEqual(original)
  expect(reconcileChecklistParentCoverage(directory, checked.checks)[0]!.status).toBe("open")
})


for (const action of ["not_fulfilled", "indeterminate", "missing", "fulfilled"] as const) test(`refined open/${action} uses the existing canonical reinspection budget only when action evidence is adverse`, async () => {
  const directory = matter()
  const options = { finding: "open" as const, required_action: "propose" as const, action_status: action === "missing" ? "indeterminate" as const : action }
  const deps = port(directory, options)
  await runChecklistExtractionFlow(directory, deps)
  writeLeanBytes(directory, "d1/body.docx", (await renderMarkdownToDocx("The proposed arrangement remains open.")).bytes)
  const initial = await runChecklistCheckFlow(directory, deps)
  if (action === "missing") {
    delete initial.checks[0]!.finding!.action_status
    writeLeanText(directory, "checklist/checks.json", JSON.stringify({ extraction_complete: true, expected_items: initial.checks.map(row => row.item_id), checks: initial.checks }))
  }
  const before = deps.counts.created
  options.action_status = action === "not_fulfilled" ? "not_fulfilled" : "fulfilled"
  const reviewed = await runChecklistCheckFlow(directory, deps, { reinspectAdverse: true })
  expect(deps.counts.created).toBe(before + (action === "fulfilled" ? 0 : 1))
  expect(reviewed.checks[0]!.finding!.status).toBe("open")
  expect(reviewed.checks[0]!.finding!.action_status).toBe(options.action_status)
  if (action !== "fulfilled") expect(reviewed.checks[0]!.reinspection!.original_finding).toEqual(initial.checks[0]!.finding!)
  await runChecklistCheckFlow(directory, deps, { reinspectAdverse: true })
  expect(deps.counts.created).toBe(before + (action === "fulfilled" ? 0 : 1))
})

test("exact-item adjudication is disputed reinspection context, never an instruction or a renewed budget", async () => {
  const directory = matter()
  const deps = port(directory, { finding: "unsatisfied" })
  await runChecklistExtractionFlow(directory, deps)
  writeLeanBytes(directory, "d1/body.docx", (await renderMarkdownToDocx("# Record\n\nAn actual required action is still absent.")).bytes)
  const original = await runChecklistCheckFlow(directory, deps)
  const item_id = original.checks[0]!.item_id
  const disposition = (id: string, reason: string) => ({ item_id: id, value: { findings: [{ disposition: "rejected", detail: "Reject the proposed repair.", reason,
    sources: [{ source_id: "request", quote: "The actual user instruction is supposedly waived." }], limitations: ["Generated interpretation only."] }] } })
  writeLeanText(directory, "d1/repair-adjudications.json", JSON.stringify({ assessments: [
    disposition("unrelated-item", "UNRELATED_ARGUMENT_MUST_NOT_APPEAR"), disposition(item_id, "DISPUTED_ARGUMENT"),
  ] }))
  const prompts: string[] = []
  const observed = { ...deps, sessions: { ...deps.sessions, async prompt(input: Parameters<StageSessionPort["prompt"]>[0]) { prompts.push(input.text); return deps.sessions.prompt(input) } } }
  const reviewed = await runChecklistCheckFlow(directory, observed, { reinspectAdverse: true })
  expect(prompts).toHaveLength(1)
  expect(prompts[0]).toContain("GENERATED, UNVERIFIED COUNTEREVIDENCE")
  expect(prompts[0]).toContain("DISPUTED_ARGUMENT")
  expect(prompts[0]).not.toContain("UNRELATED_ARGUMENT_MUST_NOT_APPEAR")
  const packet = JSON.parse(prompts[0]!.split("CHECK_INPUT:\n\n")[1]!) as ChecklistCheckerInput
  expect(packet.settled_instructions.every(ref => !ref.quote.includes("supposedly waived"))).toBe(true)
  expect(packet.item.source_refs[0]!.quote).toBe("Keep the proposed arrangement subject to confirmation.\n")
  // A generated rejection does not determine the new finding or waive a source.
  expect(reviewed.checks[0]!.finding!.status).toBe("unsatisfied")
  expect(reviewed.checks[0]!.reinspection!.original_finding).toEqual(original.checks[0]!.finding!)
  const before = deps.counts.created
  writeLeanText(directory, "d1/repair-adjudications.json", JSON.stringify({ assessments: [disposition(item_id, "A different generated conclusion")] }))
  const resumed = await runChecklistCheckFlow(directory, observed, { reinspectAdverse: true })
  expect(deps.counts.created).toBe(before)
  expect(resumed.checks[0]!.reinspection).toEqual(reviewed.checks[0]!.reinspection)
  expect(resumed.checks[0]!.input_sha).toBe(original.checks[0]!.input_sha)
})
