import { assignChecklistTargetContext, checklistTargetAssignments } from "./checklist-target-context"
import { checklistComparisonAssignmentSchema, type ChecklistComparisonAssignment } from "./checklist-contract"
import { leanClarifications } from "./requests"
import { z } from "zod"
import { readLeanPlan } from "./plan"
import { checklistAdjudicationCounterevidenceSchema, type ChecklistAdjudicationCounterevidence, type ChecklistItem } from "./checklist-contract"
import { checklistFindingRecovery, buildChecklistCheckerPrompt, buildChecklistProducerPrompt, checklistFindingSchema, checklistItemSchema, checklistRefinementSchema, buildChecklistRefinementPrompt, type ChecklistRefinementInput, type ChecklistSourceReference, type ChecklistTargetDocument } from "./checklist-contract"
import { buildChecklistDocumentContext } from "./checklist-document-context"
import { runChecklistCheckJob, runChecklistExtractionJob, runChecklistRefinementJob, type ChecklistJobDeps } from "./checklist-job"
import { checklistProducerInputs } from "./checklist-source-context"
import { matterSourceBasis } from "./quotes"
import { readLeanJson, shortHash, writeLeanText } from "./state"

const itemRecordSchema = z.object({ id: z.string(), item: checklistItemSchema, assignment: checklistComparisonAssignmentSchema.optional() })
const sourceResultSchema = z.object({
  key: z.string(), input_sha: z.string(), status: z.enum(["succeeded", "failed"]),
  items: z.array(itemRecordSchema), issues: z.array(z.string()), sessionID: z.string().optional(), attempts: z.number().optional(),
})
const extractionSchema = z.object({ expected_sources: z.array(z.string()), sources: z.array(sourceResultSchema) })
const reinspectionSchema = z.object({
  canonical_input_sha: z.string(), original_finding: checklistFindingSchema,
  status: z.enum(["succeeded", "failed"]),
  issues: z.array(z.string()), sessionID: z.string().optional(), attempts: z.number().optional(),
})
export const checklistCheckResultSchema = z.object({
  item_id: z.string(), input_sha: z.string(), status: z.enum(["succeeded", "failed"]),
  finding: checklistFindingSchema.optional(), reinspection: reinspectionSchema.optional(), issues: z.array(z.string()), sessionID: z.string().optional(), attempts: z.number().optional(),
})
const checksSchema = z.object({ extraction_complete: z.boolean(), expected_items: z.array(z.string()), checks: z.array(checklistCheckResultSchema) })
export type ChecklistItemRecord = z.infer<typeof itemRecordSchema>
export type ChecklistSourceResult = z.infer<typeof sourceResultSchema>
export type ChecklistCheckResult = z.infer<typeof checklistCheckResultSchema>

const EXTRACTION = "checklist/extraction.json"
const CHECKS = "checklist/checks.json"

/** Read-only product selection; never infer an adverse parent from prose. */
export function checklistRefinableParents(directory: string) {
  try {
    const extraction = extractionSchema.safeParse(readLeanJson(directory, EXTRACTION))
    const checks = checksSchema.safeParse(readLeanJson(directory, CHECKS))
    if (!extraction.success || !checks.success) return []
    return extraction.data.sources.flatMap(source => source.status === "succeeded" ? source.items : []).flatMap(parent => {
      const matching = checks.data.checks.filter(row => row.item_id === parent.id && row.status === "succeeded")
      const finding = matching.length === 1 ? matching[0]!.finding : undefined
      if (!finding || (finding.status !== "unsatisfied" && finding.status !== "indeterminate")) return []
      return [{ parent_id: parent.id, check: parent.item.check, target: parent.item.target,
        finding: { status: finding.status, reason: finding.reason }, scopes: ["target_body" as const] }]
    })
  } catch { return [] }
}

/** The existing runner remains the sole caller/writer. Only completed job
 * outputs are cached; an explicit rerun retries failures with unchanged input. */
export async function runChecklistExtractionFlow(directory: string, deps: ChecklistJobDeps) {
  const jobs = extractionInputs(directory, deps)
  const prior = readLeanJson(directory, EXTRACTION)
  const cached = prior === undefined ? undefined : extractionSchema.parse(prior)
  const records: (ChecklistSourceResult | undefined)[] = jobs.map(job => cached?.sources.find(row => row.key === job.key && row.input_sha === job.input_sha))
  const save = () => writeLeanText(directory, EXTRACTION, `${JSON.stringify({ expected_sources: jobs.map(job => job.key), sources: records.filter((row): row is ChecklistSourceResult => !!row) }, null, 2)}\n`)
  await parallel(jobs, async (job, index) => {
    if (records[index]?.status === "succeeded") return
    try {
      const result = await runChecklistExtractionJob(directory, deps, job.input)
      records[index] = result.status === "succeeded"
        ? { key: job.key, input_sha: job.input_sha, status: "succeeded", items: result.value.items.map((item, ordinal) => ({ id: `source-${shortHash(job.key, 12)}-${ordinal + 1}`, item })), issues: [], sessionID: result.sessionID, attempts: result.attempts }
        : { key: job.key, input_sha: job.input_sha, status: "failed", items: [], issues: result.issues, sessionID: result.sessionID, attempts: result.attempts }
    } catch (error) {
      records[index] = { key: job.key, input_sha: job.input_sha, status: "failed", items: [], issues: [message(error)] }
    }
    save()
  })
  save()
  const sources = records.filter((row): row is ChecklistSourceResult => !!row)
  return { status: sources.length === jobs.length && sources.every(row => row.status === "succeeded") ? "succeeded" as const : "incomplete" as const, sources, items: sources.flatMap(row => row.items), outputs: [EXTRACTION] }
}

/** Each unchanged item gets its own native check. Full current lookup records
 * are persisted before any sessions start, while only target hints go inline. */
export async function runChecklistCheckFlow(directory: string, deps: ChecklistJobDeps, options: { reinspectAdverse?: boolean } = {}) {
  const current = extractionInputs(directory, deps)
  const saved = readLeanJson(directory, EXTRACTION)
  if (saved === undefined) throw new Error("Checklist extraction has no saved output; run source extraction before checking.")
  const extraction = extractionSchema.parse(saved)
  const valid = current.map(job => extraction.sources.find(row => row.key === job.key && row.input_sha === job.input_sha))
  const extractionComplete = valid.every(row => row?.status === "succeeded")
  const originals = valid.flatMap(row => row?.status === "succeeded" ? row.items : [])
  const effective = readEffectiveChecklistItems(directory, originals)
  const items = effective.items
  const instructions = current.find(job => job.key === "request")!.input
  const clarifications = leanClarifications(directory)
  const settled = [...instructions.passages, ...instructions.instructions, ...clarifications]
  const contexts = items.map(record => {
    const context = buildChecklistDocumentContext(directory, { item_id: record.id, item: record.item, settled_instructions: settled,
    ...(clarifications.length ? { retrieval: "Later human clarification is supplied with its original queue locator. It clarifies source relationships or applicability for inspection only; it does not change the item, current Word evidence, or adopt a proposal as approved. A statement that text is present is not proof of the artifact: inspect it. If the clarification calls for a material document change, retain the adverse finding and describe the supported correction needed; this inspection cannot implement it." } : {}) })
    return { record, context: record.assignment ? assignChecklistTargetContext(context, record.assignment) : context }
  })
  const artifacts = new Map(contexts.flatMap(row => row.context.artifacts.map(artifact => [artifact.path, artifact] as const)))
  for (const artifact of artifacts.values()) {
    if (!artifact.path.startsWith(".dochaus/lean/")) throw new Error("Checklist lookup records must be lean output artifacts.")
    writeLeanText(directory, artifact.path.slice(".dochaus/lean/".length), `${JSON.stringify(artifact.value, null, 2)}\n`)
  }
  const missing = new Map(contexts.flatMap(row => row.context.documents.filter(document => document.status === "unavailable").map(document => [document.path, document] as const)))
  for (const [file, document] of missing) writeLeanText(directory, file.slice(".dochaus/lean/".length), `${JSON.stringify({ id: document.id, document: document.document, status: "unavailable", paragraphs: [] }, null, 2)}\n`)
  const sourceBasis = matterSourceBasis(directory)
  const jobs = contexts.map(row => ({ ...row,
    // Review allowance follows actual evidence, never a prompt/model/session revision.
    canonical_input_sha: shortHash(JSON.stringify([sourceBasis, row.record.id,
      current.map(job => job.input.passages.map(ref => ({ document: ref.document, quote: ref.quote }))),
      settled.map(ref => ({ document: ref.document, quote: ref.quote })), row.context.documents,
      row.context.artifacts.map(artifact => ({ id: artifact.value.id, document: artifact.value.document, aliases: artifact.value.aliases,
        artifact: artifact.value.artifact, sha256: artifact.value.sha256,
        annotations: artifact.value.annotations?.map(annotation => ({ artifact: annotation.artifact, sha256: annotation.sha256 })) })),
      ...(row.context.workflow.length ? [row.context.workflow] : []),
      ...(row.record.assignment ? [row.record.assignment] : [])]), 40),
    input_sha: shortHash(JSON.stringify(["checklist-check-1", deps.model, sourceBasis, buildChecklistCheckerPrompt(row.context.input), row.context.artifacts, row.context.workflow]), 40) }))
  const prior = readLeanJson(directory, CHECKS)
  const cached = prior === undefined ? undefined : checksSchema.parse(prior)
  const records: (ChecklistCheckResult | undefined)[] = jobs.map(job => cached?.checks.find(row => row.item_id === job.record.id && row.input_sha === job.input_sha))
  const save = () => writeLeanText(directory, CHECKS, `${JSON.stringify({ extraction_complete: extractionComplete, expected_items: jobs.map(job => job.record.id), checks: [...(cached?.checks.filter(row => !jobs.some(job => job.record.id === row.item_id)) ?? []), ...records.filter((row): row is ChecklistCheckResult => !!row)] }, null, 2)}\n`)
  await parallel(jobs, async (job, index) => {
    const priorReview = cached?.checks.find(row => row.item_id === job.record.id)?.reinspection
    const reinspection = priorReview?.canonical_input_sha === job.canonical_input_sha ? priorReview : undefined
    const record = records[index]
    const previous = cached?.checks.find(row => row.item_id === job.record.id)
    // A changed prompt or navigation hint is not new substantive evidence.
    // Preserve the consumed assessment before either ordinary or clarified admission.
    if (!record && reinspection && previous?.status === "succeeded" && previous.finding) {
      records[index] = { ...previous, input_sha: job.input_sha }
      save()
      return
    }
    const clarified = !record && clarifications.length > 0 && previous?.status === "succeeded" && previous.finding ? previous : undefined
    const inspected = record ?? clarified
    if (inspected?.status === "succeeded") {
      if (!clarified && (!options.reinspectAdverse || !inspected.finding || checklistFindingRecovery(job.record.item, inspected.finding) === "none" || reinspection)) return
      if (!inspected.finding) return
      // Consume the one permitted assessment in this existing completed-output
      // record before admission. Interruption remains visible and cannot loop.
      const review = { canonical_input_sha: job.canonical_input_sha, original_finding: inspected.reinspection?.original_finding ?? inspected.finding,
        status: "failed" as const, issues: ["Reinspection did not complete; no result overrides the original finding."] }
      records[index] = { ...inspected, input_sha: job.input_sha, reinspection: review }
      save()
      try {
        const result = await runChecklistCheckJob(directory, deps, job.context.input, inspected.finding, readChecklistAdjudicationCounterevidence(directory, job.record.id, job.record.item))
        records[index] = result.status === "succeeded"
          ? { ...inspected, input_sha: job.input_sha, finding: result.value, reinspection: { ...review, status: "succeeded", issues: [], sessionID: result.sessionID, attempts: result.attempts } }
          : { ...inspected, input_sha: job.input_sha, reinspection: { ...review, issues: result.issues, sessionID: result.sessionID, attempts: result.attempts } }
      } catch (error) {
        records[index] = { ...inspected, input_sha: job.input_sha, reinspection: { ...review, issues: [message(error)] } }
      }
      save()
      return
    }
    try {
      const result = await runChecklistCheckJob(directory, deps, job.context.input)
      records[index] = result.status === "succeeded"
        ? { item_id: job.record.id, input_sha: job.input_sha, status: "succeeded", finding: result.value, issues: [], sessionID: result.sessionID, attempts: result.attempts, ...(reinspection ? { reinspection } : {}) }
        : { item_id: job.record.id, input_sha: job.input_sha, status: "failed", issues: result.issues, sessionID: result.sessionID, attempts: result.attempts, ...(reinspection ? { reinspection } : {}) }
    } catch (error) {
      records[index] = { item_id: job.record.id, input_sha: job.input_sha, status: "failed", issues: [message(error)], ...(reinspection ? { reinspection } : {}) }
    }
    save()
  })
  save()
  const checks = records.filter((row): row is ChecklistCheckResult => !!row)
  const parent_coverage = reconcileChecklistParentCoverage(directory, checks, originals)
  return { parent_coverage, refinement_complete: effective.complete, status: extractionComplete && effective.complete && checks.length === jobs.length && checks.every(row => row.status === "succeeded") ? "succeeded" as const : "incomplete" as const, extraction_complete: extractionComplete, checks, outputs: [CHECKS, ...[...new Set([...artifacts.keys(), ...missing.keys()])].map(file => file.slice(".dochaus/lean/".length))] }
}

function extractionInputs(directory: string, deps: ChecklistJobDeps) {
  const jobs = checklistProducerInputs(directory)
  const basis = shortHash(JSON.stringify([matterSourceBasis(directory), jobs.map(job => job.input)]), 40)
  return jobs.map(job => ({ ...job, input_sha: shortHash(JSON.stringify(["checklist-extraction-1", deps.model, basis, buildChecklistProducerPrompt(job.input)]), 40) }))
}

async function parallel<T>(jobs: T[], work: (job: T, index: number) => Promise<void>) {
  const next = { index: 0 }
  await Promise.all(Array.from({ length: Math.min(4, jobs.length) }, async () => {
    for (;;) {
      const index = next.index++
      if (index >= jobs.length) return
      await work(jobs[index]!, index)
    }
  }))
}

function message(error: unknown) {
  return error instanceof Error ? error.message : String(error)
}


export type ChecklistRefinementSelection = {
  parent_id: string
  scope: "source_collection" | "cited_sources" | "target_body"
  units: (ChecklistRefinementInput["source_unit"] & { assignment?: ChecklistComparisonAssignment })[]
  instructions: ChecklistSourceReference[]
  targets: ChecklistTargetDocument[]
  retrieval: string
}
const refinementUnitSchema = z.object({
  unit_id: z.string(), input_sha: z.string(), status: z.enum(["succeeded", "failed"]),
  result: checklistRefinementSchema.optional(), children: z.array(itemRecordSchema), issues: z.array(z.string()),
  sessionID: z.string().optional(), attempts: z.number().optional(),
})
const refinementParentSchema = z.object({
  parent_id: z.string(), parent: checklistItemSchema, scope: z.enum(["source_collection", "cited_sources", "target_body"]),
  source_basis: z.string(), expected_units: z.array(z.string()), units: z.array(refinementUnitSchema),
})
export const checklistRefinementsSchema = z.object({ parents: z.array(refinementParentSchema) })
const REFINEMENTS = "checklist/refinements.json"

function originalChecklistItems(directory: string): ChecklistItemRecord[] {
  const saved = readLeanJson(directory, EXTRACTION)
  if (saved === undefined) throw new Error("Checklist extraction is required before refinement.")
  return extractionSchema.parse(saved).sources.flatMap(row => row.status === "succeeded" ? row.items : [])
}

/** Explicit refinement preserves authority: source scopes extract focused duties;
 * target scope prepares bounded current comparisons without a model call. */
export async function runChecklistRefinementFlow(directory: string, deps: ChecklistJobDeps, selection: ChecklistRefinementSelection) {
  const original = originalChecklistItems(directory).find(row => row.id === selection.parent_id)
  if (!original) throw new Error(`Unknown original checklist parent: ${selection.parent_id}`)
  if (new Set(selection.units.map(unit => unit.id)).size !== selection.units.length) throw new Error("Refinement source units must have unique IDs.")
  const basis = checklistRefinementSourceBasis(directory, selection.scope === "target_body" ? original.item : undefined)
  const prior = readLeanJson(directory, REFINEMENTS)
  const store = prior === undefined ? { parents: [] } : checklistRefinementsSchema.parse(prior)
  const cached = store.parents.find(row => row.parent_id === original.id)
  if (selection.scope === "target_body") {
    const expected = checklistTargetAssignments(directory, original.item).assignments
    if (JSON.stringify(selection.units.map(unit => ({ id: unit.id, assignment: unit.assignment }))) !== JSON.stringify(expected)) throw Error("Target comparison selection does not cover the actual current accepted paragraphs and process records.")
    const units = expected.map(unit => ({ unit_id: unit.id, input_sha: shortHash(JSON.stringify([basis, unit]), 40), status: "succeeded" as const,
      children: [{ id: `${original.id}-comparison-${shortHash(unit.id, 12)}`, item: original.item, assignment: unit.assignment }], issues: [] }))
    const parent = { parent_id: original.id, parent: original.item, scope: selection.scope, source_basis: basis, expected_units: expected.map(unit => unit.id), units }
    const index = store.parents.findIndex(row => row.parent_id === original.id)
    if (index < 0) store.parents.push(parent); else store.parents[index] = parent
    // Deterministic scope selection creates no legal finding and makes no model call.
    if (JSON.stringify(cached) !== JSON.stringify(parent)) writeLeanText(directory, REFINEMENTS, `${JSON.stringify(store, null, 2)}\n`)
    return { ...readEffectiveChecklistItems(directory), outputs: [REFINEMENTS] }
  }
  const jobs = selection.units.map(source_unit => {
    const input = { parent_id: original.id, parent: original.item, source_unit, instructions: selection.instructions, targets: selection.targets, retrieval: selection.retrieval }
    return { input, input_sha: shortHash(JSON.stringify(["checklist-refinement-1", basis, deps.model, selection.scope, buildChecklistRefinementPrompt(input)]), 40) }
  })
  const records = jobs.map(job => cached?.units.find(row => row.unit_id === job.input.source_unit.id && row.input_sha === job.input_sha))
  const save = () => {
    const parent = { parent_id: original.id, parent: original.item, scope: selection.scope, source_basis: basis, expected_units: selection.units.map(unit => unit.id), units: records.filter((row): row is z.infer<typeof refinementUnitSchema> => !!row) }
    const index = store.parents.findIndex(row => row.parent_id === original.id)
    if (index < 0) store.parents.push(parent); else store.parents[index] = parent
    writeLeanText(directory, REFINEMENTS, `${JSON.stringify(store, null, 2)}\n`)
  }
  // Persist selected scope even when empty: omission must not fall back to a pass.
  save()
  await parallel(jobs, async (job, index) => {
    if (records[index]?.status === "succeeded") return
    const base = { unit_id: job.input.source_unit.id, input_sha: job.input_sha }
    try {
      const result = await runChecklistRefinementJob(directory, deps, job.input)
      records[index] = result.status === "succeeded"
        ? { ...base, status: "succeeded", result: result.value, children: result.value.items.map((item, ordinal) => ({ id: `${original.id}-refined-${shortHash(job.input.source_unit.id, 12)}-${ordinal + 1}`, item })), issues: [], sessionID: result.sessionID, attempts: result.attempts }
        : { ...base, status: "failed", children: [], issues: result.issues, sessionID: result.sessionID, attempts: result.attempts }
    } catch (error) { records[index] = { ...base, status: "failed", children: [], issues: [message(error)] } }
    save()
  })
  return { ...readEffectiveChecklistItems(directory), outputs: [REFINEMENTS] }
}

export function readEffectiveChecklistItems(directory: string, originals = originalChecklistItems(directory)) {
  const saved = readLeanJson(directory, REFINEMENTS)
  const parents = saved === undefined ? [] : checklistRefinementsSchema.parse(saved).parents
  const basis = checklistRefinementSourceBasis(directory)
  const issues: string[] = []
  const items: ChecklistItemRecord[] = []
  for (const original of originals) {
    const parent = parents.find(row => row.parent_id === original.id)
    if (!parent) { items.push(original); continue }
    if (parent.source_basis !== (parent.scope === "target_body" ? checklistRefinementSourceBasis(directory, original.item) : basis) || JSON.stringify(parent.parent) !== JSON.stringify(original.item)) {
      issues.push(`${original.id}: refinement source or original requirement changed; explicit refinement is required.`)
      continue
    }
    if (parent.scope === "target_body") {
      const expected = checklistTargetAssignments(directory, original.item).assignments
      if (JSON.stringify(parent.expected_units) !== JSON.stringify(expected.map(unit => unit.id)) || parent.units.length !== expected.length) issues.push(`${original.id}: current paragraph/process assignment coverage is incomplete.`)
      for (const assignment of expected) {
        const unit = parent.units.find(row => row.unit_id === assignment.id)
        if (!unit || unit.status !== "succeeded" || unit.children.length !== 1 || JSON.stringify(unit.children[0]!.item) !== JSON.stringify(original.item) || JSON.stringify(unit.children[0]!.assignment) !== JSON.stringify(assignment.assignment)) issues.push(`${original.id}/${assignment.id}: current comparison assignment is missing or changed.`)
        if (unit?.status === "succeeded") items.push(...unit.children)
      }
      continue
    }
    if (!parent.expected_units.length) issues.push(`${original.id}: no assigned source units.`)
    for (const id of parent.expected_units) {
      const unit = parent.units.find(row => row.unit_id === id)
      if (!unit || unit.status !== "succeeded" || !unit.result || unit.result.disposition === "unresolved") issues.push(`${original.id}/${id}: missing, failed or unresolved source disposition.`)
      if (unit?.result?.disposition === "refined" && unit.children.some(child => child.item.target.document !== original.item.target.document || child.item.target.treatment !== original.item.target.treatment)) issues.push(`${original.id}/${id}: children do not preserve parent target document and treatment; parent scope remains unresolved.`)
      if (unit?.status === "succeeded") items.push(...unit.children)
    }
    if (!parent.units.some(unit => unit.children.length)) issues.push(`${original.id}: no child checks; parent scope remains unresolved.`)
  }
  for (const parent of parents) if (!originals.some(row => row.id === parent.parent_id)) issues.push(`${parent.parent_id}: original parent is unavailable.`)
  return { items, parents, complete: !issues.length, issues }
}

/** Structural coverage reconciliation only. Legal judgments remain native child findings. */
export function reconcileChecklistParentCoverage(directory: string, checks: ChecklistCheckResult[], originals = originalChecklistItems(directory)) {
  const effective = readEffectiveChecklistItems(directory, originals)
  return effective.parents.map(parent => {
    const issues = effective.issues.filter(issue => issue.startsWith(`${parent.parent_id}:`) || issue.startsWith(`${parent.parent_id}/`))
    const children = parent.units.flatMap(unit => unit.children.map(row => row.id))
    const findings = children.map(id => checks.find(row => row.item_id === id))
    for (const [index, check] of findings.entries()) if (!check || check.status !== "succeeded" || check.reinspection?.status === "failed" || !check.finding) issues.push(`${children[index]}: no completed finding.`)
    // Applied adoption cannot be certified by an annotation or a stated intention.
    for (const child of parent.units.flatMap(unit => unit.children)) {
      const finding = checks.find(row => row.item_id === child.id)?.finding
      if (checklistFindingRecovery(child.item, finding) === "inspect") issues.push(`${child.id}: action fulfillment or applicability requires inspection.`)
      if (child.item.required_action === "adopt" && child.item.target.treatment === "operative" && finding?.action_status === "fulfilled" && !finding.evidence.some(evidence => evidence.kind === "body" || evidence.kind === "artifact")) issues.push(`${child.id}: adoption lacks current operative or artifact evidence.`)
    }
    const statuses = findings.map(row => row?.finding?.status)
    const unfulfilled = parent.units.flatMap(unit => unit.children).some(child => checklistFindingRecovery(child.item, checks.find(row => row.item_id === child.id)?.finding) === "repair")
    const status = issues.length || statuses.includes("indeterminate") ? "indeterminate" as const
      : unfulfilled ? "unsatisfied" as const
      : statuses.includes("open") || findings.some(row => row?.finding?.open_decisions?.length) ? "open" as const : "satisfied" as const
    return { parent_id: parent.parent_id, scope: parent.scope, status, children, issues }
  })
}

export function checklistRefinementSourceBasis(directory: string, target?: ChecklistItem) {
  return shortHash(JSON.stringify([matterSourceBasis(directory), checklistProducerInputs(directory).map(job => job.input), ...(target ? [checklistTargetAssignments(directory, target)] : [])]), 40)
}

/** Only exact-item product adjudications are relevant. Generated history never
 * enters input/cache identity or replenishes the canonical review allowance. */
export function readChecklistAdjudicationCounterevidence(directory: string, item_id: string, item: ChecklistItem): ChecklistAdjudicationCounterevidence | undefined {
  const documents = readLeanPlan(directory)?.documents.filter(row => row.id === item.target.document || row.file === item.target.document)
  if (!documents || documents.length !== 1) return
  const file = `${documents[0]!.id}/repair-adjudications.json`
  const saved = readLeanJson(directory, file)
  const parsed = z.object({ assessments: z.array(z.object({ item_id: z.string(), value: z.object({ findings: z.array(z.object({
    disposition: z.enum(["deferred", "recommendation", "rejected"]), detail: z.string(), reason: z.string(),
    sources: z.array(z.object({ source_id: z.string(), quote: z.string() })), limitations: z.array(z.string()),
  })).optional() }) })) }).safeParse(saved)
  if (!parsed.success) return
  const matches = parsed.data.assessments.flatMap((row, index) => row.item_id === item_id ? [{ row, index }] : [])
  if (matches.length !== 1 || !matches[0]!.row.value.findings?.length) return
  const { row, index } = matches[0]!
  const locator = `.dochaus/lean/${file} /assessments/${index}`
  const candidate = checklistAdjudicationCounterevidenceSchema.safeParse({ item_id, locator,
    claims: row.value.findings!.map((finding, findingIndex) => ({ disposition: finding.disposition,
      detail: finding.detail, reason: finding.reason, limitations: finding.limitations,
      source_refs: finding.sources.map((source, sourceIndex) => ({ document: source.source_id,
        locator: `${locator}/value/findings/${findingIndex}/sources/${sourceIndex} (generated quotation; verify original source)`, quote: source.quote })),
    })),
  })
  return candidate.success ? candidate.data : undefined
}
