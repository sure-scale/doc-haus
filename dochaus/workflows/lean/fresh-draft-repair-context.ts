import { readLeanJson, readLeanState, readLeanText } from "./state"
import { readParagraphs, readChangeList, record } from "./checks"
import { checkChangePlan, type ChangePlan } from "./change-pipeline"
import { checklistProducerInputs } from "./checklist-source-context"
import { packageCompositionBlock } from "./prompts"
import { matterSourceIDs } from "./quotes"
import { reviewStages, type DocID } from "./plan"
import type { Assessment, FocusedRepairInput } from "./checklist-repair-adjudication"

const MAX_ASSIGNMENT = 6000
/** A new initial draft invalidates amendment targets, not the source-supported
 * intent of earlier work. Reuse the existing bounded adjudication workers;
 * never turn historical pointers into current targets or invent cache validity. */
export function freshDraftRepairInputs(directory: string, document: DocID, assessments: Assessment[], outputContract: string) {
  const state = readLeanState(directory)
  if (state?.stages[`${document}.draft`]?.status !== "done") throw Error("Fresh-draft repair requires an accepted current draft.")
  const raw = readLeanJson(directory, `${document}/change-plan.json`)
  if (!record(raw) || !Array.isArray(raw.groups) || !Array.isArray(raw.questions)) throw Error("Prior plan history is unavailable; cannot silently discard amendment intent.")
  const priorGroups = raw.groups
  const originals = readParagraphs(directory, document)
  if (!originals.length) throw Error("Current draft paragraphs are unavailable.")
  const staleTargets = new Set(priorGroups.flatMap((group, i) => record(group) && Array.isArray(group.pids) ? group.pids.filter(pid => !originals.some(row => row.pid === pid)).map(pid => `groups[${i}]: [${pid}] is not an initial draft paragraph.`) : []))
  const issues = checkChangePlan(directory, document, raw)
  if (!staleTargets.size || !issues.length || issues.some(issue => !staleTargets.has(issue))) throw Error(`Prior plan is invalid beyond changed draft paragraph targets: ${issues.join(" ")}`)
  // No old group, amendment, question, disposition or generated dependency is
  // accepted as the new plan. Each remains a separately accounted assignment.
  const previous: ChangePlan = { shared_decisions: [], groups: [], dismissed: [], findings: [], questions: [] }
  const assignments: { id: string; kind: string; text: string }[] = []
  const history = (field: string, index: number, summary: unknown) => ({ record: `${field}:${index}`, intent: summary, lookup: { path: `.dochaus/lean/${document}/${field === "amendments" ? "changes" : "change-plan"}.json`, pointer: `/${field}/${index}` } })
  const add = (id: string, kind: string, value: unknown) => {
    const text = typeof value === "string" ? value : JSON.stringify(value)
    if (!text.trim()) return
    if (text.length > MAX_ASSIGNMENT) throw Error(`Adjudication assignment ${id} exceeds its bounded context; decompose before admission.`)
    assignments.push({ id, kind, text })
  }
  const savedChecks = readLeanJson(directory, "checklist/checks.json") as { checks?: { item_id: string }[] } | undefined
  assessments.forEach(row => {
    const index = savedChecks?.checks?.findIndex(check => check.item_id === row.item_id) ?? -1
    if (index < 0) throw Error("Original checklist finding record is unavailable.")
    add(row.item_id, "prior checklist interpretation and assessment; not authority or proof about this refreshed draft", {
      item_id: row.item_id, item: row.item,
      prior_assessment: { status: row.finding.status, reason: row.finding.reason, action_status: row.finding.action_status, open_decisions: row.finding.open_decisions },
      full_prior_finding: { path: ".dochaus/lean/checklist/checks.json", pointer: `/checks/${index}/finding`, note: "Exact saved prior finding including its earlier-draft evidence; targeted lookup only, not current document evidence." },
    })
  })
  for (const stage of reviewStages(directory, document)) {
    if (state.stages[`${document}.${stage}`]?.status !== "done") throw Error(`Fresh draft review ${stage} has not completed.`)
    const review = readLeanText(directory, `${document}/${stage}.md`)
    if (!review?.trim()) throw Error(`Fresh draft review ${stage} is unavailable.`)
    reviewSections(review).forEach((text, index) => add(`${stage}:${index}`, "current independent review; source meaning must be adjudicated", text))
  }
  const changes = readChangeList(directory, document)
  if (!changes) throw Error("Prior amendment history is unavailable; cannot silently drop cumulative corrections.")
  const accountedGroups = new Set<number>()
  changes.amendments.forEach((value, index) => {
    const related = priorGroups.flatMap((group, i) => record(group) && Array.isArray(group.pids) && group.pids.includes(value.pid) && !accountedGroups.has(i) ? [i] : [])
    related.forEach(i => accountedGroups.add(i))
    add(`historical-amendment:${index}`, "earlier-draft cumulative correction and related planning intent; old PID/hash are historical, never current targets", {
      amendment: value, related_plan: related.map(i => history("groups", i, priorGroups[i])),
    })
  })
  priorGroups.forEach((group, index) => {
    if (!accountedGroups.has(index)) add(`historical-plan:groups:${index}`, "earlier-draft planning intent not represented by an amendment; not source authority", history("groups", index, group))
  })
  for (const field of ["findings", "questions", "dismissed"] as const) {
    if (!Array.isArray(raw[field])) continue
    raw[field].forEach((value, index) => add(`historical-plan:${field}:${index}`, "earlier-draft disposition; reassess against current source/draft without adopting its generated scope", history(field, index, value)))
  }
  if (Array.isArray(raw.shared_decisions) && raw.shared_decisions.length) add("historical-dependencies", "earlier shared planning dependencies; generated interpretations, not settled instructions", raw.shared_decisions)
  changes.withheld?.forEach((value, index) => add(`historical-withheld:${index}`, "earlier withheld intent; never implemented authority, reassess source support and disclose unresolved reasons", value))
  changes.explanations?.forEach((value, index) => add(`historical-explanation:${index}`, "earlier-draft comment intent; reassess relevance without no-op amendments", value))
  changes.retained_initial?.forEach((value, index) => add(`historical-retained-initial:${index}`, "earlier-draft restoration intent with cited support; reassess source support and current wording, never adopt its old PID as a current target", {
    retained_initial: value, lookup: { path: `.dochaus/lean/${document}/changes.json`, pointer: `/retained_initial/${index}` },
  }))
  const request = checklistProducerInputs(directory).find(row => row.key === "request")!.input
  const instructions = [...request.passages, ...request.instructions]
  const composition = packageCompositionBlock(directory, document)
  const inputs: FocusedRepairInput[] = assignments.map(assignment => {
    const named = originals.flatMap((row, index) => assignment.text.includes(row.pid) ? [index - 1, index, index + 1] : [])
    const terms = new Set(assignment.text.toLowerCase().match(/[\p{L}\p{N}]{4,}/gu) ?? [])
    const ranked = originals.map((row, index) => ({ index, score: [...new Set(row.text.toLowerCase().match(/[\p{L}\p{N}]{4,}/gu) ?? [])].filter(term => terms.has(term)).length })).sort((a, b) => b.score - a.score || a.index - b.index)
    const hints = new Set(named.length ? named : ranked.filter(row => row.score > 0).slice(0, 3).map(row => row.index))
    let hintChars = 0
    const paragraphs = originals.filter((row, index) => {
      if (!hints.has(index) || hintChars + row.text.length > MAX_ASSIGNMENT) return false
      hintChars += row.text.length
      return true
    })
    return { assessment: { item_id: assignment.id }, previous, group_indices: [], finding_indices: [], decision_indices: [], artifacts: [], prompt: [
      "Adjudicate ONE assignment against original sources and the CURRENT newly drafted document. Generated reviews, prior amendments and checklist interpretations are unverified evidence, not authority. Preserve every source-supported correction; if already present in the fresh draft, return an explicit no-change disposition. Reject unsupported scope with a source-grounded reason, or defer unresolved support honestly. Do not draft replacement text.",
      `Assignment identity and provenance: ${JSON.stringify({ id: assignment.id, kind: assignment.kind })}`,
      `<assignment>\n${assignment.text}\n</assignment>`,
      `<original_instructions_and_actual_answers>\n${JSON.stringify(instructions)}\n</original_instructions_and_actual_answers>`,
      ...(composition ? [composition] : []),
      `<current_initial_draft_hints>\n${paragraphs.map(row => `[${row.pid}] ${row.text}`).join("\n\n")}\n</current_initial_draft_hints>`,
      "These lexical hints are nonexhaustive. Historical PIDs, hashes, dispositions and quoted generated wording are not current amendment targets or client waivers. Verify the original source condition, action and requested treatment. A proposal is not adoption; unknown facts and approval remain unknown. Do not waive research from generated limitations or turn observations into disclosure duties.",
      `Current amendment targets are ONLY exact PIDs from read-record .dochaus/lean/${document}/paragraphs.json. Use targeted current_body lookup for document=${document}; saved body-applied/redline artifacts are historical when they belong to an earlier initial draft. Original uploads ${JSON.stringify(matterSourceIDs(directory))}: search-document/read-document view=source with exact identity. Inspect only necessary adjacent conditions and related documents. Missing lexical matches never prove absence.`,
      "Other assignments are adjudicated independently, including every current review section and every prior amendment intent. Do not enumerate or replay unrelated history. Return only necessary current-PID goals, or an explicit reasoned disposition; no preserve-only/no-change groups. No supersession indices are supplied in this fresh-draft assignment, so both superseded arrays must be empty. All output sources and current PIDs remain subject to strict canonical validation.",
      outputContract,
    ].join("\n\n") }
  })
  return { previous, inputs }
}

/** Preserve all review text, including qualifications, without an omnibus job.
 * Markdown table rows remain intact with their column headings in each part. */
export function reviewSections(text: string): string[] {
  const result: string[] = []
  let current = ""
  for (const section of text.split(/(?=^#{1,6}\s)/m).filter(part => part.trim())) {
    const heading = /^#{1,6}\s[^\n]+/.exec(section)?.[0] ?? "Current review context"
    for (const paragraph of section.split(/\n\s*\n/).flatMap(paragraph => reviewParagraphs(paragraph, MAX_ASSIGNMENT - heading.length - 4))) {
      if (current && current.length + paragraph.length + 2 > MAX_ASSIGNMENT) { result.push(current); current = heading }
      current += `${current ? "\n\n" : ""}${paragraph}`
    }
  }
  if (current.trim()) result.push(current)
  return result
}

function reviewParagraphs(paragraph: string, limit: number): string[] {
  if (paragraph.length <= limit) return [paragraph]
  const lines = paragraph.split("\n")
  const separator = /^\s*\|?\s*:?-+:?\s*(?:\|\s*:?-+:?\s*)+\|?\s*$/
  if (lines.length < 3 || !separator.test(lines[1]!) || lines.some(line => !line.includes("|"))) {
    throw Error("A review paragraph exceeds bounded adjudication context; decompose before admission.")
  }
  const header = lines.slice(0, 2).join("\n")
  const parts: string[] = []
  let current = header
  for (const row of lines.slice(2)) {
    if (header.length + row.length + 1 > limit) throw Error("A review table row exceeds bounded adjudication context; decompose before admission.")
    if (current.length + row.length + 1 > limit) { parts.push(current); current = header }
    current += `\n${row}`
  }
  parts.push(current)
  return parts
}
