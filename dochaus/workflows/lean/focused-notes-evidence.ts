import { escapeTagOpeners } from "../../lib/untrusted"
import type { ApplyResult, ParagraphRow } from "./checks"
import type { DocID } from "./plan"

export const NOTES_CURRENT_EVIDENCE_VERSION = "notes-current-evidence-1"
const PASSAGE_CHARS = 4_000
const SPLIT_EVIDENCE_CHARS = 3 * PASSAGE_CHARS
export type NotesTarget = { document: DocID; pids: string[] }
export type NotesBody = { document: DocID; initial: ParagraphRow[]; current: ParagraphRow[]; applied?: ApplyResult["applied"] }
export type NotesContext = { name: string; text: string }

/** Initial amendment targets may have split or been deleted at apply. The
 * current result PIDs, never the proposed replacement, locate accepted text. */
export function currentNotesPids(body: NotesBody, pids: string[]) {
  const current = new Set(body.current.map(row => row.pid))
  const missing: string[] = [], deleted: string[] = []
  const targets = pids.flatMap(pid => {
    const applied = body.applied?.find(row => row.pid === pid)
    const result = applied?.result_pids ?? [pid]
    if (applied && !result.length) { deleted.push(pid); return [] }
    const found = result.filter(id => current.has(id))
    if (found.length !== result.length) missing.push(pid)
    return found
  })
  return { pids: [...new Set(targets)], missing, deleted }
}

/** Whole current targets and nearby qualifications, with explicit inspection
 * limits. Never silently clip a target or turn a bounded view into absence. */
export function notesCurrentEvidence(body: NotesBody, pids: string[]): NotesContext[] {
  const mapped = currentNotesPids(body, pids)
  const targets = new Set(mapped.pids)
  const own = body.current.flatMap((row, index) => targets.has(row.pid) ? [index] : [])
  const survivingNearDeletion = mapped.deleted.flatMap(pid => {
    const index = body.initial.findIndex(row => row.pid === pid)
    return currentNotesPids(body, body.initial.slice(Math.max(0, index - 2), index + 3).filter(row => row.pid !== pid).map(row => row.pid)).pids
  })
  const neighbors = new Set([...own.flatMap(index => [index - 2, index - 1, index + 1, index + 2]),
    ...body.current.flatMap((row, index) => survivingNearDeletion.includes(row.pid) ? [index] : []),
  ].filter(index => index >= 0 && index < body.current.length))
  const rows = body.current.filter((row, index) => targets.has(row.pid) || neighbors.has(index))
  const beforePids = new Set([...pids, ...(body.applied ?? []).filter(row =>
    (row.result_pids ?? [row.pid]).some(pid => targets.has(pid))).map(row => row.pid)])
  const initial = body.initial.filter(row => beforePids.has(row.pid))
  const view = (rows: ParagraphRow[]) => rows.map(({ pid, text, label }) => ({ pid, ...(label ? { label } : {}), text }))
  const outcomes = (body.applied ?? []).filter(row => beforePids.has(row.pid))
    .map(({ id, pid, mode, result_pids }) => ({ id, pid, mode, result_pids: result_pids ?? [pid] }))
  const current = JSON.stringify({ document: body.document, targets: mapped.pids, deleted_initial_pids: mapped.deleted,
    baseline_pairs: initial.map(row => ({ initial: row.pid, current: currentNotesPids(body, [row.pid]).pids })),
    applied_outcomes: outcomes, paragraphs: view(rows) })
  const before = JSON.stringify({ document: body.document, paragraphs: view(initial) })
  return [
    { name: "current_inspection_scope", text: JSON.stringify({ version: NOTES_CURRENT_EVIDENCE_VERSION, document: body.document,
      target_pids: mapped.pids, nearby_qualification_pids: rows.filter(row => !targets.has(row.pid)).map(row => row.pid),
      deleted_initial_pids: mapped.deleted, unavailable_initial_pids: mapped.missing,
      scope: mapped.pids.length ? "Only the listed current targets and nearby qualifications are inlined; other dependencies require targeted inspection." : "No current passage has been established for this point. Missing inspection never proves missing body content.",
    }) },
    { name: "generated_initial_evidence", text: before },
    { name: "current_accepted_evidence", text: current },
    { name: "baseline_meaning", text: "Generated initial evidence is the earlier drafting baseline, not an original source. Current accepted evidence is the applied body now; proposed or withheld wording is not the current body. A historical finding may already be addressed elsewhere. Preserve genuine residual uncertainty, but do not restate an old repair as still missing without checking the current targets and their qualifications. A passage ending at a heading does not establish that the following content is absent. Never infer document-wide absence, unconditional effect or completed approval from this local packet. If material dependencies cannot be inspected, disclose that inspection limit rather than asserting a defect or approval." },
  ]
}

/** The locator returns only existing current PIDs. This is transient context
 * selection, not a decision that a finding is satisfied or a stored ledger. */
export function checkNotesTargets(value: unknown, bodies: NotesBody[]): { targets?: NotesTarget[]; issues: string[]; limitation?: string } {
  if (!value || typeof value !== "object") return { issues: ["Return targets and a limitation as JSON."] }
  const row = value as Record<string, unknown>
  if (!Array.isArray(row.targets) || typeof row.limitation !== "string") return { issues: ["Return targets as a list and limitation as a string."] }
  const targets: NotesTarget[] = []
  for (const target of row.targets) {
    if (!target || typeof target !== "object") return { issues: ["Every target needs document and current pids."] }
    const own = target as Record<string, unknown>
    const body = bodies.find(body => body.document === own.document)
    if (!body || !Array.isArray(own.pids) || !own.pids.length || own.pids.some(pid => typeof pid !== "string" || !body.current.some(row => row.pid === pid)))
      return { issues: ["Select only planned documents and existing current paragraph PIDs; include the actual qualifications needed for this one point."] }
    const existing = targets.find(target => target.document === body.document)
    const pids = [...new Set([...(existing?.pids ?? []), ...own.pids as string[]])]
    if (existing) existing.pids = pids
    else targets.push({ document: body.document, pids })
  }
  if (!targets.length && !row.limitation.trim()) return { issues: ["An empty target list must explain the unavailable inspection; it never establishes absence."] }
  return { targets, limitation: row.limitation.trim(), issues: [] }
}

/** Split the question, not the qualifications. Each comparison unit owns
 * one target/range while retaining its full current paragraph, mapped initial
 * counterpart, selected qualifications and application/proposal outcome.
 * Keeping that material context is necessary when a condition spans ranges. */
export function notesEvidenceUnits(text: string): { text: string; label: string }[] {
  const all = (name: string) => [...text.matchAll(new RegExp(`<${name}>\\n([\\s\\S]*?)\\n</${name}>`, "g"))].map(match => match[1]!)
  const accepted = all("current_accepted_evidence")
  const initial = all("generated_initial_evidence")
  if ([...accepted, ...initial].reduce((size, raw) => size + raw.length, 0) <= SPLIT_EVIDENCE_CHARS) return [{ text, label: "" }]
  type Row = { pid: string; label?: string; text: string }
  type Current = { document: string; targets: string[]; deleted_initial_pids?: string[]; baseline_pairs?: { initial: string; current: string[] }[]; paragraphs: Row[] }
  const current = accepted.map(raw => JSON.parse(raw) as Current)
  const before = initial.map(raw => JSON.parse(raw) as { document: string; paragraphs: Row[] })
  const units = current.flatMap(record => [
    ...record.paragraphs.filter(row => record.targets.includes(row.pid)).map(row => ({ record, row, deleted: false })),
    ...before.filter(row => row.document === record.document).flatMap(row => row.paragraphs.filter(row => record.deleted_initial_pids?.includes(row.pid)).map(row => ({ record, row, deleted: true }))),
  ])
  if (!units.length) return [{ text, label: "" }]
  const strip = ["current_accepted_evidence", "generated_initial_evidence", "initial_target"]
  let common = strip.reduce((value, name) => value.replace(new RegExp(`<${name}>[\\s\\S]*?</${name}>\\s*`, "g"), ""), text)
  const change = /<change>\n([\s\S]*?)\n<\/change>/.exec(common)
  if (change) {
    const evidence = JSON.parse(change[1]!)
    if (Array.isArray(evidence.result_paragraphs)) evidence.result_paragraphs = evidence.result_paragraphs.map(({ pid }: { pid: string }) => ({ pid, text: "Current accepted wording is supplied in the paired evidence below." }))
    common = common.replace(change[0], `<change>\n${escapeTagOpeners(JSON.stringify(evidence))}\n</change>`)
  }
  return units.flatMap(({ record, row, deleted }) => {
    const pairs = deleted ? [row.pid] : record.baseline_pairs?.filter(pair => pair.current.includes(row.pid)).map(pair => pair.initial) ?? []
    const counterpart = before.filter(value => value.document === record.document).map(value => ({ ...value,
      paragraphs: pairs.length ? value.paragraphs.filter(row => pairs.includes(row.pid)) : value.paragraphs,
    }))
    const chunks = Array.from({ length: Math.ceil(row.text.length / PASSAGE_CHARS) }, (_, index) => row.text.slice(index * PASSAGE_CHARS, (index + 1) * PASSAGE_CHARS))
    return chunks.map((passage, index) => ({
      label: `${deleted ? "Deleted initial" : "Current target"} passage${row.label ? ` ${row.label}` : ""}${chunks.length > 1 ? ` (${index + 1}/${chunks.length})` : ""}`,
      text: [common,
        `<generated_initial_evidence>\n${escapeTagOpeners(JSON.stringify(counterpart))}\n</generated_initial_evidence>`,
        `<current_accepted_evidence>\n${escapeTagOpeners(JSON.stringify({ ...record, targets: deleted ? [] : [row.pid] }))}\n</current_accepted_evidence>`,
        `<evidence_unit_scope>\n${JSON.stringify({ document: record.document, target: row.pid, effect: deleted ? "applied deletion" : "current accepted", piece: index + 1, pieces: chunks.length })}\nThis comparison unit owns only the target range below. Its complete owning paragraph, mapped earlier draft, selected qualifications and applied/unapplied outcome remain together because a condition may span ranges. Other selected paragraphs are dependency context, not separate explanation owners. Explain this range with those qualifications; do not re-explain unrelated ranges. A recorded deletion is an applied outcome, not unavailable inspection. Preserve correctly labelled history; an unapplied proposal is never the current body. Do not infer a whole-document omission or unconditional effect from this range.\n</evidence_unit_scope>`,
        `<evidence_unit_focus>\n${escapeTagOpeners(JSON.stringify(passage))}\n</evidence_unit_focus>`].join("\n\n"),
    }))
  })
}
