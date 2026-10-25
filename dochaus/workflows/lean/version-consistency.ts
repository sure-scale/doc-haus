import { readDocxBodyView } from "../../lib/docx-revision-views"
import { consistencyFindings, type ConsistencyFinding, type Para } from "./consistency"

// The cross-references and defined terms (#198) of a Word version the lawyer's
// decisions on the other side's tracked changes made (#200), as a notice at
// Apply: nothing is stored and nothing waits on it. Both versions are read with
// every pending change rejected, so a held change counts as not agreed: before
// is the text without any of their changes, after adds exactly what was
// accepted. Clause numbers are read as Word shows them. The form is not known,
// so the document is checked as `other`. Only the findings the decisions
// introduced are returned: a finding of the text before them, by its wording,
// is left out.
export async function markupConsistency(name: string, before: Uint8Array, after: Uint8Array): Promise<ConsistencyFinding[]> {
  const read = async (bytes: Uint8Array): Promise<Para[]> =>
    (await readDocxBodyView(bytes, "baseline_body", { labels: true })).map(row => ({ text: row.paragraph_text }))
  const was = await read(before)
  const now = await read(after)
  const title = name.replace(/\.docx$/i, "")
  const check = (final: Para[], baselines: Para[][]) => consistencyFindings({ title, form: "other", final, original: was, baselines, others: [] })
  const known = new Set(check(was, []).map(finding => finding.detail))
  return check(now, [was]).filter(finding => !known.has(finding.detail))
}
