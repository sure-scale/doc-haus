import type { LeanOutline } from "./plan"
import { matterSourceIDs, quoteFound } from "./quotes"

/** Selection is a source-backed planning judgment, never a filename classifier. */
export function playbookSelectionIssues(directory: string, documents: LeanOutline["documents"]) {
  const sources = matterSourceIDs(directory)
  return documents.flatMap(document => document.playbook_review?.sources.flatMap(source =>
    !sources.includes(source.source_id) || !quoteFound(directory, source)
      ? [`Playbook review for "${document.title}" must name an available canonical source and quote its scope or position exactly: ${source.source_id}. Read the source and correct the selection; do not invent guidance.`] : []) ?? [])
}
