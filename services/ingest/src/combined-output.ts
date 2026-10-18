import { combinedParts, readLeanDocuments, type LeanDocument } from "../../../dochaus/workflows/lean/deliverables"
import { deliverableDocumentName } from "./deliverable"

// The one file a combined document is delivered as with its requested memo
// (combinedDelivery): the file its delivery saved in the matter and registered
// in the run's summary (a "combined" row for that document). The run's
// registration writes it, ingest rebuilds it after a Word accept into its body
// or notes while the system's write is unchanged (server.ts), and a lawyer may
// accept changes into it directly; the web downloads it as the matter file it
// is. Without a registered combined row there is none: no file is guessed from
// the plan's name.

/** combinedParts when the run's summary registers the combined file itself (a
 * "combined" row for the same document under that name), every name plain;
 * undefined otherwise. The one rule for reporting and rebuilding the saved
 * combined file. */
export function registeredCombinedParts(directory: string, rows: readonly LeanDocument[] = readLeanDocuments(directory)?.documents ?? []) {
  const parts = combinedParts(directory, [...rows])
  if (!parts || [parts.file, parts.body, parts.notes].some(name => deliverableDocumentName(name) !== name)) return
  return rows.some(row => row.kind === "combined" && row.document === parts.document && row.deliverable === parts.file) ? parts : undefined
}
