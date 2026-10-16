import { readLeanPlan, type DocID } from "./plan"
import { instructionFound } from "./quotes"

/** A semantic interpretation of an explicitly named or titled Word output in the lawyer's
 * instructions. It selects an existing artifact, never changes its content. */
export type DeliverySelection = ({ document: DocID; role: "body" | "redline" } | { document: null; role: "drafting_notes" }) & { file: string; instruction_quote: string }

export function checkDeliverySelections(directory: string, value: unknown): string[] {
  if (!Array.isArray(value)) return ["Return delivery as an array of user-grounded artifact selections, empty when no output filename was requested."]
  const documents = readLeanPlan(directory)?.documents ?? []
  const issues = value.flatMap((selection): string[] => {
    if (!selection || typeof selection !== "object" || Array.isArray(selection)) return ["Each delivery selection needs document, role, file and instruction_quote."]
    const row = selection as Partial<DeliverySelection>
    if (row.role === "drafting_notes") {
      if (row.document !== null) return ["The run's drafting_notes artifact requires document null."]
    } else {
      if (!documents.some(document => document.id === row.document)) return ["Delivery selection must name an existing final document id."]
      if (row.role !== "body" && row.role !== "redline") return ["Delivery role must be body, redline or drafting_notes; select an existing artifact without changing its content or comparison source."]
    }
    if (typeof row.file !== "string" || !/^[^\\/:*?"<>|`\x00-\x1f]+\.docx$/i.test(row.file) || row.file !== row.file.trim()) return ["Delivery file must be the requested Word filename, without a folder path."]
    if (typeof row.instruction_quote !== "string" || !instructionFound(directory, row.instruction_quote) || !deliveryNameQuoted(row.instruction_quote, row.file)) return ["Delivery instruction_quote must quote the lawyer's actual instructions naming that output file or explicitly titling the Word output with its exact basename. Quote the actual naming instruction; source titles and input filenames do not themselves establish output intent. For an explicitly titled Word output, add only .docx to that title; do not guess, shorten or rename it."]
    return []
  })
  if (issues.length) return issues
  const rows = value as DeliverySelection[]
  if (new Set(rows.map(row => row.document)).size !== rows.length) issues.push("The plan supports one named primary artifact per document and one name for the run's notes. Report conflicting or additional requested outputs as a package finding rather than choosing silently.")
  if (new Set(rows.map(row => row.file.toLocaleLowerCase())).size !== rows.length) issues.push("Distinct named primary artifacts need distinct filenames. Report a conflicting naming instruction as a package finding.")
  return issues
}

/** Naming intent and artifact roles belong to the semantic reviewer. The
 * checker only grounds the literal name; adding Word's extension to an
 * explicitly stated title does not require a model to invent another name. */
function deliveryNameQuoted(quote: string, file: string) {
  if (quote.includes(file)) return true
  const title = file.slice(0, -5).replace(/[.*+?^${}()|[\]\\]/g, "\\$&")
  const named = new RegExp(`\\b(?:titled|entitled|named|called|with the title)\\s+["“'‘]?(${title})(?=$|[,;!?]|\\.(?:\\s|$)|["”'’])`, "gi")
  return [...quote.matchAll(named)].some(match => match[1] === file.slice(0, -5))
}

/** Only the delivery metadata changes. Stable document ids, source foundations,
 * drafting intent and all operative parts remain the accepted plan's values. */
export function selectedDeliveryPlan(directory: string, selections: DeliverySelection[]) {
  const plan = readLeanPlan(directory)
  if (!plan) return
  const selected = new Map(selections.filter(row => row.role !== "drafting_notes").map(row => [row.document, row]))
  const notes = selections.find(row => row.role === "drafting_notes")
  const documents = plan.documents.map(row => {
    const selection = selected.get(row.id)
    return selection && (row.file !== selection.file || (row.deliverable ?? "body") !== selection.role)
      ? { ...row, file: selection.file, deliverable: selection.role } : row
  })
  return documents.some((row, index) => row !== plan.documents[index]) || (notes && notes.file !== plan.notes_file)
    ? { ...plan, documents, ...(notes ? { notes_file: notes.file } : {}) } : undefined
}
