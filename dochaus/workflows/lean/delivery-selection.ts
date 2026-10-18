import { memoAnchor, memoNotesFile } from "./deliverables"
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
 * drafting intent and all operative parts remain the accepted plan's values.
 * A selected notes name replaces the notes file and never turns on a memo; a
 * requested memo keeps its name. Without one, a memo file still named as
 * memoNotesFile names it is named again after its anchor document (memoAnchor,
 * chosen on the plan), so it follows a rename of that document's body file only;
 * a redline selection, a selection on another document or any other memo file
 * leaves it. A body selection on a combined document names the combined
 * file and keeps it combined; a redline selection still makes it a redline
 * deliverable. */
export function selectedDeliveryPlan(directory: string, selections: DeliverySelection[]) {
  const plan = readLeanPlan(directory)
  if (!plan) return
  const selected = new Map(selections.filter(row => row.role !== "drafting_notes").map(row => [row.document, row]))
  const notes = selections.find(row => row.role === "drafting_notes")
  const documents = plan.documents.map(row => {
    const selection = selected.get(row.id)
    if (!selection) return row
    const role = row.deliverable === "combined" && selection.role === "body" ? "combined" : selection.role
    return row.file !== selection.file || (row.deliverable ?? "body") !== role ? { ...row, file: selection.file, deliverable: role } : row
  })
  // The memo follows only a rename of its plan-time anchor document's body
  // file, never a redline file or another document.
  const anchor = memoAnchor(plan.documents)
  const renamed = documents.find(row => row.id === anchor?.id)
  const follows = !notes && plan.memo && anchor && renamed && plan.notes_file === memoNotesFile(plan.documents, plan.memo)
  const notesFile = notes ? notes.file : follows && renamed.deliverable !== "redline" ? memoNotesFile([renamed], plan.memo!) : plan.notes_file
  return documents.some((row, index) => row !== plan.documents[index]) || notesFile !== plan.notes_file
    ? { ...plan, documents, ...(notesFile !== undefined ? { notes_file: notesFile } : {}) } : undefined
}
