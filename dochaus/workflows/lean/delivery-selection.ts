import { combinedDelivery, memoAnchor, memoNotesFile, plannedRegistration } from "./deliverables"
import { readLeanPlan, type DocID, type PlannedDocument } from "./plan"
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
  // Names are checked as registration gives them (plannedRegistration) on the
  // plan the reply leaves (selectedDeliveryPlan), held in memory. The file each
  // row selects, and the notes when the reply moves their name, must keep that
  // name, wanted by no other requested file; registration would otherwise
  // suffix or yield it. A repeat the reply leaves alone, and a matter document
  // under the name, stay registration's to suffix. The combined file is the
  // combined document's body row alone, so the notes under its name, from a
  // notes row, are a selection error rather than a conflicting instruction.
  const plan = readLeanPlan(directory)
  const applied = plan && (selectedDeliveryPlan(directory, rows) ?? plan)
  const files = applied ? plannedRegistration(directory, applied) : []
  const combined = files.find(file => file.kind === "combined")
  const notes = rows.find(row => row.role === "drafting_notes")
  const selected = (document: DocID | null) => {
    const kind = document === null ? "drafting_notes" : document === combined?.document ? "combined"
      : applied?.documents.find(row => row.id === document)?.deliverable === "redline" ? "redline" : "body"
    return files.find(file => file.kind === kind && file.document === document)
  }
  const targets = new Set([...rows.map(row => selected(row.document)), ...(applied?.notes_file !== plan?.notes_file ? [selected(null)] : [])])
  const messages = new Set([...targets].flatMap(target => {
    const group = target ? files.filter(file => file.wanted === target.wanted && (file === target || file.requested)) : []
    if (!target || (target.name === target.wanted && group.length === 1)) return []
    if (!combined || !notes || group.length !== 2 || !group.includes(combined) || !group.some(file => file.kind === "drafting_notes")) return [DISTINCT_FILES]
    const name = rows.findLast(row => row.document === combined.document)?.file ?? notes.file
    return [`"${name}" is the one file that holds the memo and document ${combined.document}. Select this file by one row: document "${combined.document}", role body, filename "${name}". Keep the other rows. The code builds it from the memo and the clean body at download; give no drafting_notes row for it.`]
  }))
  if (!messages.size && !distinctFiles(rows)) messages.add(DISTINCT_FILES)
  return [...issues, ...messages]
}

const DISTINCT_FILES = "Distinct named primary artifacts need distinct filenames. Report a conflicting naming instruction as a package finding."

function distinctFiles(rows: DeliverySelection[]) {
  return new Set(rows.map(row => row.file.toLocaleLowerCase())).size === rows.length
}

/** The words, in lower case, after which the lawyer's instructions give a
 * file's name or a Word output's title, as a regular-expression alternation;
 * the name follows them after whitespace. requested-files.ts grounds a name
 * given without ".docx" on the same words. */
export const NAMING_WORDS = String.raw`titled|entitled|named|called|with\s+the\s+title|(?:name|call|title)\s+(?:it|(?:the|this)\s+(?:file|document|draft))|save\s+(?:(?:it|(?:the|this)\s+(?:file|document|draft))\s+)?as|file\s*name:`

/** Whether a file name's stem is a format's name rather than a name: words
 * that ask for a format never name a file. Code only, never a prompt. */
export function formatStem(stem: string) {
  return /^(?:an?\s+)?(?:(?:ms|microsoft)\s+)?(?:word|docx?|pdf|rtf|odt|wpd|wordperfect|pages|xps|epub|te?xt|plain\s+text|md|markdown|html?|xml|excel|xlsx?|csv|ods|powerpoint|pptx?|odp)(?:\s+(?:document|file|format|version|copy))?$/i.test(stem.trim())
}

/** Naming intent and artifact roles belong to the semantic reviewer. The
 * checker only grounds the literal name; adding Word's extension to an
 * explicitly stated title does not require a model to invent another name,
 * and a format's name (formatStem) is no title. */
export function deliveryNameQuoted(quote: string, file: string) {
  if (formatStem(file.slice(0, -5))) return false
  if (quote.includes(file)) return true
  const title = file.slice(0, -5).replace(/[.*+?^${}()|[\]\\]/g, "\\$&")
  const named = new RegExp(`\\b(?:${NAMING_WORDS})\\s+["“'‘]?(${title})(?=$|[,;!?]|\\.(?:\\s|$)|["”'’])`, "gi")
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
 * deliverable. A body row that renames the combined file onto the saved notes
 * name, without a notes row, gives the memo its default name (memoNotesFile)
 * again: that name is the combined file's alone. Applying the reply again
 * changes nothing. */
export function selectedDeliveryPlan(directory: string, selections: DeliverySelection[]) {
  const plan = readLeanPlan(directory)
  if (!plan) return
  const selected = new Map(selections.flatMap(row => row.role === "drafting_notes" ? [] : [[row.document, row] as const]))
  const notes = selections.find(row => row.role === "drafting_notes")
  const documents = plan.documents.map((row): PlannedDocument => {
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
  // A pass's notes row may have saved the name a later body row gives the combined file.
  const combined = combinedDelivery(directory, plan)?.document
  const onto = combined && selected.get(combined)
  const reclaimed = !notes && onto?.role === "body" && onto.file !== plan.documents.find(row => row.id === combined)?.file
    && onto.file.toLocaleLowerCase() === plan.notes_file?.toLocaleLowerCase()
  const notesFile = notes ? notes.file : reclaimed ? memoNotesFile(documents, plan.memo!)
    : follows && renamed.deliverable !== "redline" ? memoNotesFile([renamed], plan.memo!) : plan.notes_file
  return documents.some((row, index) => row !== plan.documents[index]) || notesFile !== plan.notes_file
    ? { ...plan, documents, ...(notesFile !== undefined ? { notes_file: notesFile } : {}) } : undefined
}
