import { normalize, normalizeWithMap } from "../../lib/quote-verify"
import { combinedDelivery, memoAnchor, memoNotesFile, plannedRegistration } from "./deliverables"
import { readLeanPlan, type DocID, type PlannedDocument } from "./plan"
import { instructionFound, lawyerWords } from "./quotes"

/** A semantic interpretation of an explicitly named or titled Word output in the lawyer's
 * instructions. It selects an existing artifact, never changes its content. */
export type DeliverySelection = ({ document: DocID; role: "body" | "redline" } | { document: null; role: "drafting_notes" }) & { file: string; instruction_quote: string }

export function checkDeliverySelections(directory: string, value: unknown): string[] {
  if (!Array.isArray(value)) return ["Return delivery as an array of user-grounded artifact selections, empty when no output filename was requested."]
  const documents = readLeanPlan(directory)?.documents ?? []
  const issues = value.flatMap(selection => deliverySelectionIssues(directory, selection, documents))
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
    return [`"${name}" is the one file that holds the memo and document ${combined.document}. Select this file by one row: document "${combined.document}", role body, filename "${name}". Keep the other rows. The code builds it from the memo and the clean body at delivery; give no drafting_notes row for it.`]
  }))
  if (!messages.size && !distinctFiles(rows)) messages.add(DISTINCT_FILES)
  return [...issues, ...messages]
}

/** The checks of one delivery row on its own: its shape, an existing document
 * and role, a bare Word filename, and a name the lawyer's quoted words give.
 * Rows that pass them may still conflict with each other
 * (checkDeliverySelections). */
export function deliverySelectionIssues(directory: string, selection: unknown, documents: PlannedDocument[] = readLeanPlan(directory)?.documents ?? []): string[] {
  const refused = deliveryRowRefusal(directory, selection, documents)
  return refused ? [refused.issue] : []
}

/** The check of a row's own that refuses it: its shape, a notes row's document,
 * the document, the role, the file name, the quote or the name, or a format. */
export type DeliveryRefusal = "selection" | "notes_document" | "document" | "role" | "file" | "not_quoted" | "format"

/** Which of a row's own checks (deliverySelectionIssues) first refuses it, in
 * their order, with that check's message; undefined when it passes them all. */
export function deliveryRowRefusal(directory: string, selection: unknown, documents: PlannedDocument[] = readLeanPlan(directory)?.documents ?? []): { refusal: DeliveryRefusal; issue: string } | undefined {
  if (!selection || typeof selection !== "object" || Array.isArray(selection)) return { refusal: "selection", issue: "Each delivery selection needs document, role, file and instruction_quote." }
  const row = selection as Partial<DeliverySelection>
  if (row.role === "drafting_notes") {
    if (row.document !== null) return { refusal: "notes_document", issue: "The run's drafting_notes artifact requires document null." }
  } else {
    if (!documents.some(document => document.id === row.document)) return { refusal: "document", issue: "Delivery selection must name an existing final document id." }
    if (row.role !== "body" && row.role !== "redline") return { refusal: "role", issue: "Delivery role must be body, redline or drafting_notes; select an existing artifact without changing its content or comparison source." }
  }
  if (!wordFileName(row.file)) return { refusal: "file", issue: "Delivery file must be the requested Word filename, without a folder path." }
  if (typeof row.instruction_quote !== "string" || !instructionFound(directory, row.instruction_quote)) return { refusal: "not_quoted", issue: NOT_QUOTED }
  // A name the quote spells with its extension, or one the earlier check
  // accepted, is a name, even in format words: package review is never
  // stricter than before.
  if (deliveryNameQuoted(directory, row.instruction_quote, row.file) || quotedOrTitled(row.instruction_quote, row.file)) return
  // A format the lawyer's words ask for, or a name in another format, is no file to select.
  const stem = row.file.slice(0, -5)
  if (formatStem(stem) || OTHER_FORMAT.test(stem))
    return { refusal: "format", issue: `Delivery file "${row.file}" names a format, not a file. Words that ask only for a format give no file name; return no delivery row for them.` }
  return { refusal: "not_quoted", issue: NOT_QUOTED }
}

/** Whether `file` is a bare Word file name: no folder, no surrounding spaces
 * and no character a file name cannot have. */
export function wordFileName(file: unknown): file is string {
  return typeof file === "string" && /^[^\\/:*?"<>|`\x00-\x1f]+\.docx$/i.test(file) && file === file.trim()
}

/** `file`, a Word file name, when the lawyer's own words give it inside the
 * instruction its row quotes, whatever else refuses its row: the file name in
 * their spelling (`file`, a title with ".docx" added) and exactly as they write
 * it (`spelling`, a title without it). It is a name nameSpellings finds there,
 * where a name may also start after any mark outside a word but a hyphen, a
 * full stop, an ampersand or a plus sign (ASKED_OPENS), compared with the
 * whole file name, or else with its last path part, case and spaces aside: a
 * folder the lawyer writes is quoted with the name, one only the model gives
 * never is. It says only that their words give the name, not what they meant
 * by it. A name cut short from a longer one, or one only in the model's words,
 * is none. Code only, never a prompt. */
export function deliveryNameAsked(directory: string, quote: unknown, file: unknown): { file: string; spelling: string } | undefined {
  if (typeof file !== "string" || typeof quote !== "string" || !instructionFound(directory, quote)) return
  const whole = file.trim()
  for (const name of new Set([whole, whole.replace(/^.*[\\/]/, "").trim()])) {
    if (!/\.docx$/i.test(name)) continue
    const key = normalize(name).toLocaleLowerCase()
    for (const raw of lawyerWords(directory))
      for (const { written, title } of nameSpellings(raw, quote, name.slice(0, -5), ".docx", ASKED_OPENS)) {
        const spelled = title ? `${written}.docx` : written
        if (normalize(spelled).toLocaleLowerCase() === key) return { file: spelled, spelling: written }
      }
  }
}

const NOT_QUOTED = "Delivery instruction_quote must quote the lawyer's actual instructions naming that output file or explicitly titling the Word output with its exact basename. Quote the actual naming instruction; source titles and input filenames do not themselves establish output intent. For an explicitly titled Word output, add only .docx to that title; do not guess, shorten or rename it. When the lawyer's words give that output no file name or explicit title, return no delivery row for it."
const DISTINCT_FILES = "Distinct named primary artifacts need distinct filenames. Report a conflicting naming instruction as a package finding."

function distinctFiles(rows: DeliverySelection[]) {
  return new Set(rows.map(row => row.file.toLocaleLowerCase())).size === rows.length
}

/** The words, in lower case, after which the lawyer's instructions give a
 * file's name or a Word output's title, as a regular-expression alternation;
 * the name follows them after whitespace. */
const NAMING_WORDS = String.raw`titled|entitled|named|called|with\s+the\s+title|(?:name|call|title)\s+(?:it|(?:the|this)\s+(?:file|document|draft))|save\s+(?:(?:it|(?:the|this)\s+(?:file|document|draft))\s+)?as|file\s*name:`

/** A format's name, as formatStem reads it. */
const FORMAT = String.raw`(?:(?:ms|microsoft)\s+)?(?:word|docx?|pdf|rtf|odt|wpd|wordperfect|pages|xps|epub|te?xt|plain\s+text|md|markdown|html?|xml|excel|xlsx?|csv|ods|powerpoint|pptx?|odp)`
const FORMAT_STEM = new RegExp(String.raw`^(?:(?:an?\s+)?${FORMAT}(?:\s+(?:docx?|document|file|format|version|copy))?(?:\s*\(\s*\.?${FORMAT}\s*\))?)?$`, "i")

/** Whether a file name's stem is a format's name rather than a name: words
 * that ask for a format never name a file, nor do they with a format's
 * extension or name after them in brackets. An empty stem is Word's own format
 * too: ".docx" alone names no file. Code only, never a prompt. */
export function formatStem(stem: string) {
  return FORMAT_STEM.test(stem.trim())
}

/** The extensions of common document, text, sheet and slide formats other than
 * Word's own: a name's ending only counts as an extension when it is one. Code
 * only, never a prompt. */
export const OTHER_FORMAT = /\.(?:doc|docm|dot|dotx|dotm|rtf|odt|wpd|pages|pdf|xps|epub|txt|md|html?|xml|xlsx?|xlsm|csv|ods|pptx?|odp)$/i

// One grammar of a name's edges in the lawyer's text, for the files job
// (requested-files.ts) and package review alike, read as instructionFound reads
// it (normalize: curly quotation marks straight, whitespace runs one space). A
// name starts at the text's start, after a space, a colon, or an opening
// quotation mark, bracket or Markdown mark; an apostrophe opens a name only
// outside a word. A name given without ".docx" follows a naming word
// (NAMING_WORDS) after a space, a colon or an opening mark, and ends at a line
// break, a clause mark, a colon, a closing quotation mark, bracket or Markdown
// mark, a full stop before a space, a closing mark or the end, or a dash
// between spaces.
const WORD = /[\p{L}\p{N}_]/u
const NAME_OPENS = /["«(\[*`:]/
/** Where else a name the lawyer asked for (deliveryNameAsked) may start: after
 * any mark outside a word but a hyphen, a full stop, an ampersand or a plus
 * sign, which join it to a longer name; an apostrophe still only outside a
 * word. */
const ASKED_OPENS = /[^\p{L}\p{N}_\-.'&+]/u
const NAME_END = /^(?:$|[,;!?:"'»)\]*`]|\.(?:["'»)\]*`]|\s|$)|\s[-–—](?:\s|$))/
const LINE_END = /^[^\S\r\n]*[\r\n]/
const NAMING = new RegExp(String.raw`(?:^|[^\p{L}\p{N}_])(?:${NAMING_WORDS}):?\s*(?:\*\*|["'«(\[*\x60])?$`, "u")

/** Lower case where it keeps the text's length, so offsets still map. */
const fold = (text: string) => text.replace(/[\s\S]/gu, char => {
  const lower = char.toLowerCase()
  return lower.length === char.length ? lower : char
})

/** Each spelling in which `raw`, one text of the lawyer's words, gives a file
 * of `stem` and `extension` where it holds `quote` (compared as
 * instructionFound compares them: whitespace and case aside). The stem starts
 * inside the quoted words at a name's start and is followed there by the
 * extension, which makes it a name even when its words also name a format; or,
 * a Word file's given without ".docx", it follows a naming word inside the
 * quote and ends at a name's end, and is then no format's name (formatStem) and
 * ends in no other format's extension (OTHER_FORMAT). So
 * a name cut short from a longer name is refused, a quote may run on past the
 * name, and a line break inside the name does not matter. The spelling is
 * `raw` from the stem's start, whitespace runs one space, with the extension as
 * the lawyer writes it or ".docx" added. Not caught, since no rule of the text
 * says where such a name starts or ends: a name cut short at its start at a
 * space, which leaves a name the lawyer's words also end with its extension,
 * and a name given without ".docx" cut short at a full stop and a space or at a
 * comma. The grammar is not widened further: a name it does not ground reaches
 * the lawyer as a disclosed item. Package review accepts what this grammar
 * grounds or what its earlier literal or titled check accepted
 * (deliveryNameQuoted, quotedOrTitled). */
export function* spelledNames(raw: string, quote: string, stem: string, extension: string) {
  for (const { written, title } of nameSpellings(raw, quote, stem, extension, NAME_OPENS)) yield title ? `${written}.docx` : written
}

/** Each name spelledNames gives, as `raw` writes it (`written`), and whether it
 * is a Word file's given without ".docx" (`title`), which spelledNames adds. A
 * name may also start after a character `opens` matches. */
function* nameSpellings(raw: string, quote: string, stem: string, extension: string, opens: RegExp) {
  const name = fold(normalize(stem))
  const words = fold(normalize(quote))
  const ending = extension.toLowerCase()
  if (!name || !words) return
  const titled = ending === ".docx" && !formatStem(name) && !OTHER_FORMAT.test(name)
  const { text: normal, map } = normalizeWithMap(raw)
  const text = fold(normal)
  const starts = (at: number) => at === 0 || text[at - 1] === " " || opens.test(text[at - 1]!) ||
    (text[at - 1] === "'" && !WORD.test(text[at - 2] ?? ""))
  for (let start = text.indexOf(words); start >= 0; start = text.indexOf(words, start + 1)) {
    const end = start + words.length
    for (let at = text.indexOf(name, start); at >= 0 && at < end; at = text.indexOf(name, at + 1)) {
      if (!starts(at)) continue
      const close = at + name.length
      const spelled = (to: number) => raw.slice(map[at], map[to - 1]! + 1).replace(/\s+/g, " ")
      if (text.startsWith(ending, close) && !WORD.test(text[close + ending.length] ?? "")) yield { written: spelled(close + ending.length), title: false }
      else if (titled && close <= end && NAMING.test(text.slice(start, at)) && (NAME_END.test(text.slice(close)) || LINE_END.test(raw.slice(map[close - 1]! + 1))))
        yield { written: spelled(close), title: true }
    }
  }
}

/** Naming intent and artifact roles belong to the semantic reviewer. The
 * checker only grounds the literal name, by the files job's grammar
 * (spelledNames) in the lawyer's words: adding Word's extension to an
 * explicitly stated title does not require a model to invent another name.
 * The selection matches the lawyer's spelling as instructionFound compares
 * text (normalize: quotation marks and whitespace aside), in the lawyer's case:
 * unlike the files job, which keeps the lawyer's spelling of a name given in
 * another case, it names the file delivered. A name this refuses is still
 * accepted where the earlier check accepted it (quotedOrTitled), unless it
 * names a format; so package review is never stricter than before, and every
 * name the files job grounds is accepted. */
export function deliveryNameQuoted(directory: string, quote: string, file: string) {
  const name = normalize(file)
  return lawyerWords(directory).some(raw => [...spelledNames(raw, quote, file.slice(0, -5), ".docx")].some(spelled => normalize(spelled) === name))
}

/** The earlier check, on the quote alone: the quote holds the file name, or a
 * naming word in it gives the name as a title ending at the quote's end, a
 * clause mark, a full stop or a closing quotation mark. */
function quotedOrTitled(quote: string, file: string) {
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
