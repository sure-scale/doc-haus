import { leanClarifications } from "./requests"
import { reviewDispositionText } from "./change-pipeline"
import { boundPlaybookContext } from "../../lib/playbook-source"
import { checkDeliverySelections, deliveryNameAsked, deliveryRowRefusal, selectedDeliveryPlan, wordFileName, type DeliveryRefusal, type DeliverySelection } from "./delivery-selection"
import { appliedChangeEvidence, nativeExportEvidence } from "./review-evidence"
import { comparisonDescription } from "./comparison"
import { combinedDelivery, deliveryArtifactEvidence, plannedRegistration, requestedMemo } from "./deliverables"
import { finalBodyParagraphs } from "./checks"
import { readLeanPlan, type DocID } from "./plan"
import { excerpt, matterSourceIDs, openItem } from "./quotes"
import { researchHandoff } from "./research"
import { requestedFilesItemNamed } from "./requested-files"
import { answeredQuestionsText, readLeanJson, readLeanText, type OpenItem } from "./state"

export type PackageReview = {
  status: "pass" | "changes-required"
  findings: { target?: "body" | "memo" | "research"; document?: DocID; pid?: string; detail: string }[]
  limitations: string[]
  delivery: DeliverySelection[]
  /** The delivery rows the last attempt withheld (withholdUngroundedDelivery);
   * code writes them, never the reply. */
  withheld_delivery?: WithheldDelivery[]
}

/** A delivery row the last attempt withheld: what it selects, as the reply
 * gave it; why (`cause`, and `reason` in plain words); and whether the
 * lawyer's own words give its name inside the instruction it quotes (`asked`,
 * deliveryNameAsked), with the file name in their spelling (`asked_file`, a
 * title with ".docx" added) and exactly as they write it (`asked_spelling`). */
export type WithheldDelivery = { document?: string | null; role?: string; file?: string; cause: WithheldCause; reason: string; asked: boolean; asked_file?: string; asked_spelling?: string }

/** A row's own check that refuses it, or a conflict: with another row, with
 * the notes' file name, or with another planned file's name. */
export type WithheldCause = DeliveryRefusal | "conflict" | "notes_name" | "planned_name"

const WITHHELD_REASON: Record<WithheldCause, string> = {
  selection: "not a complete delivery selection",
  notes_document: "a selection of the notes that names a document",
  document: "names no planned document",
  role: "names no role the run delivers",
  file: "not a plain Word file name",
  not_quoted: "no naming instruction in the lawyer's words",
  format: "names a format, not a file",
  conflict: "conflicts with another delivery selection",
  notes_name: "the notes already have that file name",
  planned_name: "another planned file already has that name",
}

/** The semantic reviewer reads the assembled bodies and accompanying memo,
 * after amendment application. Findings are substantive, not bookkeeping. */
export function packageReviewPrompt(directory: string) {
  const documents = readLeanPlan(directory)?.documents ?? []
  const combined = combinedDelivery(directory)
  return [
    "Review this final drafting package after revisions and native comparison export, before deliverable registration. Determine whether the documents perform the requested functions, agree with the settled instructions and evidence, and agree with their accompanying drafting notes. Inspect the actual final wording, not proposed changes or previous approval. Use search-document, read-document and verify-quote to inspect supporting sources on demand; the inventory is not a fixed evidence packet. read-document returns exact pages; follow next_offset until EOF when inspecting a complete source.",
    boundPlaybookContext(directory),
    `<request>\n${readLeanText(directory, "request.md") ?? ""}\n${readLeanText(directory, "updated-notes.md") ?? ""}\n</request>`,
    ...(leanClarifications(directory).length ? [`<inspection_clarifications>\n${JSON.stringify(leanClarifications(directory))}\n</inspection_clarifications>`, "These later human clarifications were admitted for inspection, not document editing. Check their source/applicability explanation against actual artifacts. Do not assume requested changes occurred or infer approval; retain material corrections as findings."] : []),
    `<drafting_spine>\n${readLeanText(directory, "spine.json") ?? ""}\n</drafting_spine>`,
    `<discrepancies>\n${readLeanText(directory, "discrepancy.json") ?? ""}\n</discrepancies>`,
    `<answers>\n${answeredQuestionsText(directory)}\n</answers>`,
    `<sources>\n${matterSourceIDs(directory).join("\n")}\n</sources>`,
    `<delivery_artifacts>\n${JSON.stringify(deliveryArtifactEvidence(directory), null, 2)}\n</delivery_artifacts>`,
    `<legal_research>\n${researchHandoff(directory)}\n</legal_research>`,
    ...documents.map(row => `<final_document id="${row.id}" title="${row.title.replaceAll('"', "'")}">\nPurpose: ${row.purpose ?? "See request"}\n${comparisonDescription(directory, row.id)}\n${finalBodyParagraphs(directory, row.id).map(paragraph => `[${paragraph.pid}] ${paragraph.text}`).join("\n\n")}\n</final_document>`),
    ...documents.filter(row => row.playbook_review?.sources.length).map(row => `<playbook_review document="${row.id}">\nSelection: ${JSON.stringify(row.playbook_review)}\n${readLeanText(directory, `${row.id}/review-playbook.md`) ?? "Not available"}\n</playbook_review>`),
    "For documents with selected playbook guidance, verify its applicable positions against the actual final body and settled instructions. Earlier coverage is a navigation aid, not approval. Confirm corrections were incorporated or accurately disclosed as unresolved; preserve justified overrides and inapplicable positions. Do not treat a memo recommendation as completed proposed wording when the request requires that wording.",
    ...documents.map(row => `<review_dispositions document="${row.id}">\n${reviewDispositionText(directory, row.id)}\n</review_dispositions>`),
    "Check material review dispositions against the resulting package. Deferred information or research must remain visible with its practical consequence and remaining action; it is not verified evidence. Relevant optional recommendations must reach the requested advice or notes as proposals. Affected document ids identify scope, not permission to omit advice about the originating document. Reject a material omission from requested advice, while allowing justified rejections and optional protections to stay out of operative wording. Legacy decisions and unadjudicated review history are navigation; check their actual basis rather than assuming resolution. Do not reopen a supported rejection without contrary evidence.",
    ...documents.map(row => `<native_export document="${row.id}">\n${JSON.stringify(nativeExportEvidence(directory, row.id), null, 2)}\n</native_export>`),
    ...documents.map(row => `<applied_changes document="${row.id}">\n${JSON.stringify(appliedChangeEvidence(directory, row.id), null, 2)}\n</applied_changes>`),
    `<drafting_notes>\n${readLeanText(directory, "drafting-notes.md") ?? ""}\n</drafting_notes>`,
    "Distinguish source-supported proposed positions, client-approved decisions and unresolved commercial choices. A negotiating recommendation may be delivered as proposed without prior approval; never imply agreement or silently choose a genuinely missing material decision. Check that requested advice explains both the original exposure and proposed resolution, with unresolved choices visible to the user.",
    "Check operative function and package completeness, introduced factual or legal-commercial drift, definitions and cross-document dependencies, and every material memo assertion against the bodies. Check descriptions of initial revisions versus later review amendments against change_origin; initial revisions are not changes made during review. Disclosed placeholders and limitations may be acceptable for a review draft; silent unsupported terms, omitted essential functions or contradictory memo statements require changes. Do not demand information already present in accessible sources or reopen resolved decisions without contrary evidence.",
    // A requested memo is the notes; in a body it duplicates them.
    ...(requestedMemo(directory) ? ["The requested memo is the run-wide drafting notes, delivered separately from the document bodies; where a document and the memo share one file, the code places the memo before that body at download. Memo, cover-note or advice text meant for the memo that appears inside a document body is a body finding against that document. Content a document's own purpose or the request assigns to that document still belongs in its body. Advice belongs in a body only when that document's purpose is advice, and a reference to the memo only when the request requires one."] : []),
    "Check material memo assertions about review changes against the applied_changes records and their actual sources arrays. Application mode describes only mechanics and does not establish evidence support. operation is the requested amendment; result_paragraphs is actual final wording after reconciliation, including an empty result for deletion. An earlier deletion request can be superseded by surviving wording and must not be described as final deletion. Inspect source text before judging whether the passages support the change. Route corrections by the affected artifact: a separately planned advice, note or memo document uses target body with its document id; target memo without a document id is only for the run-wide drafting notes.",
    "The native_comment fields report actual Word comment export for initial revision explanations and applied review amendments, including the exported text and exact final range (all split results), exact tracked-deleted range for an explicit deletion, or an unresolved reason. Check attached explanations against their final paragraphs and the same reason/source records used by the drafting notes. An unresolved attachment must remain disclosed; it does not mean the body amendment was unapplied. Inspect the supplied comparison source and final body for material initial revisions as well as later amendments. Verify that each requested substantive revision has an accurate explanation; initial revision reasons are separate records, not applied amendments. Absence of a record is not evidence that no change occurred. Flag missing or inaccurate explanations rather than demanding fake no-op amendments.",
    "Deliverable registration follows a passing package review. delivery_artifacts projects the actual bodies, comparisons and generated drafting-notes Word file using registration's naming logic. planned_file, planned_role and planned_filename describe model selections, not verified user intent. Independently compare requested functions, names and artifact roles against that map and contents; mentioning an input filename or source title does not request that output name. A generated explanatory note may fulfill a note request when its audience, content and naming fit; distinguish that from a separately requested substantive client advice document. Do not endorse a memo's missing-deliverable claim when the mapped generated artifact fulfills it, or pass a real omission merely because it is disclosed. Registration may suffix a name to preserve existing matter documents; pending registration is the normal phase, not a missing artifact. The final_document blocks contain clean accepted wording. Native tracked insertions/deletions are supplied separately in native_export from the actual comparison file; attached Word comment records are in applied_changes. Assess those artifacts directly. Clean text without strike-through is not evidence that the separate redline is absent. Keep native markup out of operative wording; an export gap is an artifact issue, not a request to insert Markdown strike-through or bold changes into the body. Export status and application counts do not prove substantive correctness.",
    "Return delivery selections for each explicitly named or titled primary Word output in the lawyer's instructions, including on a pass. Independently select the existing document id and body or redline role that performs that function, or document null and drafting_notes for the existing run-wide explanatory notes when their reviewed content and audience fulfill the request; quote the complete actual output instruction naming its Word filename or explicitly stating its title. When only a title is supplied, preserve that exact title and append .docx; this is a format extension, not a new name. These selections will update only delivery naming/role, never wording, comparison sources, approval or new documents. Do not copy input filenames, source titles or the plan's selections as user intent. If no output filename or explicit output title is requested, return an empty array. If an instruction cannot be fulfilled by these existing artifacts, report the substantive omission in findings rather than inventing an artifact or silently dropping it.",
    ...(combined ? [`If the lawyer names the file that holds the memo and document ${combined.document}, select it by one row: document "${combined.document}", role body, that filename. The code builds that file from the memo and the clean body at download, so give no drafting_notes row for it.`] : []),
    "Check that each redline’s stated baseline matches the requested comparison and that the memo describes it accurately. A reusable template alone does not establish an original comparison source. If an expressly requested source-to-final comparison is absent, identify the outstanding deliverable; disclosure alone does not fulfill that requirement.",
    "For a genuinely new authority issue, describe precisely the legal proposition and why current approved research does not answer it; do not supply law from memory. Distinguish an essential unresolved issue from an appropriately disclosed limitation.",
    'Reply with exactly one JSON object: {"status":"pass"|"changes-required","findings":[{"target":"body|memo|research","document":"planned document id, omitted for run-wide drafting notes or package","pid":"affected final paragraph id, omitted if not applicable","detail":"specific problem, evidence and required correction"}],"limitations":["remaining disclosed limitation"],"delivery":[{"document":"existing document id, or null for run-wide notes","role":"body|redline|drafting_notes","file":"exact requested filename, or exact explicit output title plus .docx","instruction_quote":"complete exact user instruction naming or titling that output"}]}. Pass requires no findings. Use changes-required when a material correction remains. Set target body with a document id for operative wording or explanation/comment-record corrections. An explanation-only correction belongs to the affected body document and paragraph: rerun cumulative changes/explanations and native export while preserving correct body wording; never demand a no-op amendment. Use memo without a document id only for run-wide drafting-notes corrections; a separately planned advice or memo document uses body with its document id. Use research only for a genuinely new material authority question beyond current approved findings. Where a package finding affects several bodies, report one finding per affected document. Read-only tools are available; file-writing tools are unavailable.',
  ].join("\n\n")
}

export function checkPackageReview(directory: string, value: unknown): string[] {
  if (!value || typeof value !== "object" || Array.isArray(value)) return ["Return the package review JSON object."]
  const row = value as Record<string, unknown>
  if (!["pass", "changes-required"].includes(String(row.status)) || !Array.isArray(row.findings) || !Array.isArray(row.limitations) || row.limitations.some(item => typeof item !== "string"))
    return ["Use status pass or changes-required, findings array and string limitations array."]
  const documents = readLeanPlan(directory)?.documents ?? []
  const issues = [...checkDeliverySelections(directory, row.delivery), ...row.findings.flatMap((finding: unknown): string[] => {
    if (!finding || typeof finding !== "object" || Array.isArray(finding)) return ["Each finding must state a specific detail."]
    const item = finding as Record<string, unknown>
    if (typeof item.detail !== "string" || !item.detail.trim()) return ["Each finding must state a specific detail."]
    if (item.target !== undefined && !["body", "memo", "research"].includes(String(item.target))) return ["Finding target must be body, memo or research."]
    if (item.target === "body" && item.document === undefined) return ["A body finding needs the affected document id."]
    if (item.document !== undefined && !documents.some(document => document.id === item.document)) return ["Finding names an unknown document; use a final document id or omit it for a package/memo finding."]
    if (item.pid !== undefined && (typeof item.document !== "string" || !finalBodyParagraphs(directory, item.document as DocID).some(paragraph => paragraph.pid === item.pid))) return ["Finding pid must belong to its named final document."]
    return []
  })]
  if (row.status === "pass" && row.findings.length) issues.push("A pass cannot contain unresolved correction findings; use changes-required.")
  if (row.status === "changes-required" && !row.findings.length) issues.push("State the material correction needed in findings.")
  return issues
}

/** On the last attempt, delivery rows the check still refuses are withheld
 * rather than failing the run, as the changes stage withholds unverified
 * amendments: delivery only names existing artifacts, so without a row the
 * planned name stands. A file name's surrounding spaces are taken off first,
 * so they alone never refuse a row. A row that repeats another's document,
 * role and file is kept once, its file name compared case aside. A row refused
 * on its own is dropped, and so is each row that conflicts (deliveryConflicts);
 * every other row is kept. Each dropped row is saved with why
 * (withheld_delivery) and as a limitation; the gate discloses those whose name
 * the lawyer's words give (withheldDeliveryItems). A missing delivery, or
 * one that is no list or row, is read as no rename, with one limitation; one
 * row sent alone is read as a list of it. Status and findings are never
 * changed, so a reply whose other parts fail the check still fails. */
export function withholdUngroundedDelivery(directory: string, value: Record<string, unknown>): Record<string, unknown> {
  const { delivery, limitations } = value
  if (!Array.isArray(limitations)) return value
  if (delivery === undefined || delivery === null || typeof delivery !== "object")
    return { ...value, delivery: [], limitations: [...limitations, delivery === undefined ? "delivery field missing; planned names stand" : "delivery was not a list; planned names stand"] }
  const fields = (row: unknown): Partial<Record<keyof DeliverySelection, unknown>> => row && typeof row === "object" && !Array.isArray(row) ? row : {}
  const rows: unknown[] = (Array.isArray(delivery) ? delivery : [delivery]).map(row => {
    const { file } = fields(row)
    return typeof file === "string" && file !== file.trim() ? { ...fields(row), file: file.trim() } : row
  })
  const documents = readLeanPlan(directory)?.documents ?? []
  const refusals = rows.map(row => deliveryRowRefusal(directory, row, documents)?.refusal)
  const name = (row: unknown) => { const { file } = fields(row); return typeof file === "string" ? file.toLocaleLowerCase() : file }
  const same = (one: unknown, other: unknown) => fields(one).document === fields(other).document && fields(one).role === fields(other).role && name(one) === name(other)
  const passing = rows.filter((row, index): row is DeliverySelection => !refusals[index] && rows.findIndex((other, at) => !refusals[at] && same(other, row)) === index)
  const conflicts = deliveryConflicts(directory, passing)
  const kept = passing.filter(row => !conflicts.has(row))
  // A refused row selecting what a kept or conflicting row selects adds nothing.
  const records = new Map(rows.flatMap((row, index): [string, WithheldDelivery][] => {
    const cause = refusals[index] ?? conflicts.get(row as DeliverySelection)
    if (!cause || (refusals[index] && passing.some(other => same(other, row)))) return []
    const { document, role, file, instruction_quote } = fields(row)
    const asked = deliveryNameAsked(directory, instruction_quote, file)
    const record: WithheldDelivery = {
      ...(typeof document === "string" || document === null ? { document } : {}),
      ...(typeof role === "string" ? { role } : {}),
      ...(typeof file === "string" && file ? { file } : {}),
      cause, reason: WITHHELD_REASON[cause], asked: asked !== undefined, ...(asked ? { asked_file: asked.file, asked_spelling: asked.spelling } : {}),
    }
    return [[JSON.stringify(record), record]]
  }))
  const withheld = [...records.values()]
  if (Array.isArray(delivery) && !withheld.length && kept.length === delivery.length && kept.every((row, index) => row === delivery[index])) return value
  return { ...value, delivery: kept,
    limitations: [...limitations, ...withheld.map(row => `${row.file ?? "A delivery selection without a file name"}: ${row.reason}; the planned name stands`)],
    ...(withheld.length ? { withheld_delivery: withheld } : {}) }
}

/** The rows, each passing its own checks, that conflict
 * (checkDeliverySelections), and with what: while they do, each row that
 * conflicts alone, with the name the notes keep or another planned file's
 * name; then each pair that conflicts, with each other; then, if the rows left
 * still conflict, all of them. */
function deliveryConflicts(directory: string, rows: DeliverySelection[]) {
  const found = new Map<DeliverySelection, "conflict" | "notes_name" | "planned_name">()
  const conflict = (some: DeliverySelection[]) => checkDeliverySelections(directory, some).length > 0
  if (!conflict(rows)) return found
  const plan = readLeanPlan(directory)
  for (const row of rows) if (conflict([row])) {
    const files = plan ? plannedRegistration(directory, selectedDeliveryPlan(directory, [row]) ?? plan) : []
    found.set(row, row.role !== "drafting_notes" && files.some(file => file.kind === "drafting_notes" && file.wanted === row.file.toLocaleLowerCase()) ? "notes_name" : "planned_name")
  }
  const rest = rows.filter(row => !found.has(row))
  rest.forEach((row, index) => {
    for (const other of rest.slice(index + 1)) if (conflict([row, other])) found.set(row, "conflict").set(other, "conflict")
  })
  const left = rest.filter(row => !found.has(row))
  if (conflict(left)) for (const row of left) found.set(row, "conflict")
  return found
}

const SPELLING = "the name was not selected exactly as your instructions spell it"

/** Why a withheld row's name was not applied to the file it selects, in plain
 * words that hold whenever the lawyer's words give the name (asked): a row
 * refused for its name's spelling, path or case is one whose name was not
 * selected as they spell it, unless the name they write holds a folder. */
function withheldReason(row: WithheldDelivery & { asked_file: string }, memo: boolean) {
  switch (row.cause) {
    case "role": return "it was not settled which file takes the name"
    case "file": return /[\\/]/.test(row.asked_file) ? "a file name cannot include a folder"
      : wordFileName(row.asked_file.trim()) ? SPELLING : "the name is not a usable Word file name"
    case "conflict": return "the name conflicts with another file name chosen for this package"
    case "notes_name": return memo ? "the memo already has the name" : "the drafting notes already have the name"
    case "planned_name": return "another file in this package already has the name"
    default: return SPELLING
  }
}

/** "a", "a and b", "a, b and c". */
function listed(items: string[]) {
  return items.length < 2 ? items.join("") : `${items.slice(0, -1).join(", ")} and ${items.at(-1)}`
}

/** The withheld delivery rows (withheld_delivery) the lawyer reads about: only
 * names the lawyer's words give in the instruction a row quotes (asked), quoted
 * as they write them (asked_spelling), one item per file name (asked_file),
 * each disclosed with the review draft and never blocking it. Any other row
 * stays an internal record and a limitation, so no name a model made reaches
 * the lawyer and the planned name stands. Names compare trimmed, case and spaces aside. None is raised for
 * a name a kept row applies, one the planned file of a document a row selects
 * (or of the notes) already has, or one a requested-files item is about
 * (requestedFilesItemNamed). Each says in one plain sentence only what the code
 * checks: that the instructions mention the name, every file it was not
 * applied to, why, and that their planned file names stand; no id or stage. */
export function withheldDeliveryItems(directory: string): OpenItem[] {
  const review = readPackageReview(directory)
  const plan = readLeanPlan(directory)
  if (!review || !plan || !Array.isArray(review.withheld_delivery) || !review.withheld_delivery.length) return []
  const evidence = deliveryArtifactEvidence(directory)
  const plannedName = (kind: string, document: DocID | null) => {
    const row = evidence.find(other => other.kind === kind && other.document === document)
    return row && ("download_filename" in row ? row.download_filename : row.planned_filename)
  }
  const combined = combinedDelivery(directory, plan)
  const memo = requestedMemo(directory, plan) !== undefined
  const key = (name: string) => name.trim().replace(/\s+/g, " ").toLocaleLowerCase()
  const target = (row: WithheldDelivery) => {
    const notes = row.role === "drafting_notes" && row.document === null
    const document = notes || row.role === "drafting_notes" ? undefined : plan.documents.find(document => document.id === row.document)
    const planned = notes ? plannedName(combined ? "combined" : "drafting_notes", combined?.document ?? null)
      : document && plannedName(combined?.document === document.id ? "combined" : document.deliverable === "redline" ? "redline" : "body", document.id)
    return { notes, document, planned }
  }
  const rows = review.withheld_delivery.filter((row): row is WithheldDelivery & { asked_file: string; asked_spelling: string } => !!row && typeof row === "object" && row.asked === true &&
    typeof row.asked_file === "string" && !!row.asked_file.trim() && typeof row.asked_spelling === "string" && !!row.asked_spelling.trim() &&
    typeof row.cause === "string" && Object.hasOwn(WITHHELD_REASON, row.cause))
  const applied = new Set(review.delivery.map(row => key(row.file)))
  const names = new Map<string, (typeof rows)[number][]>()
  for (const row of rows) names.set(key(row.asked_file), [...names.get(key(row.asked_file)) ?? [], row])
  return [...names].flatMap(([name, group]): OpenItem[] => {
    if (applied.has(name) || requestedFilesItemNamed(directory, group[0].asked_file)) return []
    if (group.some(row => key(target(row).planned ?? "") === name)) return []
    // Each file the name was not applied to, once, under the first reason given for it.
    const reasons = new Map<string, string[]>()
    const seen = new Set<string>()
    const documents = new Set<DocID>()
    for (const row of group) {
      const { notes, document } = target(row)
      const subject = notes ? (memo ? "the memo" : "the drafting notes") : document ? `"${document.title}"` : undefined
      if (!subject || seen.has(subject)) continue
      seen.add(subject)
      if (document) documents.add(document.id)
      const reason = withheldReason(row, memo)
      reasons.set(reason, [...reasons.get(reason) ?? [], subject])
    }
    const parts = [...reasons].map(([reason, subjects]) => subjects.length > 1 ? `${listed(subjects)} keep their planned file names because ${reason}`
      : subjects[0] === "the drafting notes" ? `the drafting notes keep their planned file name because ${reason}` : `${subjects[0]} keeps its planned file name because ${reason}`)
    const detail = `Your instructions mention the name "${excerpt(group[0].asked_spelling.trim().replace(/[\x00-\x1f\x7f]/g, " "), 120)}", but it was not applied: ${parts.length ? parts.join("; ") : "the name could not be matched to a document in this package"}.`
    const only = documents.size === 1 && seen.size === 1 ? [...documents][0] : undefined
    return [{ ...openItem("package-review", "omission", detail, only ? { document: only } : {}), status: "disclosed", resolution: WITHHELD_RESOLUTION }]
  })
}

const WITHHELD_RESOLUTION = "file name in the instructions not applied; disclosed with the review draft"

/** Whether a gate item is a withheld delivery row's (withheldDeliveryItems):
 * the only package-review item disclosed rather than open. */
export function isWithheldDeliveryItem(item: OpenItem) {
  return item.stage === "package-review" && item.status === "disclosed"
}

export function readPackageReview(directory: string): PackageReview | undefined {
  const value = readLeanJson(directory, "package-review.json")
  return checkPackageReview(directory, value).length ? undefined : value as PackageReview
}
