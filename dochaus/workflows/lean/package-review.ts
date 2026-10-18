import { leanClarifications } from "./requests"
import { reviewDispositionText } from "./change-pipeline"
import { boundPlaybookContext } from "../../lib/playbook-source"
import { checkDeliverySelections, type DeliverySelection } from "./delivery-selection"
import { appliedChangeEvidence, nativeExportEvidence } from "./review-evidence"
import { comparisonDescription } from "./comparison"
import { deliveryArtifactEvidence, requestedMemo } from "./deliverables"
import { finalBodyParagraphs } from "./checks"
import { readLeanPlan, type DocID } from "./plan"
import { matterSourceIDs } from "./quotes"
import { researchHandoff } from "./research"
import { answeredQuestionsText, readLeanJson, readLeanText } from "./state"

export type PackageReview = {
  status: "pass" | "changes-required"
  findings: { target?: "body" | "memo" | "research"; document?: DocID; pid?: string; detail: string }[]
  limitations: string[]
  delivery: DeliverySelection[]
}

/** The semantic reviewer reads the assembled bodies and accompanying memo,
 * after amendment application. Findings are substantive, not bookkeeping. */
export function packageReviewPrompt(directory: string) {
  const documents = readLeanPlan(directory)?.documents ?? []
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

export function readPackageReview(directory: string): PackageReview | undefined {
  const value = readLeanJson(directory, "package-review.json")
  return checkPackageReview(directory, value).length ? undefined : value as PackageReview
}
