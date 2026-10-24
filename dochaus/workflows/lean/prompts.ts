import { draftingStyleWarningsBlock, styleScopedRows } from "./style-warnings"
import { boundPlaybookContext } from "../../lib/playbook-source"
import { delegationSettles } from "../../lib/question-delegation"
import { readChangePlan, repairBodyText, reviewDispositionText, writerPlan } from "./change-pipeline"
import { appliedChangeEvidence, nativeExportEvidence } from "./review-evidence"
import { comparisonDescription } from "./comparison"
import { combinedDelivery, deliveryArtifactEvidence, requestedMemo } from "./deliverables"
import { researchHandoff } from "./research"
import { finalBodyParagraphs, readChangeList, readParagraphs, type ApplyResult, shownText } from "./checks"
import { reviewStages, FIRST_DOCUMENT, earlierDocuments, planActingFor, operativeDocument, plannedDocument, readLeanPlan, runPosture, stepDocument, stepStage, type DocID, type StepID } from "./plan"
import { REQUEST_SOURCE, matterSourceIDs } from "./quotes"
import { nameWithProvenance, sourceProvenance } from "../../lib/source-provenance"
import { readRepairRound, changesRepairSections, notesRepair } from "./repairs"
import { MEMO_TARGET, REVIEW_FINDINGS_CONTRACT, findingsText, forwardedFindings, reviewText } from "./review-findings"
import { disclosableItems } from "./stages"
import { QUESTION_KIND_CONTRACT, WORKING_DRAFT, isReviewStage, answeredQuestions, answeredQuestionsText, delegatedQuestions, questionOptions, readLeanJson, readLeanText, readQuestions, type OpenItem, type Question, type StageID } from "./state"
import { plannedContract } from "./template-contract"

// Prompts of the lean model steps. Each step's first prompt carries everything
// the step needs; retries carry only the checker's issues. A document after the
// first consumes the final text and accepted source-linked explanations of
// earlier documents. Writers receive no earlier paragraph pointers; semantic
// review separately identifies an earlier body owner when a repair is needed.

const REVIEW_FOCUS = {
  review: "Function, anatomy and package completeness: derive the document's intended purpose from the request and sources, then test whether its operative wording actually performs it. Look for missing operative provisions, dependencies and requested deliverables, not merely accurate existing sentences. A heading or recital alone does not perform an operative function.",
  "review-independent": "Source and legal-commercial fidelity: verify the actual evidence for facts, agreed terms and authority. Identify unintended changes in rights, obligations, remedies and risk. Do not assume a drafting spine or another stage has verified them.",
  "review-playbook": "Playbook adherence: read the complete selected guidance, not only the planning quotes or spine. Check each position against this document, retaining its scope, conditions, exceptions, fallbacks and cross-clause dependencies. Record coverage by position: conforms, explicit instruction override, correction required, unresolved applicability, or not applicable. Give the actual source basis and draft location for each material conclusion. Locate required corrections at an existing paragraph pointer. A deferred recommendation alone does not fulfill a request for proposed wording. Preserve supplied settled decisions; preferred terms are proposals, not agreement. Do not invent legal risk to justify a policy preference. Check unresolved material conflicts against instructions and answers before raising a focused human question.",
  "review-wording": "Problematic wording: identify ambiguity, inconsistency or imprecision that changes how a provision operates. Explain the competing readings and their material effect. Separately assess supplied date/number preference warnings as presentation feedback, preserving exact quotations and expressly retained wording.",
  "review-client-disadvantage": "Client disadvantage: assess concrete adverse effects on the stated client's position within the requested purpose. Distinguish a defect from an agreed allocation of risk and an optional negotiating improvement. When no client is stated, assess the parties even-handedly without selecting one.",
}

const SOURCE_RULES = [
  "- Read the sources you rely on with search-document or read-document. State only what the request or the sources support: do not invent facts, figures or dates about the parties or the matter, commercial or strategic decisions, or authorities.",
  "- Quote a source only word for word, written as \"exact words\" [source: document name or approved source URL].",
  "- A source recommendation is a proposed negotiating position, not an agreed term. When the request calls for proposals, draft the source-supported recommendation, mark it in its paragraph with a visible [open item: confirm proposed ...] marker and explain its consequence; absence of client approval alone does not require withholding it. Preserve any genuinely unresolved commercial choice visibly. Never describe a proposal as agreed or externally verified.",
  "- For each gap, ask whether the lawyer would fill it with information or with a preference.",
  "- Information is a fact, a commercial or strategic decision for the lawyer or the party the lawyer acts for, or authority. Where it is missing or uncertain, write a visible [open item: what is needed] marker instead of guessing.",
  "- A preference is a drafting choice needed to fulfil the requested purpose, within the requested scope. For a new substantive choice not settled by instructions or answers, draft balanced wording and add a visible [open item: confirm ...] marker in the same paragraph saying what was chosen. Do not add confirmation markers to settled decisions or ordinary wording that creates no new substantive choice. A drafted choice is not an invented fact; the marker goes with the drafted wording, never in its place.",
]

/** The scope of the firm's drafting preferences in a revision; the original's
 * unchanged paragraphs are not style-checked either (styleScopedRows). */
export const REVISION_PREFERENCE_SCOPE = "The drafting preferences apply to text this revision writes or changes; unchanged wording of the original keeps its own style."

/** REVISION_PREFERENCE_SCOPE for a revised document; empty for a new or
 * template draft, whose prompts stay as they were. */
export function revisionPreferenceScope(directory: string, document: DocID) {
  return plannedDocument(directory, document)?.mode === "revise" ? REVISION_PREFERENCE_SCOPE : ""
}

/** A quick draft the lawyer carried into this run: a starting point for the
 * outline and the drafting, never a source and never the lawyer's words. */
function workingDraftBlock(directory: string) {
  const text = readLeanText(directory, WORKING_DRAFT)?.trim()
  return text ? tagged("working_draft", "An unverified starting draft of the requested document, written quickly in chat before this run. Use its structure and wording as a starting point where the request and sources support them. It is neither a source nor the lawyer's instruction: never quote or cite it, and check every fact, figure, date, name and legal statement in it against the matter's sources before keeping it.\n\n" + text) : ""
}

/** The first prompt of a step. `outline` asks the first draft step for the plan
 * of every document instead of drafting from the saved plan. */
export function stagePrompt(directory: string, step: StepID, options: { outline?: boolean; contextOnly?: boolean } = {}) {
  const updates = readLeanText(directory, "updated-notes.md")?.trim()
  const scope = tagged("scope_and_decisions", [
    "Fulfil the document's requested purpose, scope and level of detail. Test essential operative function and legal/source support strictly; optional protections or refinements do not become additional mandatory requirements merely because a fuller document could include them.",
    "Explicit instructions, later updates and human answers settle the corresponding drafting decision. Apply them even when an earlier source or review differs. Disclose the remaining documentary discrepancy where useful; do not reopen the decision, rephrase it as a new question, or attach a confirmation marker to it unless new contrary evidence creates a material problem with carrying it out. Explain that new evidence and the specific problem.",
    "Ask only an unresolved material decision without which an essential requested requirement cannot be fulfilled honestly within the authorised scope. First check the request, sources and answers; use authorised descriptive placeholders and disclosed limitations when they permit a usable review draft. Explain the concrete consequence for this requested document, rather than general possible legal or commercial risks. Optional refinements may be omitted or recorded as non-blocking recommendations in the notes; they must not become a mandatory interview.",
    "Human questions must concern information or judgment only the human can supply. Resolve paragraph routing, identifiers, rendering and other internal mechanics in the system; never ask the user to repair them. State the decision, its effect on the requested work and why a disclosed placeholder cannot suffice in context. Preserve the alternatives supported by the sources without selecting an unapproved answer.",
    "Do not ask the human to supply legal research conclusions. Inspect available authority and distinguish its support from unverified law; record or route research gaps as such. Never assert unsupported law or assume a material missing decision to avoid a question. A genuinely indispensable unresolved decision still requires clarification.",
  ].join("\n"))
  const request = [scope, boundPlaybookContext(directory), tagged("legal_research", "Only the approved retrieved propositions below may support external legal assertions. Preserve jurisdiction, currency and source limitations. Missing research is an open item, never an invitation to use model memory. Cite supporting exact words as \"quote\" [source: source URL].\n" + researchHandoff(directory)), tagged("request", readLeanText(directory, "request.md") ?? ""), updates ? tagged("user_updates",
    "Ongoing instructions from the lawyer, in arrival order. Consider these alongside the original request throughout this run. An explicit later change supersedes the earlier instruction on that point; ask about unresolved contradictions. Do not treat a question or a tentative suggestion as an agreed instruction. Do not ask again for information these updates already supply. Quote the lawyer’s words under source name \"request\".\n\n" + updates) : ""].filter(Boolean).join("\n\n")
  const answers = answeredQuestionsText(directory)
  const working = workingDraftBlock(directory)
  const spine = readLeanText(directory, "spine.json")
  const spineBlock = spine ? tagged("drafting_spine", "Source-linked drafting decisions and unresolved points. Retrieve the underlying evidence as needed; this handoff does not replace it.\n" + spine) : ""
  const answerBlock = answers ? tagged("answers", `The lawyer answered these decisions; their answers remain controlling and must not be re-asked under narrower or different wording:\n${answers}`) : ""
  const provenance = sourceProvenance(directory)
  const sources = tagged("sources", [
    "Matter documents (use these exact names with search-document and read-document, and as source names):",
    ...matterSourceIDs(directory).map(name => `- ${nameWithProvenance(name, provenance)}`),
    `To quote the request or the lawyer's answers, use the source name "${REQUEST_SOURCE}".`,
  ].join("\n"))
  if (step === "finalize") return draftingNotesSections(directory, request).filter(Boolean).join("\n\n")
  const document = stepDocument(step) ?? FIRST_DOCUMENT
  const stage = stepStage(step)
  const title = plannedDocument(directory, document)?.title
  const finished = completedDocumentsBlock(directory, document)
  const paragraphs = tagged("draft", readParagraphs(directory, document).map(row => `[${row.pid}] ${shownText(row)}`).join("\n\n"))
  const adjudicated = stage === "changes" ? readChangePlan(directory, document) : undefined
  const named = title ? ` of "${title}"` : ""
  const template = stage === "draft" || stage === "changes" ? plannedContract(directory, document)?.source : undefined
  const acting = actingForBlock(directory)
  const actingFor = planActingFor(directory)
  const planned = plannedDocumentsBlock(directory, document)
  const composition = packageCompositionBlock(directory, document)
  // What the scope guards say other output owns: another planned document, the memo.
  const assigned = [planned ? "another planned document" : "", composition ? "the accompanying memo" : ""].filter(Boolean).join(" or ")
  const outline = tagged("document_outline", JSON.stringify(operativeDocument(plannedDocument(directory, document))))
  const preferenceScope = revisionPreferenceScope(directory, document)
  if (options.contextOnly) return [request, spineBlock, acting, answerBlock, sources, planned, composition, outline, preferenceScope, finished].filter(Boolean).join("\n\n")
  const sections: Partial<Record<StageID, string[]>> = {
    draft: options.outline ? [
      "This is a lean drafting run. Plan the substantive document bodies to draft, then draft each in parts. The workflow generates delivery artifacts from those bodies.",
      request, working, spineBlock, answerBlock, sources,
      tagged("output_contract", [
        ...SOURCE_RULES,
        "- Delivery capabilities: for every planned body the workflow renders the final Word document, with the firm's draft line and legend on every page from its settings, and generates its tracked redline; one final drafting-notes memo, listing the open points, accompanies the package. A request for a redline/comparison or ordinary drafting notes is fulfilled by those generated artifacts, not by another planned body. Set Compare against on the revised body to select a supplied original for comparison. Plan a separate substantive memo only when the request requires its own distinct audience, purpose or analysis beyond the accompanying drafting notes; a different delivery format does not create another substantive document. Do not draft simulated tracked text as a separate comparison document.",
        "- Write one line \"Memo: <its name in the request's own words>\" before the first \"# \" title only when the request asks for a memo or note to accompany the documents; otherwise leave it out. Do not also plan the memo it names as a \"# \" document.",
        "- These rules apply to every part you draft later.",
        "- Begin the outline with one line \"Acting for: <party>\" when the request says which party the lawyer acts for, naming that party; " +
        "otherwise write \"Acting for: not stated\". Do not infer the party from the sources.",
        "- Reply now with the outline only, in Markdown. For each substantive body the request asks for or clearly implies, in the order to draft them " +
        "(a document that refers to or depends on another comes after it): one \"# \" line with that document's title, then one line saying what " +
        "the document is for, then one \"## \" line per top-level part of that document in order, each followed by one line saying what the part covers. " +
        "Do not draft the text yet.",
        "- Under each title write its form: \"Form: agreement\" for an instrument its parties execute to create, vary or end rights or obligations, " +
        "\"Form: advice\" for legal analysis or advice set out as a memorandum or note rather than correspondence, \"Form: advice letter\" for a letter " +
        "giving the client legal advice, \"Form: letter\" for other correspondence addressed to a recipient and signed by the sender, otherwise \"Form: other\".",
        "- Only when the request asks for a page header legend on a document, write under its title \"Legend: <the legend in the request's own words>\"; " +
        "only when it asks for a document without page numbers, \"Page numbers: none\".",
        "- Under each title write \"Executed by: <party>; <party>\" naming each party who executes that document, as the request or sources name them, or \"Executed by: none\" for a document that takes no legal act. Leave the line out when neither says who executes it. " +
        "Plan no document the instructions leave out of this package. For each legal act the instructions direct, unless they leave out the document of the party who can take it, " +
        "plan the act in a document that party executes, adding a companion document where none is planned.",
        "- An answer that leaves a point undecided keeps the instructions as they stand: plan every act they direct. It does not license narrowing the package.",
        "- Under each title write \"Mode: revise\" when revising a supplied original, otherwise \"Mode: new\". For a revision, name the supplied original with Compare against and retain it as the drafting foundation. Also write \"Deliverable: redline\" when the requested named artifact is an annotated markup/comparison, \"Deliverable: combined\" when the request asks for this document and its memo in one named file, otherwise \"Deliverable: body\"; combined needs the \"Memo:\" line. These select the requested artifact role; filenames do not determine it.",
        '- Inspect the supplied documents and bound library guidance before outlining. Decide which are playbooks or drafting guidance from their contents, scope and the request, not filenames alone. Under each title add Playbook review: {"sources":[{"source_id":"exact canonical guidance source","quote":"exact source words establishing its scope or a relevant position"}],"reason":"why this guidance applies to this body"}. Select all applicable guidance, and use sources: [] with a reason when none applies. An uploaded playbook stays a matter source; do not publish it to the shared library. This schedules a dedicated check after drafting only for selected bodies. Read the whole guidance later; these quotes select sources, not an exhaustive position list.',
        "- Titles become document text. Each part's ## label identifies the drafting job and is not printed; the part's text carries the document's own headings in its style. To print a heading in a label's place, add Display heading: <publication heading> immediately under that label. For a letter whose text uses headings, write \"Headings: shown\" under its title to print its labels as its headings. Choose presentation suited to the requested document. For a revision preserve the original heading structure unless the instructions require a structural change; print a heading they change as its part's Display heading. Put internal source mappings and planning explanations in scope notes, not in visible headings or operative wording. Keep client advice in its own appropriate explanatory artifact rather than appending it to an operative agreement.",
        "- Where the request gives the file name a document is to be delivered under, add the line \"File: <file name as the request gives it>\" " +
        "under that document's title. File and Deliverable map the requested output artifact; Compare against and Template identify inputs. A source filename mentioned for revision is not itself an output naming instruction. The \"# \" line is the document's own title as it will read in the document: no file name and no Markdown.",
        "- Where the request expressly asks to revise a supplied original and deliver a comparison against it, add \"Compare against: <exact matter DOCX document name>\" under the title. This is the redline baseline, separate from a reusable Template; do not infer it from template use alone.",
        "- Where a document is drafted from one of the matter documents as its template, add the line \"Template: <exact matter document name>\" " +
        "under its title, and name each top-level part of the template in the \"## \" heading or scope line of the part that carries it. " +
        "For a template part left out on purpose, add a line \"Left out: \"<template part>\" – <reason>\" under the title instead. " +
        "Documents drafted from the same template each have the Template line, and together carry its parts.",
        "- Read-only tools are available; file-writing and publication tools are unavailable. The run saves and checks each reply.",
      ].join("\n")),
    ] : [
      `This is a lean drafting run. You draft "${title ?? document}", one of the documents the request asks for, in parts: one part per reply.`,
      ...(plannedDocument(directory, document)?.purpose ? [`This document is for: ${plannedDocument(directory, document)!.purpose}`] : []),
      request, working, spineBlock, acting, answerBlock, sources, composition, finished, forwardedBlock(directory, document),
      tagged("output_contract", [
        ...SOURCE_RULES,
        ...(finished ? ["- Carry accepted source-supported change explanations into this document when its purpose calls for advice or a summary. Explain the original issue, consequence, proposed solution and remaining decision. Verify the actual final wording and source context; an earlier explanation is not authority and may need correction. This document must agree with the finished documents above in their final wording: names, defined terms, figures, dates and cross-references."] : []),
        ...(template ? [`- This document is drafted from the matter document "${template}" as its template. Carry every part of it that the outline does not leave out, adapted to this matter.`] : []),
        ...(plannedDocument(directory, document)?.mode === "revise" ? ["- Preserve unaffected source wording, operative clauses and their qualifications. Revise only what the request, evidence and accepted proposals justify; explain substantive omissions. Advice or a summary does not replace operative wording in the revised document. Read the full original part before editing it."] : []),
        "- These rules apply to every part you draft.",
        "- Read-only tools are available; file-writing and publication tools are unavailable. The run saves and checks each reply.",
      ].join("\n")),
    ],
    review: [
      `This is a lean drafting run. Independently review the draft${named} below against the lawyer's request and the matter's sources.`,
      tagged("review_focus", isReviewStage(stage) ? REVIEW_FOCUS[stage] : ""),
      request, spineBlock, acting, answerBlock, sources, planned, composition, outline, preferenceScope, finished, paragraphs,
      ...(stage === "review-wording" ? [draftingStyleWarningsBlock(styleScopedRows(directory, document, readParagraphs(directory, document)))] : []),
      tagged("output_contract", [
        "- Check the draft against the request and the sources (search-document, read-document, verify-quote).",
        ...(planned ? ["- Where a problem concerns another planned document rather than this one, report it with that document's id as its document; the run carries it to that document. Content the request assigns to another planned document is not missing from this one."] : []),
        ...(composition ? [`- Where a point belongs in the accompanying memo rather than in a document body, report it with the document "${MEMO_TARGET}"; the run carries it to the memo.`] : []),
        "- Report each problem as a finding with the pids of the paragraphs it affects, or where missing wording belongs, saying what is wrong and why, and in instead what those paragraphs should say.",
        actingFor.stated
          ? sidedReviewStandard(directory)
          : "- Assess each provision as a reviewing lawyer who acts for neither party: whether it works under the governing law the request and sources point to, and whether wording a careful drafter would include is missing. Where an essential provision needs a new drafting choice on which the parties' interests diverge, propose balanced wording within the requested scope and mark that new choice for confirmation; retain decisions already supplied without another confirmation.",
        "- Ask for governing law only when an unresolved jurisdiction is necessary for the requested document function or legal analysis. A factual record or other work whose stated purpose does not require that analysis needs no governing-law question; never assume a jurisdiction for legal conclusions.",
        "- Review every provision, marked or not; wording without an open item is not presumed to work. Report defects affecting rights, obligations, remedies or risk within the requested purpose, including a missing protection essential to that purpose; distinguish optional protections as non-blocking recommendations. Leave out pure style points that change none of these, except the supplied date/number preference warnings assigned to the wording reviewer.",
        "- Preserve agreed terms and answered decisions. Report a specific supported legal defect or material conflict with the requested purpose as a finding; an optional improvement or generic risk is not grounds to reopen the decision as a question.",
        "- Where sources disagree, reconcile the point with explicit instructions, later updates and answers first. Say which evidence or instruction controls. Ask only if a material conflict remains unresolved and prevents fulfilling an essential requested requirement; otherwise disclose the source discrepancy without another question.",
        ...(finished ? ["- The finished documents are not under review. Where the draft disagrees with one of them, report a finding on the draft's paragraph when the draft should change; otherwise report the inconsistency with that finished document's id as its document, for the lawyer to resolve, raising it as a decision only if it leaves an indispensable material decision unresolved under scope_and_decisions."] : []),
        "- Support each finding from the sources in its sources list, copying their words exactly.",
        "- Where a finding rests on your view of the law rather than a passage in the sources, say so; it will be flagged for the lawyer, not applied silently.",
        "- Report as a decision for the lawyer only genuinely unresolved material information or decisions that only the lawyer can supply, subject to scope_and_decisions. Do not turn routine placeholders, optional refinements, already answered choices or legal research gaps into decisions for the lawyer. Report each indispensable decision once, even where several paragraphs depend on it.",
        `- ${REVIEW_FINDINGS_CONTRACT}`,
        ...(stage === "review-playbook" ? ['- Also return "coverage": your coverage of the guidance positions, in Markdown, beside "findings".'] : []),
        "- Read-only tools are available; file-writing and publication tools are unavailable. Your final reply is saved as the review.",
      ].join("\n")),
    ],
    changes: stage === "changes" ? [
      request,
      `This is a lean drafting run. Turn the review findings into paragraph amendments for the draft${named} below.`,
      tagged("change_plan", adjudicated ? JSON.stringify(writerPlan(adjudicated), null, 2) : "No prior adjudication."),
      "The change plan and shared_decisions are generated source-adjudication records, not independent client instructions or waivers. Original sources and actual human answers determine scope. Preserve supported goals and dependencies, but do not let a generated research limit or prior decision silently override them; retain a genuine unresolved issue explicitly.",
      tagged("prepared_changes", readLeanText(directory, `${document}/change-preparation.json`) ?? "No prepared amendments."),
      ...(adjudicated ? ["Act as the final global reconciler. Prepared amendments are proposals, not authority. Preserve every supported correction, reconcile dependent paragraphs and common definitions, and check actual source support. Emit one explicit replacement, deletion or split per initial paragraph incorporating every accepted change; do not select the last proposal or silently discard a supported correction. Retain still-supported prior amendments during repair. If new evidence reveals an indispensable unanswered decision, ask it instead of assuming an answer."] : []),
      paragraphs, ...(readRepairRound(directory)?.items.some(item => (item.stage === "package-review" || item.stage === "change-review" || item.stage === "checklist-check") && item.ref.document === document) ? [tagged("current_final_body", "The semantic review refers to this resulting body (projected before apply or actual final after apply). Translate its finding to the corresponding initial draft paragraph above; output amendments using only initial draft pids. Your list replaces the whole change list: retain prior supported amendments and incorporate the correction so applying the cumulative list to the initial draft produces the intended final body.\n\n" + repairBodyText(directory, document, Boolean(readRepairRound(directory)?.items.some(item => item.stage === "change-review" && item.ref.document === document)), Boolean(readRepairRound(directory)?.items.some(item => item.stage === "checklist-check" && item.ref.document === document))))] : []), ...(adjudicated ? [] : [tagged("findings", reviewStages(directory, document).map(stage => {
        const finding = reviewText(directory, document, stage)
        return finding ? `${stage}:\n${finding}` : ""
      }).filter(Boolean).join("\n\n"))]), ...changesRepairSections(directory, document),
      spineBlock, acting, answerBlock, sources, planned, composition, outline, preferenceScope, finished,
      tagged("output_contract", [
        "Reply with exactly one JSON object and nothing else:",
        '{"amendments":[{"pid":"<pid from the draft>","updated":"<full new text of that paragraph>","reason":"<why it changes>","sources":[{"source_id":"<exact document name, or request>","quote":"<words copied exactly from that source>"}]}],"questions":[{"id":"<id of an open question you are repeating; omit for a new one>","pid":"<pid of the paragraph whose text depends on the answer>","question":"<what only the lawyer can decide>","kind":"<fact|instruction|judgment>","context":"<why it matters>"}]}',
        `- ${QUESTION_KIND_CONTRACT}`,
        '- Also return "explanations":[{"pid":"initial draft paragraph where the change is explained","reason":"actual source-to-initial-draft substantive change, rationale and consequence","sources":[{"source_id":"canonical source or request","quote":"exact supporting words"}]}]. Review the supplied original against the initial draft and explain every material initial revision, including omissions. These are separate comment records, not body text or amendments; never submit an identical replacement merely to create a comment. Anchor a removed provision’s explanation to its related surviving paragraph and explicitly identify the omission. Retain still-valid explanations during repair and correct them against the projected final text. Where no initial source comparison exists, this array may be empty.',
        "- One amendment replaces the whole text of one existing paragraph and stays one paragraph. To add wording, extend the most closely related paragraph. To remove a whole paragraph, use operation delete with an empty updated string and supporting source or settled-instruction passages. To split one initial paragraph into several, use operation split with updated as an array of at least two complete nonempty single-paragraph strings. Keep the initial pid as the operation target; the runner supplies distinct final pids. Empty replacement strings never imply deletion. Unsupported or uncertain deletion retains the wording with a new [open item: confirm ...] marker instead. Do not collapse a split or discard any accepted effect.",
        ...(finished ? ["- Amend only paragraphs of the draft below. The finished documents are shown for agreement and cannot be amended here."] : []),
        ...(assigned ? [`- Amend only this document. Content the request assigns to ${assigned} belongs ${composition ? "there" : "in that document"}, and findings about it reach it separately; do not draft it here.`] : []),
        "- Quote a passage that supports an amendment whenever one exists, copied exactly from a matter document, or from the request or the lawyer's answers under the source name \"request\". The passage must itself support the change.",
        "- Where no passage supports the change, or the change goes beyond what its passage says, it rests on drafting judgment or a view of the law: its updated paragraph must carry a new visible [open item: confirm ...] marker saying what the lawyer is to confirm. Sources may then be empty.",
        ...(template ? [`- Where an amendment takes out or shortens text the draft carries from its template "${template}", its reason says why.`] : []),
        adjudicated ? "- Reconcile prepared group candidates into one cumulative change list. Check their combined effects against adjudicated goals, source context and settled decisions. Give one explicit operation per initial pid incorporating all accepted corrections and preserving unaffected text. Correct unsupported candidates rather than copying them; do not silently drop an accepted goal or still-supported previous amendment." : "- Reconcile all independent reviews into one change list. Consolidate duplicate findings and resolve conflicting suggestions against the request, sources and answers. An exact quotation confirms its words, not the reviewer's conclusion: verify that the passage supports the proposed legal or commercial effect before amending. Reviewer agreement alone is not evidence.",
        adjudicated ? "- Adjudication has resolved the review findings and essential questions. Apply the current answers. Ask a newly discovered indispensable decision only when actual evidence shows it remains unanswered; never assume an answer or reopen a settled decision. Questions pause before application, each decision once, with the affected pid and a visible paragraph open-item marker. Explain why it is essential in context." : "- Independently assess the review findings against scope_and_decisions and the current answers; a reviewer's suggested question is not automatically a blocker. Leave out findings that need no change or seek optional expansion beyond the requested purpose. Questions pause the run: ask only unresolved decisions needed to fulfil an essential requested requirement, not optional refinements or routine details that can safely remain disclosed placeholders. Put in \"questions\" only information the request, the sources and the answers do not give and only the lawyer can supply; do not guess it. Ask each decision once, even where several paragraphs depend on it. Word each question so the lawyer can answer it without the draft open: one decision per question, in plain words, with abbreviations spelled out and no [pid] or other internal reference; put why it matters in \"context\". Give each question the \"pid\" of the paragraph whose text depends on the answer and, if that paragraph does not already show an open item, amend it so it shows a visible [open item: ...] marker.",
        "- A change that supplies what an open-item marker in the paragraph asks for removes that marker.",
        "- Read-only tools are available; file-writing and redline tools are unavailable. Your final reply is parsed as the change list.",
      ].join("\n")),
    ] : [],
  }
  return (sections[isReviewStage(stage) ? "review" : stage] ?? []).filter(Boolean).join("\n\n")
}

export function retryPrompt(stage: StageID | "advise" | "answer", issues: string[]) {
  return [
    `Your reply did not pass the ${stage} checks:`,
    ...issues.map(issue => `- ${issue}`),
    "",
    "Reply again with the complete corrected output. The output contract above still applies.",
  ].join("\n")
}

/** The standard every reviewer applies for the party the lawyer acts for, by
 * the run's Posture: Client-favorable looks for every protection a careful
 * drafter would include for that party; Balanced and Conservative hold it to
 * what a neutral, mid-market document of the kind, prepared for neither party,
 * would contain (a market document for that side would itself lean). Each finding's
 * balance, not this wording, keeps what goes beyond market out of the draft
 * (review-findings.ts). */
export function sidedReviewStandard(directory: string) {
  const assess = "- Assess each provision as the reviewing lawyer for the party the lawyer acts for: whether it works under the governing law the request and sources point to, "
  return runPosture(directory) === "client-favorable"
    ? `${assess}whether it protects that party's position, and whether a protection a careful drafter would include for that party is missing.`
    : `${assess}whether it has a real adverse effect on that party's position, and whether it lacks a protection for that party that a neutral, mid-market document of this kind, prepared for neither party, would contain.`
}

/** Whom the lawyer acts for, as data every model stage reads. When neither the
 * request nor the matter's recorded client says, the stages take no side. */
export function actingForBlock(directory: string) {
  const actingFor = planActingFor(directory)
  return tagged("acting_for", actingFor.stated
    ? `The lawyer acts for: ${actingFor.party}.`
    : "The request does not say whom the lawyer acts for. Work even-handedly: where the parties' interests diverge, draft wording balanced between them " +
      "marking only a new substantive choice not already settled by instructions or answers for confirmation, rather than favouring either party. " +
      "Ask the lawyer only for information the request and the sources do not give.")
}

/** What the package holds besides this body, from two flags: a memo is
 * requested, and this document is the body of the combined file. The memo is
 * the drafting notes, written once the documents are final; the combined file
 * is built at delivery. Empty without a memo, so the prompt and stage inputs
 * stay as they were. */
export function packageCompositionBlock(directory: string, document: DocID) {
  if (!requestedMemo(directory)) return ""
  return tagged("package_composition", [
    "The request asks for a memo to accompany the documents. The run writes that memo separately, as the drafting notes, after the documents are final; it is not part of this document's body.",
    ...(combinedDelivery(directory)?.document === document ? ["The request asks for that memo and this document in one file. The code builds that file at delivery: the memo first, then this document's clean body."] : []),
    "Do not report the memo as missing from this body, and do not insert memo, cover-note or advice text meant for the memo into this body.",
    "Content this document's own purpose or the request assigns to it still belongs in this body. Advice belongs in this body only when this document's purpose is advice, and a reference to the memo only when the request requires one.",
  ].join("\n"))
}

/** Every planned document by id, title and purpose, marking the one under
 * work, so review and changes do not treat another document's content as
 * missing from this one. Empty when the plan has one document. */
function plannedDocumentsBlock(directory: string, document: DocID) {
  const plan = readLeanPlan(directory)?.documents ?? []
  if (plan.length < 2) return ""
  return tagged("planned_documents", [
    "The request is drafted as these documents, each drafted, reviewed and amended in its own steps:",
    ...plan.map(row => `- ${row.id} "${row.title}"${row.id === document ? " (this document)" : ""}${row.purpose ? `: ${row.purpose}` : ""}`),
  ].join("\n"))
}

/** Earlier final text, accepted explanations, adjudicated dispositions and selected guidance review.
 * Review remains source-navigation context, not approval. Strip its earlier
 * paragraph pointers so a later writer cannot borrow them as amendment targets. */
export function completedDocumentsText(directory: string, document: DocID) {
  return earlierDocuments(directory, document)
    .map(row => {
      const guidanceReview = row.playbook_review?.sources.length ? reviewText(directory, row.id, "review-playbook", { strip: true }) : undefined
      return `<document title="${attribute(row.title)}">\n${finalBodyParagraphs(directory, row.id).map(shownText).join("\n\n")}\n<accepted_change_explanations>\n${JSON.stringify(appliedChangeEvidence(directory, row.id).map(change => ({
      change_origin: change.change_origin, operation: change.operation, deleted_text: change.deleted_text, result_text: change.result_paragraphs.map(row => row.text), reason: change.reason, sources: change.sources, evidence_basis: change.evidence_basis,
      final_text: (paragraph => paragraph ? shownText(paragraph) : "")(finalBodyParagraphs(directory, row.id).find(paragraph => paragraph.pid === change.pid)),
      comment_status: change.native_comment.status,
    })), null, 2)}\n</accepted_change_explanations>\n<review_dispositions>\n${reviewDispositionText(directory, row.id, true)}\n</review_dispositions>${guidanceReview ? `\n<prior_playbook_review>\nEarlier guidance review is a navigation aid, not approval. Verify its conclusions against the canonical guidance, current instructions and final wording. When this document provides advice, explain the material original exposure and proposed effect; preserve conditions, exceptions and proposal status. Earlier paragraph references are not amendment targets in this document.\n${guidanceReview}\n</prior_playbook_review>` : ""}\n</document>`
    })
    .join("\n\n")
}

/** Earlier documents' review findings about this one, shown to its draft. Its
 * change plan adjudicates them; the draft key leaves them out,
 * so a drafted document is not drafted again for them. */
function forwardedBlock(directory: string, document: DocID) {
  const forwarded = forwardedFindings(directory, document)
  return forwarded.length ? tagged("findings_for_this_document", `Reviews of the finished documents found these points for this document. Deal with them where this document's purpose calls for it; its change plan checks them again.\n${findingsText(forwarded, document, { ids: false })}`) : ""
}

function completedDocumentsBlock(directory: string, document: DocID) {
  const text = completedDocumentsText(directory, document)
  return text ? tagged("completed_documents", `These documents of this request are finished, in their final wording. Carry material deferred findings and optional recommendations into the appropriate requested advice, with their origin, practical consequence and remaining action. Affected document ids do not exclude advice about another document. Verify source support; dispositions are decisions, not authority or proof of application. Rejected ideas need not be restated unless material to the advice. Legacy decisions and unadjudicated review history are navigation, not proof of rejection or resolution:\n\n${text}`) : ""
}

function draftingNotesSections(directory: string, request: string) {
  const plan = readLeanPlan(directory)?.documents ?? []
  const items = disclosableItems(directory)
  const planned = new Set(plan.map(row => row.id))
  const loose = items.filter(item => !item.ref.document || !planned.has(item.ref.document))
  const questions = readQuestions(directory)
  const delegated = delegatedDecisionsText(directory)
  const delegatedIDs = new Set(delegatedQuestions(directory).map(question => question.id))
  const memo = requestedMemo(directory)
  // The notes never name a question by its id, so their prompt shows none: an
  // open question is disclosed by its open item's id.
  const answers = answeredQuestions(directory).filter(question => !delegatedIDs.has(question.id)).map(question => `- Q: ${question.question}\n  A: ${question.answer}`).join("\n")
  const provenance = sourceProvenance(directory)
  return [
    memo
      ? `This is a lean drafting run. The request asks for a memo to accompany the drafted documents, "${memo.name}". Write the drafting notes as that memo.`
      : "This is a lean drafting run. Write the drafting notes that go to the lawyer with the drafted documents.",
    tagged("required_repairs", (readRepairRound(directory)?.items ?? []).filter(item => item.stage === "package-review" || item.ref.repair_target === "memo" || notesRepair(item)).map(item => `- ${item.detail}`).join("\n") || "None."),
    request, actingForBlock(directory), answers ? tagged("answers", answers) : "",
    tagged("sources", matterSourceIDs(directory).map(name => nameWithProvenance(name, provenance)).join("\n")),
    tagged("delivery_artifacts", JSON.stringify(deliveryArtifactEvidence(directory), null, 2)),
    tagged("drafting_spine", readLeanText(directory, "spine.json") ?? "None."),
    tagged("discrepancies", readLeanText(directory, "discrepancy.json") ?? "None."),
    ...plan.map(row => `<document id="${row.id}" title="${attribute(row.title)}">\n${documentSections(directory, row.id, items).join("\n\n")}\n</document>`),
    loose.length ? tagged("open_items", openItemLines(loose)) : "",
    tagged("questions", questions.map(question => delegatedIDs.has(question.id)
      ? `- (${question.answer_class === "declined" ? "not answered" : "delegated"}; see delegated_decisions) ${question.question}`
      : `- (${question.status}) ${question.question}${question.answer ? `\n  Answer: ${question.answer}` : ""}`).join("\n") || "None."),
    delegated ? tagged("delegated_decisions", delegated) : "",
    tagged("output_contract", [
      "- Follow the lawyer's explanation-detail, formality and house-style preferences when writing the notes. Translate the supplied records into their practical meaning rather than copying workflow field names, raw paths or stage status. Retain source support, comparison baselines, applied/unapplied outcomes, limitations and every required open-item id. Name documents where this helps locate them.",
      "- Separate actual unresolved instructions, corrections and decisions from professional review before issue. Do not turn settled instructions or internal output-handling steps into new confirmation requirements.",
      "- Retain the material deferred findings and relevant optional recommendations in review_dispositions, explaining their consequence and remaining action in the appropriate advice or notes. Do not describe deferred research as verified law, a recommendation as agreed wording, or a rejected suggestion as an outstanding mandatory change. The final body and applied records determine what was implemented. Legacy decisions are navigation; inspect underlying review and sources when their disposition is unclear.",
      "- Organise the notes by material issues and decisions, tied to the final document and relevant wording. Explain the effect, source basis, applied resolution and remaining action; distinguish agreed terms from choices requiring confirmation.",
      "- Read the final bodies below before making any statement about what they contain, omit or change. The final wording controls; review findings and proposed amendments are history, not proof that a change was applied. Check sources with the available read/search tools whenever the basis is unclear.",
      "- delivery_artifacts projects the bodies, comparisons and generated drafting-notes Word file from the same naming logic registration uses. planned_file, planned_role and planned_filename are the model's selections, not proof of what the user requested; compare them with the instructions. The notes file is the output you are writing, so its pending rendering is normal. A generated note can fulfill an explanatory-note request when its audience, content and naming fit that instruction; a distinct client advice function must be fulfilled by its own appropriate body. Do not describe an explanatory artifact as missing merely because it has no separately planned body. Registration may suffix names to preserve existing matter documents.",
      "- The final body is clean accepted text; native_export describes the separate comparison file, naming plan (registration occurs after package review; this internal phase is not a client open item or outstanding lawyer action), and actual review-amendment application counts and separate withheld-proposal counts. Use that record for artifact presence and applied/unapplied review-amendment counts and the separate withheld-proposal count, and keep source support and legal effect separate. Do not infer export absence from clean body text or fabricate application status from an empty list of explanatory records.\n- Describe each delivered redline using its stated comparison baseline below. Do not imply that template reuse selects an original comparison source. Identify any requested comparison not covered by that baseline as outstanding.",
      "- Explain what each drafted document does. Use change_origin to distinguish source-to-initial-draft revisions from later review amendments; an initial revision is not a change made during review.",
      "- State whom the documents were drafted for or, where the prompt says the request does not say whom the lawyer acts for, that it did not and the drafting is even-handed.",
      ...(memo ? ["- These notes are the requested memo; title it as that memo. Lead with the discrepancies found and the choices made, then give the recommendations and the open points, in plain language. The memo carries only what these records show."] : []),
      "- Say what each open item is and what the lawyer needs to do about it, and end that statement with the item's exact id as a trailing \"[ref: <id>]\". Write internal ids only in these markers, and name each document by its title or file name, never by its internal id; describe each change by what it does, never by its change id.",
      "- For each marked drafted default, say what was chosen and that the lawyer must confirm it.",
      ...(delegated ? ['- delegated_decisions lists the questions the lawyer delegated to the drafter\'s judgment or did not answer. Give them a section of their own. For each decision taken, say in words what was decided, why, the evidence it rests on and the alternatives not taken, that the lawyer delegated it, and that it is for the lawyer to confirm. For each with no decision taken, say that the point stays open, why, and what the lawyer needs to do. End each statement with the question\'s exact id as "[ref: <id>]".'] : []),
      "- Use applied_changes.evidence_basis and the actual sources array to describe a change’s evidential basis. application_mode describes only how text was applied; code or already_applied says nothing about source support. Where source passages are provided, assess them before claiming support or a gap. For no-source-passage changes, disclose drafting judgment or a view of the law and the required lawyer confirmation. Do not invent an open-item marker or confirmation requirement that the final body and evidence do not show.",
      "- operation describes the requested amendment, the only one for its paragraph; result_paragraphs reports the paragraph's actual final wording, and an empty result records deletion. Describe the actual result and disclose unresolved explanations honestly.",
      "- Quote sources only word for word as \"exact words\" [source: document name], and cite only authority supported by matter sources or the approved legal research handoff.",
      "- Reply with the drafting notes in Markdown only, beginning with a \"# \" title.",
      "- Read-only tools are available; file-writing tools are unavailable. Your final reply is saved as the drafting notes.",
    ].join("\n")),
  ]
}

/** Each question the lawyer delegated or did not answer (delegatedQuestions),
 * as the drafting notes disclose it: who delegated it and, for an adopted
 * recommendation, the decision taken with the advisor's checked reason,
 * evidence and the alternatives not taken; otherwise why no decision was
 * taken. Empty when there is none. */
export function delegatedDecisionsText(directory: string) {
  return delegatedQuestions(directory).map(question => [`- [${question.id}] Q: ${question.question}`, ...delegatedLines(question).map(line => `  ${line}`)].join("\n")).join("\n")
}

function delegatedLines(question: Question) {
  const words = question.lawyer_words ?? (question.answer_kind === "adopted" ? undefined : question.answer)
  const who = question.answer_class === "declined" ? "Not answered by the lawyer" : "Delegated by the lawyer"
  const attribution = words === undefined ? `${who}.` : `${who}, in their words: "${words}"`
  const advice = question.advice
  if (question.answer_kind === "adopted") {
    const recommendation = advice && "recommendation" in advice ? advice.recommendation : undefined
    return [attribution, `Decision taken, for the lawyer to confirm: ${question.answer}`,
      ...(recommendation ? [`Reason: ${recommendation.rationale}`, ...recommendation.basis.map(row => `Evidence: "${row.quote}" [source: ${row.source_id}]`)] : []),
      ...questionOptions(question).filter(option => option.id !== question.option_id).map(option => `Alternative not taken: ${option.text}`)]
  }
  const why = question.answer_class === "declined" ? "the point stays open for the lawyer."
    : !delegationSettles(question) ? "it needs a fact or an approval that only the lawyer can give."
    : advice && "unavailable" in advice ? `no option was supported: ${advice.unavailable.reason}`
    : "no checked recommendation is available."
  return [attribution, `No decision taken: ${why}`]
}

function documentSections(directory: string, document: DocID, items: OpenItem[]) {
  const list = readChangeList(directory, document)
  const amendments = new Map((list?.amendments ?? []).map(amendment => [amendment.id, amendment]))
  const apply = readLeanJson(directory, `${document}/apply.json`) as ApplyResult | undefined
  return [
    tagged("comparison_baseline", comparisonDescription(directory, document)),
    tagged("native_export", JSON.stringify(nativeExportEvidence(directory, document), null, 2)),
    tagged("final_body", finalBodyParagraphs(directory, document).map(row => `[${row.pid}] ${shownText(row)}`).join("\n\n")),
    draftingStyleWarningsBlock(styleScopedRows(directory, document, finalBodyParagraphs(directory, document)), true),
    tagged("applied_changes", JSON.stringify(appliedChangeEvidence(directory, document), null, 2)),
    tagged("review_dispositions", reviewDispositionText(directory, document)),
    tagged("unapplied_changes", [
      ...(apply?.unapplied ?? []).map(row => `- ${row.id} [${row.pid}] (${row.reason}): ${amendments.get(row.id)?.reason ?? ""}${row.detail?.issue ? ` Applier: ${row.detail.issue}` : ""}`),
      ...(list?.withheld ?? []).map(row => `- ${row.amendment.id} [${row.amendment.pid}] (withheld, not applied): ${row.amendment.reason}`),
    ].join("\n") || "None."),
    tagged("open_items", openItemLines(items.filter(item => item.ref.document === document)) || "None."),
  ]
}

function openItemLines(items: OpenItem[]) {
  return items.map(item => `- ${item.id}: ${item.detail}`).join("\n")
}

function attribute(text: string) {
  return text.replaceAll("\"", "'")
}

export function tagged(name: string, body: string) {
  return `<${name}>\n${body.trim()}\n</${name}>`
}
