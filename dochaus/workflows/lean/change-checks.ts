import { amendmentSchema, amendmentParagraphs, amendmentText, normalizeParagraphText, type Amendment } from "../../lib/drafting-amendment"
import { normalize } from "../../lib/quote-verify"
import { markerIssues } from "./change-markers"
import { record, type ChangeList, type ParagraphRow, type StageCheck } from "./checks"
import { paragraphShowsOpenItem } from "./gate"
import { approvedResearch } from "./research"
import { REQUEST_SOURCE, excerpt, matterSourceIDs, openItem, quoteFound, quoteMissIssue, sourceQuotes } from "./quotes"
import type { DocID } from "./plan"
import { readQuestions, shortHash, type Question } from "./state"

// Admission of the changes step's reply: the change list is parsed, checked
// against the draft's paragraphs, its quotations and its markers, and either
// fed back to the model or, on the last attempt, admitted with what still
// fails withheld as open items.


/** Parses and validates the model's change list. Values code can derive (the
 * amendment id and the paragraph's current hash) are filled in here; the model
 * never restates them. A quotation that the named source does not contain, or
 * markers that do not show the lawyer the change (`markerIssues`), are retry
 * feedback; on the last allowed attempt (`withholdUnverified`) that amendment is
 * withheld instead, unapplied and recorded as an open item. An
 * unanswered question that names no paragraph, or whose paragraph will not show
 * it as an open item, is also retry feedback; on the last attempt it is recorded
 * as asked and settled by the gate as before. An amendment that takes out or
 * cuts down a template part is not an issue here: its reason is required, and
 * the apply step records the removal with that reason for the drafting notes. */
export function checkChanges(directory: string, content: string, paragraphs: ParagraphRow[], options: { withholdUnverified?: boolean; document?: DocID } = {}):
  StageCheck & { changes?: ChangeList } {
  const parsed = parseJsonObject(content)
  if (!parsed.ok)
    return { issues: [`The reply must be one JSON object {"amendments": [...], "questions": [...]}: ${parsed.error}`], openItems: [] }
  const value = parsed.value
  const rows = new Map(paragraphs.map(row => [row.pid, row]))
  const sources = new Set([...matterSourceIDs(directory), REQUEST_SOURCE, ...approvedResearch(directory).map(row => row.url)])
  const asked = new Map(readQuestions(directory).map(question => [question.id, question]))
  const amendments = Array.isArray(value.amendments) ? value.amendments : undefined
  const explanations = value.explanations === undefined ? [] : Array.isArray(value.explanations) ? value.explanations : undefined
  const questions = value.questions === undefined ? [] : Array.isArray(value.questions) ? value.questions : undefined
  const questionPids = new Set((questions ?? []).filter(record).flatMap(item => typeof item.pid === "string" ? [item.pid.trim()] : []))
  const issues = [
    ...(!amendments ? ['"amendments" must be an array (it may be empty).'] : []),
    ...(!explanations ? ['"explanations" must be an array when present.'] : []),
    ...(explanations ?? []).flatMap((item, index) => explanationIssues(directory, item, index, rows, sources)),
    ...(!questions ? ['"questions" must be an array when present.'] : []),
    ...(amendments ?? []).flatMap((item, index) => amendmentIssues(item, index, rows, sources)),
    ...duplicateAmendmentIssues(amendments ?? []),
    ...(questions ?? []).flatMap((item, index) => questionIssues(item, index, rows, asked)),
  ]
  if (issues.length || !amendments || !questions || !explanations) return { issues, openItems: [] }
  const filled: Amendment[] = amendments.filter(record).map((item, index) => ({
    id: `a${index + 1}`,
    pid: String(item.pid),
    existing_sha256: rows.get(String(item.pid))!.sha256,
    ...(item.operation !== undefined ? { operation: item.operation as Amendment["operation"] } : {}),
    updated: Array.isArray(item.updated) ? item.updated.map(text => String(text).trim()) : String(item.updated).trim(),
    reason: String(item.reason).trim(),
    sources: (item.sources as Record<string, unknown>[]).map(source => ({ source_id: String(source.source_id).trim(), quote: String(source.quote) })),
  }))
  // The applier validates the same contract; agree with it here so a defect is
  // fed back to the authoring model instead of failing the apply stage.
  const contract = amendmentSchema.array().safeParse(filled)
  if (!contract.success)
    return { issues: contract.error.issues.slice(0, 10).map(issue => `amendments.${issue.path.join(".")}: ${issue.message}`), openItems: [] }
  const unverified = filled.map((amendment, index) => ({
    amendment,
    issues: [...markerIssues(amendment, index, rows.get(amendment.pid)!.text, questionPids), ...amendmentQuoteIssues(directory, amendment, index)],
  })).filter(row => row.issues.length)
  const proposed = askedQuestions(questions.filter(record), asked, options.document)
  const unmarked = proposed.flatMap((question, index) => question ? questionMarkerIssues(question, index, filled, rows) : [])
  if ((unverified.length || unmarked.length) && !options.withholdUnverified)
    return { issues: [...unverified.flatMap(row => row.issues), ...unmarked], openItems: [] }
  const withheld = new Set(unverified.map(row => row.amendment.id))
  const kept = filled.filter(amendment => !withheld.has(amendment.id))
  const questionList = proposed.filter((question): question is Question => Boolean(question))
  // A question whose paragraph only a withheld amendment would have marked: the
  // paragraph keeps its drafted text, so, like the withheld change, the
  // question is disclosed through the drafting notes.
  const carriers = new Map(questionList.flatMap(question => {
    if (!questionMarkerIssues(question, 0, kept, rows).length) return []
    const carrier = unverified.find(row => row.amendment.pid === question.pid && paragraphShowsOpenItem(amendmentText(row.amendment)))
    return carrier ? [[question.id, carrier.amendment.id] as const] : []
  }))
  const shared = [...Map.groupBy(kept, amendment => amendment.pid)]
    .filter(([, group]) => group.length > 1)
    .map(([pid, group]) => ({ pid, amendment_ids: group.map(amendment => amendment.id) }))
  const explained = explanations.filter(record).map((item, index) => ({
    id: `i${index + 1}`, pid: String(item.pid), reason: String(item.reason).trim(),
    sources: (item.sources as { source_id: string; quote: string }[]).map(source => ({ source_id: source.source_id.trim(), quote: source.quote })),
  }))
  return {
    issues: [],
    changes: { ...(value.explanations !== undefined ? { explanations: explained } : {}), amendments: kept, withheld: unverified, questions: questionList, shared_paragraphs: shared },
    openItems: [
      ...unverified.map(row => openItem("changes", "unapplied_amendment", [
        `Amendment ${row.amendment.id} to paragraph [${row.amendment.pid}] was withheld and not applied: it did not pass the change checks after the allowed retries.`,
        `Proposed text: "${excerpt(amendmentText(row.amendment))}" Reason given: ${excerpt(row.amendment.reason)}`,
        `Problems: ${row.issues.join(" ")} The paragraph keeps its drafted text.`,
      ].join(" "), { amendment_id: row.amendment.id, pid: row.amendment.pid })),
      ...questionList.map(question => {
        const carrier = carriers.get(question.id)
        return openItem("changes", "question", `Question for the lawyer: ${question.question}` +
          (carrier ? ` Its paragraph keeps its drafted text without a marker: amendment ${carrier}, which would have marked it, was withheld.` : ""),
        { question_id: question.id, ...(question.pid ? { pid: question.pid } : {}), ...(carrier ? { amendment_id: carrier } : {}) })
      }),
    ],
  }
}

/** An explanation annotates an existing revision, never a no-op amendment.
 * Exact passages are checked here; their entailment and coverage remain semantic review. */
function explanationIssues(directory: string, item: unknown, index: number, rows: Map<string, ParagraphRow>, sources: Set<string>) {
  const at = `explanations[${index}]`
  if (!record(item)) return [`${at} must be an object {pid, reason, sources}.`]
  return [
    ...(typeof item.pid !== "string" || !rows.has(item.pid) ? [`${at}.pid is not a paragraph of the draft; select the paragraph explaining this initial revision.`] : []),
    ...(typeof item.reason !== "string" || !item.reason.trim() ? [`${at}.reason must explain the actual revision and its consequence.`] : []),
    ...(!Array.isArray(item.sources) ? [`${at}.sources must be a list of {source_id, quote}.`] : item.sources.flatMap((source, n) => {
      if (!record(source) || typeof source.source_id !== "string" || !sources.has(source.source_id.trim()) || typeof source.quote !== "string" || !source.quote.trim())
        return [`${at}.sources[${n}] must name a canonical source and a non-empty exact passage.`]
      return quoteFound(directory, { source_id: source.source_id.trim(), quote: source.quote }) ? [] : [quoteMissIssue(directory, `${at}.sources[${n}]`, { source_id: source.source_id.trim(), quote: source.quote })]
    })),
    ...(typeof item.reason === "string" ? sourceQuotes(item.reason).filter(quote => !quoteFound(directory, quote)).map(quote => quoteMissIssue(directory, `${at}.reason`, quote)) : []),
  ]
}

function amendmentIssues(item: unknown, index: number, rows: Map<string, ParagraphRow>, sources: Set<string>) {
  const at = `amendments[${index}]`
  if (!record(item)) return [`${at} must be an object {pid, updated, reason, sources}.`]
  const row = typeof item.pid === "string" ? rows.get(item.pid) : undefined
  const updated = typeof item.updated === "string" ? item.updated.trim() : ""
  const split = item.operation === "split" && Array.isArray(item.updated) && item.updated.length >= 2 && item.updated.every(text => typeof text === "string" && text.trim() && !/[\r\n]/.test(text))
  const deletion = item.operation === "delete" && item.updated === ""
  const sourceList = Array.isArray(item.sources) ? item.sources : undefined
  return [
    ...(!row ? [`${at}.pid ${JSON.stringify(item.pid ?? null)} is not a paragraph of the draft; use a [pid] value listed with the draft.`] : []),
    ...(!deletion && !split && !updated ? [`${at}.updated must be the full replacement text; use operation delete with updated \"\" for an explicit deletion, or operation split with an array of at least two nonempty paragraphs.`] : []),
    ...(item.operation !== undefined && !["replace", "delete", "split"].includes(String(item.operation)) ? [`${at}.operation must be replace, delete or split.`] : []),
    ...(/[\r\n]/.test(updated) ? [`${at}.updated must be one paragraph without line breaks; use operation split with an array of resulting paragraphs for a deliberate split.`] : []),
    ...(row && updated && normalizeParagraphText(updated) === normalizeParagraphText(row.text) ? [`${at}.updated is identical to the current paragraph; drop it or change the text.`] : []),
    ...(typeof item.reason !== "string" || !item.reason.trim() ? [`${at}.reason must say why the paragraph changes.`] : []),
    ...(!sourceList ? [`${at}.sources must be a list of {source_id, quote}.`] : []),
    ...(sourceList ?? []).flatMap((source, sourceIndex) => {
      const where = `${at}.sources[${sourceIndex}]`
      if (!record(source) || typeof source.source_id !== "string" || !source.source_id.trim() || typeof source.quote !== "string" || !source.quote.trim())
        return [`${where} must be {source_id, quote} with non-empty strings.`]
      return sources.has(source.source_id.trim()) ? [] : [`${where}.source_id ${JSON.stringify(source.source_id)} is not a matter document or an approved research URL; use an exact source name, or \"${REQUEST_SOURCE}\" for the request and the lawyer's answers.`]
    }),
  ]
}

function duplicateAmendmentIssues(amendments: unknown[]) {
  const keys = amendments.filter(record).map(item => `${item.pid}\0${item.operation ?? "replace"}\0${JSON.stringify((Array.isArray(item.updated) ? item.updated : [String(item.updated)]).map(text => normalizeParagraphText(String(text))))}`)
  const repeated = [...Map.groupBy(keys, key => key)].filter(([, group]) => group.length > 1).map(([key]) => key.split("\0")[0])
  return repeated.length ? [`Duplicate amendments for paragraphs ${repeated.map(pid => `[${pid}]`).join(", ")}; keep one per change.`] : []
}

function amendmentQuoteIssues(directory: string, amendment: Amendment, index: number) {
  const at = `amendments[${index}] ([${amendment.pid}])`
  return [
    ...amendment.sources.filter(source => !quoteFound(directory, source))
      .map(source => quoteMissIssue(directory, `${at}.sources`, source)),
    ...sourceQuotes(amendmentText(amendment)).filter(quote => !quoteFound(directory, quote))
      .map(quote => quoteMissIssue(directory, `${at}.updated`, quote)),
  ]
}

export function questionIssues(item: unknown, index: number, rows: Map<string, ParagraphRow>, asked: Map<string, Question>) {
  const at = `questions[${index}]`
  if (!record(item) || typeof item.question !== "string" || !item.question.trim())
    return [`${at} must be {question, pid, context?, id?} with a non-empty question.`]
  return [
    ...(item.context !== undefined && typeof item.context !== "string" ? [`${at}.context must be text.`] : []),
    ...(given(item.id) && (typeof item.id !== "string" || !asked.has(item.id.trim()))
      ? [`${at}.id ${JSON.stringify(item.id)} is not a question asked so far; omit "id" for a new question.`] : []),
    ...(given(item.pid) && (typeof item.pid !== "string" || !rows.has(item.pid.trim()))
      ? [`${at}.pid ${JSON.stringify(item.pid)} is not a paragraph of the draft; use the [pid] listed with the draft for the paragraph the question affects.`] : []),
  ]
}

/** The gate discloses an unanswered question only through an open item shown
 * in the paragraph it affects, and only an amendment from this stage can put one
 * there: apply and finalize never add body text. The paragraph is checked as it
 * will read after the change list: the text of its amendments, or else its
 * current text. */
function questionMarkerIssues(question: Question, index: number, amendments: Amendment[], rows: Map<string, ParagraphRow>) {
  const at = `questions[${index}]`
  const row = question.pid ? rows.get(question.pid) : undefined
  if (!row)
    return [`${at} is tied to no paragraph of the draft. Set its "pid" to the paragraph whose text depends on the answer, ` +
      "and, if that paragraph does not already show an open item, amend it so it shows a visible [open item: what the lawyer must decide] marker."]
  const matching = amendments.filter(amendment => amendment.pid === row.pid)
  const amended = matching.flatMap(amendmentParagraphs)
  if ((matching.length ? amended : [row.text]).some(text => paragraphShowsOpenItem(text))) return []
  return [`${at} affects paragraph [${row.pid}], but that paragraph will not show the open item. Add an amendment to [${row.pid}] ` +
    "(or change your amendment to it) whose updated text puts a visible [open item: what the lawyer must decide] marker there, " +
    "citing the source passage the paragraph rests on."]
}

/** A repeated question keeps its id, so rewording it does not raise a new
 * question; a question already answered, by id or by its exact wording, is
 * never asked again (undefined at its index). A repeat that names no paragraph
 * keeps the earlier paragraph only when it was asked about the same document.
 * The final gate's question stays the gate's: a repeat of its id is dropped. */
export function askedQuestions(items: Record<string, unknown>[], asked: Map<string, Question>, document: DocID | undefined): (Question | undefined)[] {
  const answered = new Set([...asked.values()].filter(question => question.status === "answered").map(question => normalize(question.question)))
  return items.map(item => {
    const text = String(item.question).trim()
    const prior = given(item.id) ? asked.get(String(item.id).trim()) : undefined
    if (prior?.status === "answered" || prior?.stage === "gate" || answered.has(normalize(text))) return
    const pid = given(item.pid) ? String(item.pid).trim() : prior?.document === document ? prior?.pid : undefined
    return {
      id: prior?.id ?? `q-${shortHash(normalize(text))}`,
      stage: "changes" as const,
      question: text,
      context: typeof item.context === "string" ? item.context.trim() : "",
      ...(pid ? { pid } : {}),
      status: "open" as const,
    }
  })
}

function given(value: unknown) {
  return value !== undefined && value !== null && value !== ""
}

export function parseJsonObject(content: string): { ok: true; value: Record<string, unknown> } | { ok: false; error: string } {
  // Prefer the whole value so an array containing an object cannot be mistaken
  // for the stage's object contract. Never select one of multiple results.
  try {
    const value: unknown = JSON.parse(content.trim())
    return record(value) ? { ok: true, value } : { ok: false, error: "the JSON value is not an object" }
  } catch { /* A prose wrapper can still contain one complete JSON object. */ }
  const start = content.indexOf("{")
  const end = content.lastIndexOf("}")
  if (start < 0 || end < start) return { ok: false, error: "no JSON object found" }
  const blocks = [...content.matchAll(/```(?:json)?\s*([\s\S]*?)```/g)]
  if (blocks.length > 1) return { ok: false, error: "multiple fenced results found; return one JSON object" }
  const candidate = blocks[0]?.[1]?.trim() ?? content.slice(start, end + 1)
  if (!blocks.length && (content.slice(0, start).trimEnd().endsWith("[") || content.slice(end + 1).trimStart().startsWith("]")))
    return { ok: false, error: "the JSON value is not an object" }
  try {
    const value: unknown = JSON.parse(candidate)
    if (blocks.length && record(value) && content.slice(start, end + 1) !== candidate)
      return { ok: false, error: "ambiguous content outside the fenced result; return one JSON object" }
    return record(value) ? { ok: true, value } : { ok: false, error: "the JSON value is not an object" }
  } catch (error) {
    return { ok: false, error: error instanceof Error ? error.message : "invalid JSON" }
  }
}
