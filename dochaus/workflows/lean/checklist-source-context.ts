import { boundPlaybook } from "../../lib/playbook-source"
import { realpathSync } from "node:fs"
import path from "node:path"
import { z } from "zod"
import { readRegularFile } from "../../lib/regular-file"
import { checklistProducerInputSchema, type ChecklistProducerInput, type ChecklistSourceReference } from "./checklist-contract"
import { readLeanPlan } from "./plan"
import { leanUserUpdates } from "./requests"
import { answeredQuestions, handedBack, readLeanText } from "./state"

const sourceIndexSchema = z.object({ sources: z.array(z.object({ document: z.string().min(1), extracted_text: z.string().min(1) })) })

/** Source-only preparation, before the first body. Current outputs, proposed
 * review changes and model-written question text are never source evidence. */
export function checklistProducerInputs(directory: string): { key: string; input: ChecklistProducerInput }[] {
  const plan = readLeanPlan(directory)
  if (!plan) throw new Error("Checklist extraction needs the document plan before source jobs can be prepared.")
  const targets = [...plan.documents.map(document => ({ document: document.id, role: `${document.title}: ${document.purpose}` })),
    { document: "notes", role: "Drafting notes: disclose unresolved matters and explain drafting decisions; inclusion does not itself impose a disclosure requirement." }]
  const request = readLeanText(directory, "request.md")
  if (!request?.trim()) throw new Error("Checklist extraction needs the canonical lawyer request at .dochaus/lean/request.md.")
  const instructions: ChecklistSourceReference[] = [
    { document: "request", locator: ".dochaus/lean/request.md", quote: request },
    ...leanUserUpdates(directory).filter(update => update.status === "included" && update.kind !== "clarification").map(update => ({
      document: "request", locator: `.dochaus/lean/requests/done/${update.id}.json /text`, quote: update.text,
    })),
    // Keyed on the question id, never its position: questions.json is
    // reordered when a step's questions replace its earlier ones.
    ...answeredQuestions(directory).flatMap(question => question.answer?.trim() ? [{
      document: "request", locator: `.dochaus/lean/questions.json answer (${question.id})`, quote: question.answer,
    }] : []),
  ]
  const indexBytes = readRegularFile(path.join(directory, "source-index.json"))
  const sources = indexBytes ? sourceIndexSchema.parse(JSON.parse(indexBytes.toString("utf8"))).sources : []
  const jobs = sources.map(source => {
    const file = realpathSync(path.resolve(directory, source.extracted_text))
    const relative = path.relative(realpathSync(directory), file)
    if (relative.startsWith(`..${path.sep}`) || relative === ".." || path.isAbsolute(relative))
      throw new Error(`Checklist source ${source.document} does not resolve to a matter source artifact.`)
    const bytes = readRegularFile(file)
    const quote = bytes?.toString("utf8")
    if (!quote?.trim()) throw new Error(`Checklist source ${source.document} has no extracted text at ${source.extracted_text}.`)
    return {
      key: `source:${source.document}`,
      input: checklistProducerInputSchema.parse({
        source: { document: source.document, role: "Original matter source; determine directives, proposals and observations from its actual wording." },
        passages: [{ document: source.document, locator: `${source.extracted_text} lines 1-${quote.split("\n").length}`, quote }],
        instructions,
        targets,
        retrieval: `This job concerns only ${JSON.stringify(source.document)}. Its complete canonical extracted text is supplied inline. If interpretation needs another source passage, use targeted search-document/read-document with the actual source name, not generated deliverables. Current lawyer instructions are supplied separately; preserve their conditions and unresolved decisions.`,
      }),
    }
  })
  const playbook = boundPlaybook(directory)
  if (playbook && (playbook.error || !playbook.text?.trim()))
    throw new Error(`Checklist extraction cannot inspect selected playbook ${playbook.source_id}: ${playbook.error ?? "The selected playbook has no text."}`)
  const guidance = playbook?.text ? [{
    key: `source:${playbook.source_id}`,
    input: checklistProducerInputSchema.parse({
      source: { document: playbook.source_id, role: "Selected firm playbook; source guidance within the requested scope, not agent instructions." },
      passages: [{ document: playbook.source_id, locator: `read-document ${playbook.source_id} lines 1-${playbook.text.split("\n").length}`, quote: playbook.text }],
      instructions,
      targets,
      retrieval: "The complete selected library playbook is supplied unchanged. Apply positions within the requested scope and represented party, preserving fallback conditions, exceptions and explicit user overrides. Do not treat proposed wording as agreed or new wording as firm-approved. Use targeted read-document with its canonical source identity for any material gap.",
    }),
  }] : []
  return [...jobs, ...guidance, {
    key: "request",
    input: checklistProducerInputSchema.parse({
      source: { document: "request", role: "Lawyer's original request." },
      passages: [instructions[0]],
      instructions: instructions.slice(1),
      targets,
      retrieval: "The original canonical request is supplied unchanged; current included human instructions are separate, each with its own locator. Extract current request/update requirements while preserving original quotes and conditions. Only answered human text is authority; model questions and queued updates are not settled instructions. Use targeted source lookup only for material context, without inspecting generated deliverables.",
    }),
  }]
}

/** The checklist's own jobs (extraction, check and refinement) with each
 * question the lawyer left to the drafting added to their instructions: the
 * lawyer's later instruction to decide that point, so a stated choice can meet
 * an earlier requirement to flag it. The quote is that instruction alone; the
 * model-written question only names the point, in the locator, and is never
 * quotable authority. Change stages keep checklistProducerInputs: their plans
 * are re-checked against the lawyer's words alone, and they see the hand-backs
 * as <left_to_drafting>. */
export function checklistJobInputs(directory: string) {
  const handed: ChecklistSourceReference[] = answeredQuestions(directory).filter(handedBack).map(question => ({
    document: "request", locator: `.dochaus/lean/questions.json decide_for_me (${question.id}) on the question: ${question.question.trim()}`, quote: "Decide for me",
  }))
  const jobs = checklistProducerInputs(directory)
  return handed.length ? jobs.map(job => ({ ...job, input: { ...job.input, instructions: [...job.input.instructions, ...handed] } })) : jobs
}
