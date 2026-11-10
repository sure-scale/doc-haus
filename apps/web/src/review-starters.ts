import { COMPARE_AGREEMENTS_EXAMPLE, type StarterExample } from "./review-starters-example"
import type { Document } from "./api/ingest/documents"

// Review-table starters are definitions only: the selected matter supplies all
// documents and the grid remains the sole store for questions and answers. A
// starter's example is fictional, read-only and never runs a model.

// `compare: false` keeps a question out of the AI comparison after Fill: its
// answers describe or identify each document, so they differ by nature. The
// flag stays with the question when the lawyer edits its text, and on its column.
export type StarterQuestion = { text: string; compare?: false }

export type ReviewStarterDefinition = {
  id: string
  name: string
  purpose: string
  inputs: string
  output: string
  limits: string
  questions: readonly StarterQuestion[]
  /** Answers are keyed by the index of `questions`. */
  example?: StarterExample
}

export const REVIEW_STARTERS: readonly ReviewStarterDefinition[] = [
  {
    id: "document-overview-v1",
    name: "Document overview",
    purpose: "Start a review table with questions about each document in this matter.",
    inputs: "All documents in this matter and this starter's questions.",
    output: "A review table with an answer and source for each document and question.",
    limits: "Answers may be missing or incomplete. Check the sources and tailor the questions to your matter; this is not a complete legal review.",
    // Each answer describes its own document, so none is compared.
    questions: [
      { text: "What is the purpose of this document?", compare: false },
      { text: "What obligations are described in this document?", compare: false },
      { text: "What dates or deadlines are stated in this document?", compare: false },
    ],
  },
  {
    id: "compare-agreements-v1",
    name: "Compare agreements",
    purpose: "Compare the key terms of the agreements in this matter side by side.",
    inputs: "All documents in this matter and this starter's questions.",
    output: "A review table with an answer and source for each agreement and question. After Fill, terms that differ between agreements are marked.",
    limits: "Answers and comparisons may be missing or incomplete. Check each answer against its source and tailor the questions to your matter; this is not a complete legal review.",
    // The parties and date identify each agreement, so they are not compared.
    questions: [
      { text: "Who are the parties, and when was the agreement made?", compare: false },
      { text: "Does the agreement renew at the end of its term, and how can renewal be stopped?" },
      { text: "When can each party terminate the agreement, and what notice is required?" },
      { text: "What limits and exclusions apply to each party's liability?" },
      { text: "Which law governs the agreement, and how are disputes resolved?" },
    ],
    example: COMPARE_AGREEMENTS_EXAMPLE,
  },
]

export function normaliseStarterQuestion(question: string) {
  return question.trim().replace(/\s+/g, " ").toLocaleLowerCase()
}

// Discovery searches the descriptions already shown to the lawyer. It changes
// the view only, without reading documents or asking a model to rank them.
export function workflowSearchMatches(query: string, ...description: string[]) {
  const text = description.join(" ").toLocaleLowerCase()
  return query.trim().toLocaleLowerCase().split(/\s+/).every((word) => text.includes(word))
}

// Only report what ingest actually knows; missing quality metadata is unknown.
export function reviewInputNotice(documents: readonly Pick<Document, "name" | "quality">[]) {
  if (!documents.length) return "Add documents to this matter before adding questions or filling answers. You can still inspect a starter and its fictional example."
  const unreadable = documents.filter((doc) => doc.quality?.status === "no_readable_text").map((doc) => doc.name)
  const degraded = documents.filter((doc) => doc.quality?.status === "degraded").map((doc) => doc.name)
  return [
    unreadable.length ? `No readable text was extracted from: ${unreadable.join(", ")}. Add a readable version before relying on answers from these documents.` : "",
    degraded.length ? `Extracted text may be incomplete in: ${degraded.join(", ")}. Check the original documents when reviewing answers.` : "",
  ].filter(Boolean).join(" ") || undefined
}

// Recommended from the lawyer's own documents only: generated deliverables are
// not agreements to compare. No model call; the counts are already loaded.
export function recommendStarter(documents: readonly { role?: "source" | "deliverable" }[]): { id: string; reason: string } {
  const sources = documents.filter((doc) => doc.role !== "deliverable").length
  const kind = sources === documents.length ? "" : "source "
  const reason = sources ? `${sources} ${kind}${sources === 1 ? "document" : "documents"} in this matter` : `No ${kind}documents in this matter yet`
  return { id: sources >= 2 ? "compare-agreements-v1" : "document-overview-v1", reason }
}
