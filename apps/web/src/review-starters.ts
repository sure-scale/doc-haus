import { COMPARE_AGREEMENTS_EXAMPLE, type StarterExample } from "./review-starters-example"

// Review-table starters are definitions only: the selected matter supplies all
// documents and the grid remains the sole store for questions and answers. A
// starter's example is fictional, read-only and never runs a model.

export type ReviewStarterDefinition = {
  id: string
  name: string
  purpose: string
  inputs: string
  output: string
  limits: string
  questions: readonly string[]
  /** Answers are keyed by the index of `questions`. */
  example?: StarterExample
}

export const REVIEW_STARTERS: readonly ReviewStarterDefinition[] = [
  {
    id: "document-overview-v1",
    name: "Document overview",
    purpose: "Start a review table with questions about each document in this matter.",
    inputs: "All documents in this matter and the questions below.",
    output: "A review table for answers and sources for each document and question.",
    limits: "Answers may be missing or incomplete. Check the sources and tailor the questions to your matter; this is not a complete legal review.",
    questions: [
      "What is the purpose of this document?",
      "What obligations are described in this document?",
      "What dates or deadlines are stated in this document?",
    ],
  },
  {
    id: "compare-agreements-v1",
    name: "Compare agreements",
    purpose: "Compare the key terms of the agreements in this matter side by side.",
    inputs: "All documents in this matter and the questions below.",
    output: "A review table with an answer and source for each agreement and question. After Fill, answers that differ between agreements are marked.",
    limits: "Answers and comparisons may be missing or incomplete. Check each answer against its source and tailor the questions to your matter; this is not a complete legal review.",
    questions: [
      "Who are the parties, and when was the agreement made?",
      "Does the agreement renew at the end of its term, and how can renewal be stopped?",
      "When can each party terminate the agreement, and what notice is required?",
      "What limits and exclusions apply to each party's liability?",
      "Which law governs the agreement, and how are disputes resolved?",
    ],
    example: COMPARE_AGREEMENTS_EXAMPLE,
  },
]

export function normaliseStarterQuestion(question: string) {
  return question.trim().replace(/\s+/g, " ").toLocaleLowerCase()
}

// Recommended from the lawyer's own documents only: generated deliverables are
// not agreements to compare. No model call; the counts are already loaded.
export function recommendStarter(documents: readonly { role?: "source" | "deliverable" }[]): { id: string; reason: string } {
  const sources = documents.filter((doc) => doc.role !== "deliverable").length
  const kind = sources === documents.length ? "" : "source "
  const reason = sources ? `${sources} ${kind}${sources === 1 ? "document" : "documents"} in this matter` : `No ${kind}documents in this matter yet`
  return { id: sources >= 2 ? "compare-agreements-v1" : "document-overview-v1", reason }
}
