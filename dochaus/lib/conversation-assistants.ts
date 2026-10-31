/** Assistant identities and routing descriptions shared by every chat host. */
export const DOCUMENT_CHAT_ASSISTANTS = [
  { name: "qa", label: "Q&A", description: "Answers questions about this matter's documents, always with citations." },
  { name: "redliner", label: "Redline", description: "Makes the edits you ask for in this matter's documents, as tracked changes you accept or reject in Word." },
] as const
