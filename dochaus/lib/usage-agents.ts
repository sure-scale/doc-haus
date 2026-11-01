// Plain names for usage reports. Background agents start without a click and
// match the Settings > Background agents switches; everything else is work the
// lawyer asked for, including the checks that run inside it.
export type UsageGroup = "background" | "work"
export type UsageAgent = { key: string; label: string; group: UsageGroup }

// Usage with no agent of its own: conversation titles and compaction (native
// session cost not carried by any message), and stateless generation jobs.
export const UPKEEP_AGENT = "conversation-upkeep"
export const PREPARE_CHANGES_AGENT = "prepare-changes"
export const REPLY_DRAFTING_AGENT = "comment-replies"
export const WORD_SUGGESTIONS_AGENT = "word-suggestions"

const BACKGROUND: Record<string, string> = {
  "chronology-assess": "Suggest a chronology",
  "review-compare": "Compare review answers",
  narrator: "Narrate drafting progress",
  [WORD_SUGGESTIONS_AGENT]: "Suggest ideas in Word",
}

const WORK: Record<string, string> = {
  [UPKEEP_AGENT]: "Conversation titles and compaction",
  [PREPARE_CHANGES_AGENT]: "Prepare changes",
  [REPLY_DRAFTING_AGENT]: "Comment replies",
  router: "Choose an assistant",
  "lean-advise": "Question check",
  "document-reviewer": "Word edit review",
  "chronology-extract": "Build a chronology",
  "chronology-match": "Build a chronology",
  qa: "Questions",
  extract: "Review table",
}

type SessionLike = { metadata?: Record<string, unknown> | null } | undefined

// Word suggestions run on the Q&A agent; their session metadata names the kind.
export function usageAgentKey(agent: string | undefined, session?: SessionLike): string {
  if (session?.metadata?.kind === "suggestions") return WORD_SUGGESTIONS_AGENT
  return agent || "unknown"
}

export function usageAgent(agent: string | undefined, session?: SessionLike): UsageAgent {
  const key = usageAgentKey(agent, session)
  if (BACKGROUND[key]) return { key, label: BACKGROUND[key], group: "background" }
  return { key, label: WORK[key] ?? humanize(key), group: "work" }
}

export const USAGE_GROUP_LABEL: Record<UsageGroup, string> = { background: "Background agents", work: "Your work" }

// Sentence case for an id with no plain name of its own (agents, tools).
export function humanize(key: string) {
  const words = key.replace(/^lean-/, "").replace(/[-_]+/g, " ").trim()
  return words ? words[0]!.toUpperCase() + words.slice(1) : "Other"
}
