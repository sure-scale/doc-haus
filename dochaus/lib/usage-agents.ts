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

// The assistants and jobs the lawyer starts, by the name the lawyer knows them by.
const WORK: Record<string, string> = {
  [UPKEEP_AGENT]: "Conversation titles and compaction",
  [PREPARE_CHANGES_AGENT]: "Prepare changes",
  [REPLY_DRAFTING_AGENT]: "Comment replies",
  router: "Choose an assistant",
  "document-reviewer": "Word edit review",
  chronology: "Build a chronology",
  "chronology-extract": "Build a chronology",
  "chronology-match": "Build a chronology",
  qa: "Questions",
  "word-chat": "Questions in Word",
  redliner: "Redline",
  redactor: "Redact",
  research: "Research",
  "quick-draft": "Quick draft",
  "client-summary": "Client summary",
  compare: "Compare",
  obligations: "Obligations",
  summarizer: "Summaries",
  "legal-review": "Full review",
  "playbook-reviewer": "Full review",
  "assumption-challenger": "Full review",
  extract: "Review table",
  "review-suggest": "Review table",
  dashboard: "Build dashboards",
  "dashboard-chart": "Build dashboards",
  "playbook-importer": "Playbook import",
  "template-builder": "Template builder",
  "workflow-builder": "Workflow builder",
  "skill-builder": "Skill builder",
  "agent-builder": "Agent builder",
}

// A drafting run's steps, grouped by the task each serves so the lawyer sees a
// few plain tasks rather than one row per internal step.
const DRAFTING_TASK: Record<string, string> = {
  draft: "Drafting",
  "lean-instructions": "Planning",
  "lean-intake": "Planning",
  "lean-discrepancy": "Planning",
  "lean-act-record": "Planning",
  "lean-advise": "Planning",
  "lean-answer": "Planning",
  "lean-answer-reconcile": "Planning",
  "lean-spine": "Planning",
  "lean-spine-split": "Planning",
  "lean-spine-level": "Planning",
  "lean-placement": "Planning",
  "lean-coverage": "Planning",
  "lean-files": "Planning",
  "lean-research": "Research",
  "lean-research-review": "Research",
  "lean-draft": "Drafting",
  "lean-drafting-notes": "Drafting",
  "lean-review": "Review",
  "lean-review-independent": "Review",
  "lean-review-playbook": "Review",
  "lean-review-wording": "Review",
  "lean-review-client-disadvantage": "Review",
  "lean-package-review": "Review",
  "lean-change-plan": "Review",
  "lean-change-prepare": "Review",
  "lean-changes": "Review",
  "lean-change-review": "Review",
  "lean-spine-absence": "Checks",
  "lean-spine-absence-verify": "Checks",
  "lean-checklist-extract": "Checks",
  "lean-checklist-check": "Checks",
  "lean-reader-check": "Checks",
  "lean-notes-check": "Checks",
}

// Usage under a name with no plain label: never shown as its internal id.
export const OTHER_WORK = "Other work"

type SessionLike = { metadata?: Record<string, unknown> | null } | undefined

// Word suggestions run on the Q&A agent; their session metadata names the kind.
export function usageAgentKey(agent: string | undefined, session?: SessionLike): string {
  if (session?.metadata?.kind === "suggestions") return WORD_SUGGESTIONS_AGENT
  return agent || "unknown"
}

export function usageAgent(agent: string | undefined, session?: SessionLike): UsageAgent {
  const key = usageAgentKey(agent, session)
  if (BACKGROUND[key]) return { key, label: BACKGROUND[key], group: "background" }
  return { key, label: DRAFTING_TASK[key] ?? WORK[key] ?? OTHER_WORK, group: "work" }
}

export const USAGE_GROUP_LABEL: Record<UsageGroup, string> = { background: "Background agents", work: "Your work" }

// Sentence case for a tool id.
export function humanize(key: string) {
  const words = key.replace(/[-_]+/g, " ").trim()
  return words ? words[0]!.toUpperCase() + words.slice(1) : "Other"
}
