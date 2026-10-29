/** Matter-local chronology contracts. Dates describe the source's precision;
 * evidence verification is a text match, never a finding that an event is true. */
export type ChronologyPrivilege = "all-cleared" | "mixed" | "stop"
export type ChronologyDate = {
  original: string
  precision: "exact" | "approximate" | "month" | "year" | "ambiguous" | "relative" | "undated"
  start?: string
  end?: string
  rangeKind?: "duration" | "uncertainty"
  time?: string
  timezone?: string
  qualifier?: string
  documentDate?: string
  relative?: {
    expression: string
    anchorDocument: string
    anchorDate?: string
    offsetDays?: number
    calculation?: string
    interpretation?: "document-stated" | "legal-interpretation"
  }
  correction?: { original: ChronologyDate; at: number; actor: string }
}
export type ChronologyEvidence = {
  id: string
  documentName: string
  sourceId: string
  version: string
  excerpt: string
  section: string
  charStart: number
  charEnd: number
  verified: boolean
  freshness: "current" | "stale" | "unavailable" | "unverified"
  provenance: "source" | "supplied-chronology"
  sourceKind?: string
  issue?: string
}
export type ChronologyAccount = {
  id: string
  title: string
  detail: string
  date: ChronologyDate
  evidence: ChronologyEvidence[]
  attribution?: string
  privilege: "cleared" | "uncertain" | "privileged" | "not-privileged"
  significance: "key" | "relevant" | "background"
  people: string[]
  status: "reported" | "scheduled" | "required"
  obligationId?: string
  gaps: string[]
}
export type ChronologyEvent = {
  id: string
  title: string
  detail: string
  date: ChronologyDate
  accounts: ChronologyAccount[]
  origin: "generated" | "manual"
  original?: { title: string; detail: string; date: ChronologyDate }
  reviewed: boolean
  reviewedAt?: number
  reviewedBy?: string
  opened: boolean
  hidden: boolean
  key: boolean
  disputed: boolean
  dispute?: string
  note: string
  lawyerEdited: boolean
  createdAt: number
  updatedAt: number
  previousIds: string[]
  supersededBy?: string[]
}
export type ChronologyCoverage = {
  sourceId: string
  documentName: string
  version: string
  status: "pending" | "processing" | "complete" | "partial" | "failed" | "unreadable" | "unsupported" | "removed"
  uploaded: boolean
  textExtracted: boolean
  sections: number
  processedSections: number[]
  completedKeys: string[]
  events: number
  issues: string[]
  suppliedChronology: boolean
}
export type ChronologyProposal = {
  id: string
  kind: "addition" | "amendment" | "conflict" | "removal" | "merge" | "split"
  targetIds: string[]
  baseEventHashes: Record<string, string>
  /** Removal decisions expire when the authorised source set changes. */
  sourceSetKey?: string
  events: ChronologyEvent[]
  reason: string
  recommendation: string
  status: "pending" | "deferred" | "accepted" | "kept"
  origin: "refresh" | "chat" | "source-date"
  createdAt: number
  decidedAt?: number
  actor?: string
}
export type ChronologyRun = {
  id: string
  mode: "initial" | "refresh"
  status: "running" | "stopped" | "failed" | "complete"
  privilege: Exclude<ChronologyPrivilege, "stop">
  startedAt: number
  completedAt?: number
  model?: string
  sourceKeys: string[]
  completedKeys: string[]
  errors: string[]
  usage?: { input: number; output: number; cost?: number }
}
export type ChronologySuggestion = {
  status: "unknown" | "relevant" | "irrelevant" | "insufficient" | "hidden" | "later"
  reason: string
  assessedKey?: string
  dismissedKey?: string
  assessedAt?: number
}
export type ChronologyHistory = {
  id: string
  at: number
  actor: string
  action: string
  eventIds: string[]
  detail?: string
}
export type Chronology = {
  schemaVersion: 1
  revision: number
  created: boolean
  events: ChronologyEvent[]
  coverage: ChronologyCoverage[]
  proposals: ChronologyProposal[]
  sourceDates: Record<string, { date: string; at: number; actor: string }>
  suggestion: ChronologySuggestion
  run?: ChronologyRun
  history: ChronologyHistory[]
  framing: string
  gaps: string[]
}
export type ChronologyAction =
  | { type: "open"; eventId: string }
  | { type: "review"; eventId: string; reviewed: boolean }
  | { type: "edit"; eventId: string; title: string; detail: string; date: ChronologyDate }
  | { type: "note"; eventId: string; note: string }
  | { type: "key"; eventId: string; key: boolean }
  | { type: "hide"; eventId: string; hidden: boolean }
  | { type: "dispute"; eventId: string; dispute: string; disputed: boolean }
  | { type: "add"; title: string; detail: string; date: ChronologyDate }
  | { type: "merge"; eventIds: string[]; title: string; detail: string; date: ChronologyDate; preview?: boolean }
  | {
      type: "split"
      eventId: string
      parts: { title: string; detail: string; date: ChronologyDate; accountIds: string[] }[]
      preview?: boolean
    }
  | { type: "source-date"; documentName: string; date: string }
  | { type: "invitation"; choice: "hide" | "later" | "show" }
  | { type: "discard-refresh" }
export type ChronologyDecision = "accept" | "keep" | "defer"

export function chronologyNeedsAttention(event: ChronologyEvent) {
  return (
    event.disputed ||
    (event.origin === "generated" && event.accounts.length === 0) ||
    event.accounts.some(
      (a) => a.evidence.length === 0 || a.evidence.some((e) => !e.verified || e.freshness !== "current"),
    )
  )
}
export function chronologyCounters(state: Chronology) {
  const visible = state.events.filter((e) => !e.hidden)
  return {
    new: visible.filter((e) => !e.opened).length,
    unreviewed: visible.filter((e) => !e.reviewed).length,
    attention: visible.filter(chronologyNeedsAttention).length,
    updates: state.proposals.filter((p) => p.status === "pending" || p.status === "deferred").length,
    sources: state.coverage.filter((c) => c.status !== "complete").length,
  }
}

/** A model-supplied day count is usable only when the recorded expression
 * supplies the same explicit calendar-day rule. Business/calendar conventions
 * and legal interpretations need a lawyer rather than a guessed conversion. */
export function chronologyCalendarDayOffset(relative: NonNullable<ChronologyDate["relative"]>): number | undefined {
  const expression = relative.expression.toLocaleLowerCase()
  if (
    relative.interpretation === "legal-interpretation" ||
    /business|working|court|statut|under\s+(?:the\s+)?law|except|excluding|inclusive|exclusive|weekend|holiday|clear\s+days|hundred|thousand|million|\bor\b|\bbetween\b/.test(
      expression,
    )
  )
    return
  const numeric = expression.match(
    /\b(\d+|(?:zero|one|two|three|four|five|six|seven|eight|nine|ten|eleven|twelve|thirteen|fourteen|fifteen|sixteen|seventeen|eighteen|nineteen|twenty|thirty|forty|fifty|sixty|seventy|eighty|ninety)(?:[ -](?:one|two|three|four|five|six|seven|eight|nine))?)\s+(?:calendar\s+)?days?\s+(after|before|from|of)\b/,
  )
  if (!numeric || relative.offsetDays === undefined) return
  const words: Record<string, number> = {
    zero: 0,
    one: 1,
    two: 2,
    three: 3,
    four: 4,
    five: 5,
    six: 6,
    seven: 7,
    eight: 8,
    nine: 9,
    ten: 10,
    eleven: 11,
    twelve: 12,
    thirteen: 13,
    fourteen: 14,
    fifteen: 15,
    sixteen: 16,
    seventeen: 17,
    eighteen: 18,
    nineteen: 19,
    twenty: 20,
    thirty: 30,
    forty: 40,
    fifty: 50,
    sixty: 60,
    seventy: 70,
    eighty: 80,
    ninety: 90,
  }
  const count = /^\d+$/.test(numeric[1]!)
    ? Number(numeric[1])
    : numeric[1]!.split(/[ -]/).reduce((sum, word) => sum + (words[word] ?? NaN), 0)
  const days = count * (numeric[2] === "before" ? -1 : 1)
  return days === relative.offsetDays && Math.abs(days) <= 36525 ? days : undefined
}
