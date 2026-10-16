import type { Part } from "../../api/session-view"
import { CHAT_ASSISTANTS, TEMPLATE_BUILDER } from "../../agents"

export function humanizeTool(tool: string) {
  return tool.charAt(0).toUpperCase() + tool.slice(1).replace(/-/g, " ")
}

// The agent name as a byline on an answer. Prefer the registry's own label (so
// "qa" reads "Q&A" and "redliner" reads "Redline"); anything off-registry
// humanizes cleanly ("redline" -> "Redline").
const AGENT_LABELS: Record<string, string> = Object.fromEntries(
  [...CHAT_ASSISTANTS, TEMPLATE_BUILDER].map((a) => [a.name, a.label]),
)
export const agentLabel = (name: string) => AGENT_LABELS[name] ?? humanizeTool(name)

// Verbs that turn a raw tool name + its target into a readable action line
// ("Read Source Document.docx"), so the timeline narrates each step the
// way a familiar legal-agent UI does rather than dumping the tool's own title.
const TOOL_VERBS: Record<string, string> = {
  read: "Read",
  write: "Wrote",
  edit: "Edited",
  list: "Listed",
  glob: "Searched files",
  grep: "Searched",
  bash: "Ran",
  redline: "Redlined",
}

export const basename = (s: string) => s.split("/").filter(Boolean).pop() ?? s

// The matter's internal bookkeeping files are plumbing, not legal work — name
// them in plain language instead of leaking raw filenames into the timeline.
const INTERNAL_FILES: Record<string, string> = {
  "matter.json": "Reviewed the matter details",
  "grid.json": "Reviewed the document grid",
  "profile.md": "Reviewed the practice profile",
}

// The reviewers a workflow spawns, named by legal role rather than subagent id.
const SUBAGENT_ROLES: Record<string, string> = {
  "document-reviewer": "Reviewer",
  "playbook-reviewer": "Playbook",
  "assumption-challenger": "Challenger",
  summarizer: "Summary",
  compare: "Compare",
  obligations: "Obligations",
}

// One readable label for a tool step. search-document is the matter's core
// retrieval call, so it reads as a search for its query (not the raw passage
// count); file tools read as a verb plus the file's basename; the matter folder
// and its internal files read in plain language; anything else falls back to a
// cleaned-up title.
export function stepLabel(p: Extract<Part, { type: "tool" }>, matter: string): string {
  const input = "input" in p.state ? (p.state.input as Record<string, unknown> | undefined) : undefined
  const title = p.state.status === "completed" || p.state.status === "running" ? p.state.title : undefined

  if (p.tool === "search-document") {
    const query = (input?.query as string | undefined) ?? title?.match(/"([^"]+)"/)?.[1]
    if (!query) return "Searching documents"
    const doc = input?.document as string | undefined
    return doc ? `Searched ${basename(doc)} for "${query}"` : `Searched documents for "${query}"`
  }

  // A workflow coordinates its reviewers through the task tool; name each by its
  // legal role rather than the raw subagent id so the timeline reads like a team.
  if (p.tool === "task") {
    const role = input?.subagent_type as string | undefined
    return role ? `Consulted the ${SUBAGENT_ROLES[role] ?? humanizeTool(role)}` : "Coordinated the review"
  }

  // webfetch's upstream title is "<url> (<content-type>)"; the generic basename
  // fallback would split on the slash inside "text/html" and label the step
  // with a content-type fragment. Name it by the host it fetched instead.
  if (p.tool === "webfetch") {
    const url = (input?.url as string | undefined) ?? title?.split(" ")[0] ?? ""
    return URL.canParse(url) ? `Fetched ${new URL(url).hostname}` : "Fetched a web page"
  }

  const verb = TOOL_VERBS[p.tool]
  const path = [input?.filePath, input?.path].find((value) => typeof value === "string" && value)
  const target = title || (typeof path === "string" ? path : undefined)
  if (!target) return humanizeTool(p.tool)
  const name = basename(target)
  if (name === matter) return "Browsed the matter files"
  if (INTERNAL_FILES[name]) return INTERNAL_FILES[name]
  return verb ? `${verb} ${name}` : name
}
