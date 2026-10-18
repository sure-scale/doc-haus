import { existsSync, mkdirSync, readdirSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import path from "node:path"
import { workspaceRoot } from "./matter"
import { DOCHAUS_DIR, NAME_RE, RESERVED_NAMES, listWorkflows } from "./workflow"
import { disabledAgents, setAgentDisabled } from "./engine-config"

// The firm's custom specialist agents: subagents composed conversationally by the
// agent-builder and written as dochaus/agent/<name>.md, the same namespace the
// repo-shipped specialists live in. A registry (dochaus/agents.json) records which
// agents are custom, mirroring workflows.json — the .md is the engine artifact,
// the registry is the source of truth for the library UI. Custom agents are
// mode: subagent with a fixed read-only research toolset, so they immediately
// become available as workflow pipeline steps and `task` targets.

// The agent-builder chat runs in this directory (engine sessions are
// directory-scoped, mirroring WORKFLOWS_DIR).
export const AGENTS_DIR = () => path.join(workspaceRoot(), ".agents")

const REGISTRY_FILE = () => path.join(DOCHAUS_DIR(), "agents.json")
const AGENT_DIR = () => path.join(DOCHAUS_DIR(), "agent")

export type CustomAgent = {
  name: string
  label: string
  description: string
  // The specialist's task instructions — the body the builder composed. The
  // frontmatter and citation/output discipline are rendered around it here.
  instructions: string
  created_at: number
}

export type AgentSummary = {
  name: string
  label: string
  description: string
  instructions: string
  builtin: boolean
  enabled: boolean
  created_at: number
}

class AgentError extends Error {
  status: number
  constructor(message: string, status: number) {
    super(message)
    this.status = status
  }
}

function readRegistry(): CustomAgent[] {
  if (!existsSync(REGISTRY_FILE())) return []
  return JSON.parse(readFileSync(REGISTRY_FILE(), "utf8")) as CustomAgent[]
}

function writeRegistry(records: CustomAgent[]) {
  writeFileSync(REGISTRY_FILE(), JSON.stringify(records, null, 2))
}

function agentPath(name: string) {
  return path.join(AGENT_DIR(), `${name}.md`)
}

export function renderCustomAgentMarkdown(record: CustomAgent): string {
  const description = record.description.trim() || `Custom specialist: ${record.label}`
  const frontmatter = [
    "---",
    // JSON-quoted so a colon or newline in the description cannot inject YAML keys.
    `description: ${JSON.stringify(description)}`,
    "mode: subagent",
    "temperature: 0.2",
    "color: info",
    "tools:",
    '  "*": false',
    "  read: true",
    "  glob: true",
    "  grep: true",
    "  list: true",
    "  skill: true",
    "  search-document: true",
    "  cite: true",
    "---",
  ].join("\n")

  const intro = [
    `You are the doc.haus "${record.label}" specialist subagent. You analyze a matter's`,
    `documents for the focus described below and report findings grounded in the`,
    `actual text. The task prompt tells you which documents to cover and any extra focus.`,
  ].join("\n")

  const citation = [
    "<citation>",
    "- Cite every finding as `[<Document> § <section>]` with the supporting excerpt",
    "  quoted verbatim. Never invent a section number or quote.",
    "- Anchor every quoted excerpt with the `cite` tool before it appears in a finding,",
    "  passing the verbatim quote, a `reason`, and a `confidence` (1-5). Never quote a",
    "  passage `cite` failed to verify.",
    "</citation>",
  ].join("\n")

  const output = [
    "<output>",
    "A list of findings. For each: a one-line headline, the citation + excerpt, and a",
    "short explanation of why it matters. No preamble.",
    "</output>",
  ].join("\n")

  return [frontmatter, "", intro, "", "<task>", record.instructions.trim(), "</task>", "", citation, "", output].join(
    "\n",
  )
}

function slugify(label: string) {
  return label.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "")
}

// Frontmatter of a hand-authored dochaus/agent/*.md — enough to list the
// repo-shipped specialists beside the custom ones (same parse as list-workflows).
function parseFrontmatter(src: string): Record<string, string> {
  const lines = src.split("\n")
  if (lines[0]?.trim() !== "---") return {}
  const end = lines.findIndex((l, i) => i > 0 && l.trim() === "---")
  if (end === -1) return {}
  const result: Record<string, string> = {}
  for (const line of lines.slice(1, end)) {
    const colon = line.indexOf(":")
    if (colon === -1) continue
    result[line.slice(0, colon).trim()] = line
      .slice(colon + 1)
      .trim()
      .replace(/^"|"$/g, "")
  }
  return result
}

function humanize(name: string) {
  return name
    .split("-")
    .map((w) => w.charAt(0).toUpperCase() + w.slice(1))
    .join(" ")
}

function stripFrontmatter(src: string) {
  const lines = src.split("\n")
  if (lines[0]?.trim() !== "---") return src.trim()
  const end = lines.findIndex((l, i) => i > 0 && l.trim() === "---")
  if (end === -1) return src.trim()
  return lines.slice(end + 1).join("\n").trim()
}

// All specialist subagents: the repo-shipped ones (read from their .md, builtin)
// plus the registry's custom ones. Workflow orchestrators are mode: primary and
// the registry's own files are skipped by name, so neither double-lists.
export function listAgents(): AgentSummary[] {
  const disabled = disabledAgents()
  const custom = readRegistry()
  const customNames = new Set(custom.map((a) => a.name))
  const builtins = !existsSync(AGENT_DIR())
    ? []
    : readdirSync(AGENT_DIR())
        .filter((f) => f.endsWith(".md") && !customNames.has(f.slice(0, -3)))
        .flatMap((f) => {
          const src = readFileSync(path.join(AGENT_DIR(), f), "utf8")
          const fm = parseFrontmatter(src)
          if (fm["mode"] !== "subagent" || fm["hidden"] === "true") return []
          const name = f.slice(0, -3)
          return [
            {
              name,
              label: humanize(name),
              description: fm["description"] ?? "",
              instructions: stripFrontmatter(src),
              builtin: true,
              enabled: !disabled.has(name),
              created_at: 0,
            },
          ]
        })
  return [
    ...builtins.sort((a, b) => a.name.localeCompare(b.name)),
    ...custom
      .map((a) => ({ ...a, builtin: false, enabled: !disabled.has(a.name) }))
      .sort((a, b) => a.name.localeCompare(b.name)),
  ]
}

// Flip an agent on or off — builtin and custom alike. Off sets
// `agent.<name>.disable` in dochaus/opencode.json, the same flag the config
// already uses for the engine's stock agents, so the engine drops it from the
// task/workflow roster; the .md and registry entry stay put.
export function setAgentEnabled(name: string, enabled: boolean): AgentSummary {
  const agent = listAgents().find((a) => a.name === name)
  if (!agent) throw new AgentError(`agent "${name}" not found`, 404)
  // Disabling an agent a workflow step still names would break that workflow at
  // launch (and updateWorkflow's enabled-agent check would reject re-saving it) —
  // refuse while any workflow references it, same as deleteAgent.
  if (!enabled) {
    const dependents = listWorkflows()
      .filter((w) => w.steps.some((s) => s.agent === name))
      .map((w) => w.name)
    if (dependents.length)
      throw new AgentError(`agent "${name}" is used by workflow(s): ${dependents.join(", ")} — update them first`, 409)
  }
  setAgentDisabled(name, !enabled)
  return { ...agent, enabled }
}

function validate(input: { label: string; instructions: string }) {
  const name = slugify(input.label)
  if (!name || !NAME_RE.test(name)) throw new AgentError("label slugifies to an empty or invalid name", 400)
  if (!input.instructions.trim()) throw new AgentError("instructions must be non-empty", 400)
  return name
}

export function createAgent(input: { label: string; description: string; instructions: string }): CustomAgent {
  const name = validate(input)
  if (RESERVED_NAMES.includes(name)) throw new AgentError(`"${name}" is a reserved agent name`, 409)
  const registry = readRegistry()
  if (registry.some((r) => r.name === name)) throw new AgentError(`agent "${name}" already exists`, 409)
  if (existsSync(agentPath(name))) throw new AgentError(`agent file "${name}.md" already exists`, 409)
  const record: CustomAgent = {
    name,
    label: input.label,
    description: input.description,
    instructions: input.instructions,
    created_at: Date.now(),
  }
  mkdirSync(AGENT_DIR(), { recursive: true })
  writeFileSync(agentPath(name), renderCustomAgentMarkdown(record))
  writeRegistry([...registry, record])
  return record
}

export function updateAgent(name: string, input: { label: string; description: string; instructions: string }): CustomAgent {
  validate(input)
  const registry = readRegistry()
  const existing = registry.find((r) => r.name === name)
  if (!existing) throw new AgentError(`agent "${name}" not found`, 404)
  const record: CustomAgent = {
    ...existing,
    label: input.label,
    description: input.description,
    instructions: input.instructions,
  }
  writeFileSync(agentPath(name), renderCustomAgentMarkdown(record))
  writeRegistry(registry.map((r) => (r.name === name ? record : r)))
  return record
}

export function deleteAgent(name: string) {
  const registry = readRegistry()
  const idx = registry.findIndex((r) => r.name === name)
  if (idx === -1) throw new AgentError(`agent "${name}" not found`, 404)
  // A workflow step that names this agent would break at launch — refuse while
  // any workflow still references it.
  const dependents = listWorkflows()
    .filter((w) => w.steps.some((s) => s.agent === name))
    .map((w) => w.name)
  if (dependents.length)
    throw new AgentError(`agent "${name}" is used by workflow(s): ${dependents.join(", ")} — update them first`, 409)
  rmSync(agentPath(name), { force: true })
  writeRegistry(registry.filter((_, i) => i !== idx))
  // Drop any disable flag so a later agent reusing the name starts enabled.
  setAgentDisabled(name, false)
}
