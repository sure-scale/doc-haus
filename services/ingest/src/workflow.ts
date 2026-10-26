import { existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import { fileURLToPath } from "node:url"
import path from "node:path"
import { WORKSPACE_ROOT } from "./matter"
// Cycle with agent.ts (which imports listWorkflows for deleteAgent's dependent
// guard) is safe: both modules only call across the boundary at request time.
import { listAgents } from "./agent"

// Resolved relative to this module so it does not depend on the ingest process cwd.
// DOCHAUS_DIR env override lets tests point at a throwaway temp dir; read lazily
// because bun test loads every test file into one process — a module-load capture
// would freeze whichever file imported first.
export const DOCHAUS_DIR = () =>
  process.env.DOCHAUS_DIR ??
  path.join(path.dirname(fileURLToPath(import.meta.url)), "..", "..", "..", "dochaus")

// The workflow-builder chat runs in this directory (engine sessions are directory-scoped,
// mirroring how template-builder runs in TEMPLATES_DIR).
export const WORKFLOWS_DIR = path.join(WORKSPACE_ROOT, ".workflows")

const REGISTRY_FILE = () => path.join(DOCHAUS_DIR(), "workflows.json")
const AGENT_DIR = () => path.join(DOCHAUS_DIR(), "agent")

// A dochaus/agent/<name>.md that overrides a name the opencode engine reserves at
// config merge would silently hijack that built-in agent. Block them here. Shared
// with agent.ts — custom specialist agents land in the same namespace.
export const RESERVED_NAMES = ["auto", "build", "plan", "general", "explore", "summary", "title", "compaction", "triage"]

// A lowercase, hyphenated slug that names a file or folder, never a Windows
// device name (safe-file-name.ts).
export const NAME_RE = /^(?!(?:con|prn|aux|nul|com[0-9]|lpt[0-9])$)[a-z0-9][a-z0-9-]*$/

export type WorkflowStep = {
  agent: string
  instructions?: string
}

export type Workflow = {
  name: string
  label: string
  description: string
  scope: "matter" | "document"
  prompt: string
  steps: WorkflowStep[]
  created_at: number
}

class WorkflowError extends Error {
  status: number
  constructor(message: string, status: number) {
    super(message)
    this.status = status
  }
}

function readRegistry(): Workflow[] {
  if (!existsSync(REGISTRY_FILE())) return []
  return JSON.parse(readFileSync(REGISTRY_FILE(), "utf8")) as Workflow[]
}

function writeRegistry(records: Workflow[]) {
  writeFileSync(REGISTRY_FILE(), JSON.stringify(records, null, 2))
}

function agentPath(name: string) {
  return path.join(AGENT_DIR(), `${name}.md`)
}

function humanize(agent: string) {
  return agent
    .split("-")
    .map((w) => w.charAt(0).toUpperCase() + w.slice(1))
    .join(" ")
}

export function renderAgentMarkdown(record: Workflow): string {
  const description = record.description.trim() || `Custom workflow: ${record.label}`
  const frontmatter = [
    "---",
    // JSON-quoted so a colon or newline in the description cannot inject YAML keys.
    `description: ${JSON.stringify(description)}`,
    "mode: primary",
    "temperature: 0.2",
    "color: primary",
    "tools:",
    '  "*": false',
    "  read: true",
    "  task: true",
    "  search-document: true",
    "---",
  ].join("\n")

  const intro = [
    `You are the doc.haus "${record.label}" workflow orchestrator. You run a multi-agent`,
    `workflow by delegating to specialist subagents through the \`task\` tool, then`,
    `synthesize their output into one report. You do not perform the analysis`,
    `yourself — you coordinate it.`,
  ].join("\n")

  const pipelineLines = ["<pipeline>", "Run these steps in order, threading each result into the next:", ""]
  record.steps.forEach((step, i) => {
    const title = humanize(step.agent)
    let line = `${i + 1}. **${title}** — call \`task\` with \`subagent_type: "${step.agent}"\`.`
    if (i === 0) {
      line += " In the prompt, state the documents/matter and any focus from the user."
    } else {
      line += " Pass the output of every prior step verbatim in the prompt."
    }
    if (step.instructions && step.instructions.trim()) {
      line += ` ${step.instructions.trim()}`
    }
    line += " Capture its output."
    pipelineLines.push(line)
  })
  pipelineLines.push("</pipeline>")

  const outputLines = [
    "<output>",
    "Return a single combined report with clearly-labeled sections, preserving the",
    "citations each subagent produced:",
    "",
  ]
  record.steps.forEach((step) => {
    outputLines.push(`- **${humanize(step.agent)}**`)
  })
  outputLines.push(
    "- **Bottom line** — your own 2-4 sentence synthesis of where the workflow nets out.",
    "",
    "Do not drop or rewrite the subagents' citations.",
    "</output>",
  )

  return [frontmatter, "", intro, "", pipelineLines.join("\n"), "", outputLines.join("\n")].join("\n")
}

function slugify(label: string) {
  return label.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "")
}

function validate(input: { label: string; description: string; scope: string; prompt: string; steps: WorkflowStep[] }) {
  const name = slugify(input.label)
  if (!name || !NAME_RE.test(name)) throw new WorkflowError("label slugifies to an empty or invalid name", 400)
  if (!input.steps.length) throw new WorkflowError("steps must be non-empty", 400)
  // A step naming a nonexistent or disabled subagent would break the workflow at
  // every launch — the inverse of deleteAgent's dependent-workflow guard.
  const enabled = listAgents().filter((a) => a.enabled).map((a) => a.name)
  for (const step of input.steps) {
    if (!NAME_RE.test(step.agent)) throw new WorkflowError(`step agent "${step.agent}" is not a valid name`, 400)
    if (!enabled.includes(step.agent))
      throw new WorkflowError(
        `step agent "${step.agent}" is not an enabled subagent — valid agents: ${enabled.join(", ")}`,
        400,
      )
  }
  return name
}

export function listWorkflows(): Workflow[] {
  return readRegistry()
}

export function createWorkflow(input: {
  label: string
  description: string
  scope: "matter" | "document"
  prompt: string
  steps: WorkflowStep[]
}): Workflow {
  const name = validate(input)
  if (RESERVED_NAMES.includes(name)) throw new WorkflowError(`"${name}" is a reserved agent name`, 409)
  const registry = readRegistry()
  if (registry.some((r) => r.name === name)) throw new WorkflowError(`workflow "${name}" already exists`, 409)
  if (existsSync(agentPath(name))) throw new WorkflowError(`agent file "${name}.md" already exists`, 409)
  const record: Workflow = { name, label: input.label, description: input.description, scope: input.scope, prompt: input.prompt, steps: input.steps, created_at: Date.now() }
  mkdirSync(AGENT_DIR(), { recursive: true })
  writeFileSync(agentPath(name), renderAgentMarkdown(record))
  writeRegistry([...registry, record])
  return record
}

export function updateWorkflow(
  name: string,
  input: {
    label: string
    description: string
    scope: "matter" | "document"
    prompt: string
    steps: WorkflowStep[]
  },
): Workflow {
  validate(input)
  const registry = readRegistry()
  const existing = registry.find((r) => r.name === name)
  if (!existing) throw new WorkflowError(`workflow "${name}" not found`, 404)
  const record: Workflow = { ...existing, label: input.label, description: input.description, scope: input.scope, prompt: input.prompt, steps: input.steps }
  writeFileSync(agentPath(name), renderAgentMarkdown(record))
  writeRegistry(registry.map((r) => (r.name === name ? record : r)))
  return record
}

export function deleteWorkflow(name: string) {
  const registry = readRegistry()
  const idx = registry.findIndex((r) => r.name === name)
  if (idx === -1) throw new WorkflowError(`workflow "${name}" not found`, 404)
  rmSync(agentPath(name), { force: true })
  writeRegistry(registry.filter((_, i) => i !== idx))
}
