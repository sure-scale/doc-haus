import { tool } from "./legacy-definition"
import { readFileSync, readdirSync, existsSync } from "node:fs"
import { fileURLToPath } from "node:url"
import path from "node:path"

// doc.haus list-workflows tool. Returns two things the workflow-builder needs:
// (1) the registered custom workflows from the ingest service, and (2) the
// available subagent building blocks from dochaus/agent/*.md so the builder can
// compose pipelines from real agent names only.

const ingestUrl = process.env.INGEST_URL ?? "http://127.0.0.1:4500"
const agentDir = path.join(path.dirname(fileURLToPath(import.meta.url)), "..", "agent")

function parseFrontmatter(src: string): Record<string, string> {
  const lines = src.split("\n")
  if (lines[0].trim() !== "---") return {}
  const end = lines.findIndex((l, i) => i > 0 && l.trim() === "---")
  if (end === -1) return {}
  const result: Record<string, string> = {}
  for (const line of lines.slice(1, end)) {
    const colon = line.indexOf(":")
    if (colon === -1) continue
    result[line.slice(0, colon).trim()] = line.slice(colon + 1).trim().replace(/^"|"$/g, "")
  }
  return result
}

export default tool({
  description:
    "List the firm's existing custom workflows (name, label, description, scope, and steps) AND the available specialist subagents that can be used as pipeline steps. Always call this first before creating or updating a workflow — it is the only way to discover real subagent names. Never invent a subagent name.",
  args: {},
  async execute() {
    const res = await fetch(`${ingestUrl}/workflows`)
    const { workflows } = (await res.json()) as {
      dir: string
      workflows: {
        name: string
        label: string
        description: string
        scope: string
        prompt: string
        steps: { agent: string; instructions?: string }[]
      }[]
    }

    const subagents: { name: string; description: string }[] = []
    if (existsSync(agentDir)) {
      for (const file of readdirSync(agentDir).filter((f) => f.endsWith(".md"))) {
        const src = readFileSync(path.join(agentDir, file), "utf8")
        const fm = parseFrontmatter(src)
        if (fm["mode"] !== "subagent") continue
        if (fm["hidden"] === "true") continue
        subagents.push({ name: file.slice(0, -3), description: fm["description"] ?? "" })
      }
    }

    const workflowLines =
      workflows.length === 0
        ? "No custom workflows yet."
        : workflows
            .map((w) => {
              const steps = w.steps
                .map((s, i) => `  ${i + 1}. ${s.agent}${s.instructions ? ` — ${s.instructions}` : ""}`)
                .join("\n")
              return `[${w.name}] ${w.label} (${w.scope})\n  ${w.description}\n  Launch prompt: ${w.prompt}\n  Steps:\n${steps}`
            })
            .join("\n\n")

    const subagentLines =
      subagents.length === 0
        ? "No subagents available."
        : subagents.map((s) => `  ${s.name} — ${s.description}`).join("\n")

    return {
      title: `${workflows.length} workflow(s), ${subagents.length} subagent(s)`,
      output: `Custom workflows:\n${workflowLines}\n\nAvailable subagents (use exact names as step agents):\n${subagentLines}`,
      metadata: {
        workflows: workflows.map((w) => w.name),
        subagents: subagents.map((s) => s.name),
      },
    }
  },
})
