import { INGEST_URL } from "../../config"
import { ingestError } from "./errors"
import type { Playbook } from "./matters"

export async function listPlaybookLibrary(): Promise<{ dir: string; playbooks: Playbook[] }> {
  const res = await fetch(`${INGEST_URL}/playbooks/library`)
  if (!res.ok) throw new Error(await ingestError(res, "Could not load playbooks"))
  return res.json()
}

export async function uploadPlaybookSource(file: File): Promise<{ name: string; source: string }> {
  const form = new FormData()
  form.append("file", file)
  const res = await fetch(`${INGEST_URL}/playbooks/source`, { method: "POST", body: form })
  if (!res.ok) throw new Error(await ingestError(res, "Could not upload this playbook"))
  return res.json()
}

export async function savePlaybook(input: { name: string; description: string; content: string }, create: boolean): Promise<void> {
  const res = await fetch(`${INGEST_URL}/playbooks${create ? "" : `/${encodeURIComponent(input.name)}`}`, {
    method: create ? "POST" : "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(input),
  })
  if (!res.ok) throw new Error(await ingestError(res, "Could not save this playbook"))
}

export async function deletePlaybook(name: string): Promise<void> {
  const res = await fetch(`${INGEST_URL}/playbooks/${encodeURIComponent(name)}`, { method: "DELETE" })
  if (!res.ok) throw new Error(await ingestError(res, "Could not delete this playbook"))
}

// A drafting template in the firm's global library. placeholders are the
// bracketed terms a draft fills in (text, with a kind and an optional hint).
export type TemplatePlaceholder = { text: string; kind: string; hint?: string }
// description is a one-line summary of what the template is for, stored in the
// ingest-owned manifest beside the .docx (empty string when none was set).
export type Template = { name: string; description: string; placeholders: TemplatePlaceholder[] }

export type WorkflowStep = { agent: string; instructions: string }

export type CustomWorkflow = {
  name: string
  label: string
  description: string
  scope: "matter" | "document"
  prompt: string
  steps: WorkflowStep[]
  created_at: number
}

// The firm's global template library, owned by the ingest service. Templates are
// shared across every matter, so unlike documents they are not scoped to a matter.
// The response carries the library directory alongside the list — the templates
// chat scopes its engine sessions to that directory (there is no matter to scope
// to), so the page needs it from the same call.
export async function listTemplates(): Promise<{ dir: string; templates: Template[] }> {
  const res = await fetch(`${INGEST_URL}/templates`)
  if (!res.ok) throw new Error(await ingestError(res, "Could not load templates"))
  return res.json()
}

export async function uploadTemplate(file: File): Promise<Template> {
  const form = new FormData()
  form.append("file", file)
  const res = await fetch(`${INGEST_URL}/templates`, { method: "POST", body: form })
  if (!res.ok) throw new Error(await ingestError(res, `Could not upload ${file.name}`))
  return res.json()
}

export async function updateTemplateDescription(name: string, description: string): Promise<Template> {
  const res = await fetch(`${INGEST_URL}/templates?name=${encodeURIComponent(name)}`, {
    method: "PATCH",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ description }),
  })
  if (!res.ok) throw new Error(`Could not update ${name} (${res.status})`)
  return res.json()
}

export async function deleteTemplate(name: string): Promise<void> {
  const res = await fetch(`${INGEST_URL}/templates?name=${encodeURIComponent(name)}`, { method: "DELETE" })
  if (!res.ok) throw new Error(await ingestError(res, `Could not delete ${name}`))
}

// Raw .docx bytes for a template, rendered client-side by the template viewer via
// WASM — the file is fetched from our own ingest service and converted in the
// browser, so it never leaves for any third-party service.
export async function fetchTemplateBytes(name: string): Promise<Uint8Array> {
  const res = await fetch(`${INGEST_URL}/templates/content?name=${encodeURIComponent(name)}`)
  if (!res.ok) throw new Error(`Could not load ${name} (${res.status})`)
  return new Uint8Array(await res.arrayBuffer())
}

// The firm's custom workflow library, owned by the ingest service. The response
// carries the library directory alongside the list — the workflow builder chat
// scopes its engine sessions to that directory (there is no matter to scope to),
// so the page needs it from the same call.
export async function listWorkflows(): Promise<{ dir: string; workflows: CustomWorkflow[] }> {
  const res = await fetch(`${INGEST_URL}/workflows`)
  if (!res.ok) throw new Error(await ingestError(res, "Request failed"))
  return res.json()
}

export async function deleteWorkflow(name: string): Promise<void> {
  const res = await fetch(`${INGEST_URL}/workflows/${encodeURIComponent(name)}`, { method: "DELETE" })
  if (!res.ok) throw new Error(await ingestError(res, "Request failed"))
}

// The firm's skill library, owned by the ingest service: repo-shipped reference
// skills (read-only) plus the custom ones under WORKSPACE_ROOT/.skills. The
// response carries the library directory alongside the list — the skill builder
// chat scopes its engine sessions to that directory (there is no matter to scope
// to), so the page needs it from the same call.
export type Skill = { name: string; description: string; content: string; builtin: boolean; enabled: boolean }

export async function listSkills(): Promise<{ dir: string; skills: Skill[] }> {
  const res = await fetch(`${INGEST_URL}/skills`)
  if (!res.ok) throw new Error(await ingestError(res, "Request failed"))
  return res.json()
}

// Import an uploaded file as a skill: .md keeps its own frontmatter, .docx/.pdf/.txt
// are text-extracted raw for the skill builder chat to refine afterwards.
export async function importSkill(file: File): Promise<Skill> {
  const form = new FormData()
  form.append("file", file)
  const res = await fetch(`${INGEST_URL}/skills/import`, { method: "POST", body: form })
  if (!res.ok) throw new Error(await ingestError(res, "Request failed"))
  return res.json()
}

export async function deleteSkill(name: string): Promise<void> {
  const res = await fetch(`${INGEST_URL}/skills/${encodeURIComponent(name)}`, { method: "DELETE" })
  if (!res.ok) throw new Error(await ingestError(res, "Request failed"))
}

// Flip a skill on or off — builtin and custom alike. A disabled skill stays in
// the library but the engine hides it from every agent.
export async function setSkillEnabled(name: string, enabled: boolean): Promise<Skill> {
  const res = await fetch(`${INGEST_URL}/skills/${encodeURIComponent(name)}`, {
    method: "PATCH",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ enabled }),
  })
  if (!res.ok) throw new Error(await ingestError(res, "Request failed"))
  return res.json()
}

// The firm's specialist subagents, owned by the ingest service: repo-shipped
// specialists (read-only) plus the custom ones the agent builder composed.
// Custom agents are workflow building blocks — deletion is refused while a
// workflow references one.
export type AgentRecord = {
  name: string
  label: string
  description: string
  instructions: string
  builtin: boolean
  enabled: boolean
  created_at: number
}

export async function listAgents(): Promise<{ dir: string; agents: AgentRecord[] }> {
  const res = await fetch(`${INGEST_URL}/agents`)
  if (!res.ok) throw new Error(await ingestError(res, "Request failed"))
  return res.json()
}

export async function deleteAgent(name: string): Promise<void> {
  const res = await fetch(`${INGEST_URL}/agents/${encodeURIComponent(name)}`, { method: "DELETE" })
  if (!res.ok) throw new Error(await ingestError(res, "Request failed"))
}

// Flip an agent on or off — builtin and custom alike. A disabled agent stays in
// the library but the engine drops it from the task/workflow roster.
export async function setAgentEnabled(name: string, enabled: boolean): Promise<AgentRecord> {
  const res = await fetch(`${INGEST_URL}/agents/${encodeURIComponent(name)}`, {
    method: "PATCH",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ enabled }),
  })
  if (!res.ok) throw new Error(await ingestError(res, "Request failed"))
  return res.json()
}
