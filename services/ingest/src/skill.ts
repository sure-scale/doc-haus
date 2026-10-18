import { existsSync, mkdirSync, readdirSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import path from "node:path"
import { workspaceRoot } from "./matter"
import { DOCHAUS_DIR } from "./workflow"
import { extractDocumentText } from "./ingest"
import { disabledSkills, setSkillDisabled } from "./engine-config"

// The firm's custom skill library: reference knowledge (clause standards,
// checklists, drafting guidance) the engine's agents load on demand. Custom
// skills live in WORKSPACE_ROOT/.skills/<name>/SKILL.md — a skills path in
// dochaus/opencode.json — and ingest is their single writer, mirroring the
// .playbooks library. Repo-shipped skills under dochaus/skill/ are listed
// read-only beside them. Playbooks (playbook-*) are skills too but have their
// own library and lifecycle, so they are excluded here.
// Read lazily (workspaceRoot) so test files can point WORKSPACE_ROOT at their
// own temp dirs regardless of which test file imported this module first.
export const SKILLS_DIR = () => path.join(workspaceRoot(), ".skills")

const BUILTIN_SKILL_DIR = () => path.join(DOCHAUS_DIR(), "skill")

const NAME_RE = /^[a-z0-9][a-z0-9-]*$/

export type Skill = {
  name: string
  description: string
  content: string
  builtin: boolean
  enabled: boolean
}

class SkillError extends Error {
  status: number
  constructor(message: string, status: number) {
    super(message)
    this.status = status
  }
}

function skillFile(name: string) {
  return path.join(SKILLS_DIR(), name, "SKILL.md")
}

// SKILL.md = frontmatter (name + description) + markdown body. The frontmatter
// is owned here — description is JSON-quoted so a colon or newline cannot
// inject YAML keys; callers supply only the body.
function renderSkillMarkdown(name: string, description: string, content: string) {
  return `---\nname: ${name}\ndescription: ${JSON.stringify(description)}\n---\n\n${content.trim()}\n`
}

function parseSkillMarkdown(source: string) {
  const lines = source.split("\n")
  if (lines[0]?.trim() !== "---") return { name: "", description: "", content: source.trim() }
  const end = lines.findIndex((l, i) => i > 0 && l.trim() === "---")
  if (end === -1) return { name: "", description: "", content: source.trim() }
  const read = (key: string) => {
    const line = lines.slice(1, end).find((l) => l.startsWith(`${key}:`))
    if (!line) return ""
    const raw = line.slice(key.length + 1).trim()
    if (raw.startsWith('"')) return JSON.parse(raw) as string
    return raw
  }
  return { name: read("name"), description: read("description"), content: lines.slice(end + 1).join("\n").trim() }
}

function readSkillDir(dir: string, builtin: boolean) {
  if (!existsSync(dir)) return []
  return readdirSync(dir)
    .map((entry) => path.join(dir, entry, "SKILL.md"))
    .filter((file) => existsSync(file))
    .map((file) => {
      const parsed = parseSkillMarkdown(readFileSync(file, "utf8"))
      const name = parsed.name || path.basename(path.dirname(file))
      return { name, description: parsed.description, content: parsed.content, builtin }
    })
    .filter((s) => !s.name.startsWith("playbook-"))
}

export function listSkills(): Skill[] {
  const disabled = disabledSkills()
  return [...readSkillDir(BUILTIN_SKILL_DIR(), true), ...readSkillDir(SKILLS_DIR(), false)]
    .map((s) => ({ ...s, enabled: !disabled.has(s.name) }))
    .sort((a, b) => a.name.localeCompare(b.name))
}

// Flip a skill on or off — builtin and custom alike. Off means the engine
// permission-denies it, so it leaves every agent's available list; the file
// stays on disk and the library keeps showing it.
export function setSkillEnabled(name: string, enabled: boolean): Skill {
  const skill = listSkills().find((s) => s.name === name)
  if (!skill) throw new SkillError(`skill "${name}" not found`, 404)
  setSkillDisabled(name, !enabled)
  return { ...skill, enabled }
}

function validateName(name: string) {
  if (!NAME_RE.test(name)) throw new SkillError(`"${name}" is not a valid skill name (lowercase, hyphenated)`, 400)
  if (name.startsWith("playbook-")) throw new SkillError('Playbooks are managed separately — run the import-playbook command (or ask the assistant to import a playbook) for "playbook-" names', 400)
}

export function createSkill(input: { name: string; description: string; content: string }): Skill {
  validateName(input.name)
  if (!input.content.trim()) throw new SkillError("content must be non-empty", 400)
  if (listSkills().some((s) => s.name === input.name)) throw new SkillError(`skill "${input.name}" already exists`, 409)
  mkdirSync(path.join(SKILLS_DIR(), input.name), { recursive: true })
  writeFileSync(skillFile(input.name), renderSkillMarkdown(input.name, input.description, input.content))
  return { name: input.name, description: input.description, content: input.content.trim(), builtin: false, enabled: true }
}

export function updateSkill(name: string, input: { description: string; content: string }): Skill {
  validateName(name)
  if (!existsSync(skillFile(name))) {
    if (listSkills().some((s) => s.name === name && s.builtin)) throw new SkillError(`skill "${name}" is built-in and read-only`, 403)
    throw new SkillError(`skill "${name}" not found`, 404)
  }
  if (!input.content.trim()) throw new SkillError("content must be non-empty", 400)
  writeFileSync(skillFile(name), renderSkillMarkdown(name, input.description, input.content))
  return { name, description: input.description, content: input.content.trim(), builtin: false, enabled: !disabledSkills().has(name) }
}

export function deleteSkill(name: string) {
  validateName(name)
  if (!existsSync(skillFile(name))) {
    if (listSkills().some((s) => s.name === name && s.builtin)) throw new SkillError(`skill "${name}" is built-in and cannot be deleted`, 403)
    throw new SkillError(`skill "${name}" not found`, 404)
  }
  rmSync(path.join(SKILLS_DIR(), name), { recursive: true, force: true })
  // Drop any deny rule so a later skill reusing the name starts enabled.
  setSkillDisabled(name, false)
}

// Import an uploaded file as a skill. A .md keeps its own frontmatter name and
// description when present; .docx/.pdf/.txt are text-extracted and imported raw,
// for the skill builder chat to refine afterwards.
export async function importSkill(fileName: string, buffer: Buffer): Promise<Skill> {
  const base = path.basename(fileName)
  const ext = path.extname(base).toLowerCase()
  const fallbackName = base
    .slice(0, base.length - ext.length)
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-|-$/g, "")
  if (ext === ".md" || ext === ".txt") {
    const parsed = parseSkillMarkdown(buffer.toString("utf8"))
    return createSkill({
      name: parsed.name || fallbackName,
      description: parsed.description,
      content: parsed.content,
    })
  }
  const text = await extractDocumentText(base, buffer)
  return createSkill({ name: fallbackName, description: "", content: text })
}
