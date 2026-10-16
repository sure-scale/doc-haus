import { readEngineConfig as awaitConfig } from "./engine-config"
import { test, expect, afterAll } from "bun:test"
import { mkdtempSync, rmSync, existsSync, mkdirSync, writeFileSync, readFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"

// Set both roots before the dynamic import so skill.ts captures the temp dirs.
const dochaus = mkdtempSync(path.join(tmpdir(), "dochaus-skill-config-"))
const workspace = mkdtempSync(path.join(tmpdir(), "dochaus-skill-workspace-"))
process.env.DOCHAUS_DIR = dochaus
process.env.DOCHAUS_RUNTIME_DIR = path.join(dochaus, "runtime")
process.env.WORKSPACE_ROOT = workspace

// Seed a repo-shipped (builtin) skill and a playbook that must stay excluded.
mkdirSync(path.join(dochaus, "skill", "clause-library"), { recursive: true })
writeFileSync(
  path.join(dochaus, "skill", "clause-library", "SKILL.md"),
  '---\nname: clause-library\ndescription: "Builtin clause refs"\n---\n\nBuiltin body',
)
mkdirSync(path.join(dochaus, "skill", "playbook-nda"), { recursive: true })
writeFileSync(
  path.join(dochaus, "skill", "playbook-nda", "SKILL.md"),
  '---\nname: playbook-nda\ndescription: "NDA playbook"\n---\n\nPlaybook body',
)

const sk = await import("./skill")

afterAll(() => {
  rmSync(dochaus, { recursive: true, force: true })
  rmSync(workspace, { recursive: true, force: true })
})

test("listSkills: builtins included, playbooks excluded", () => {
  const skills = sk.listSkills()
  const builtin = skills.find((s) => s.name === "clause-library")
  expect(builtin).toBeDefined()
  expect(builtin!.builtin).toBe(true)
  expect(builtin!.description).toBe("Builtin clause refs")
  expect(builtin!.content).toBe("Builtin body")
  expect(skills.some((s) => s.name === "playbook-nda")).toBe(false)
})

test("createSkill writes SKILL.md with quoted frontmatter", () => {
  const skill = sk.createSkill({
    name: "indemnity-standards",
    description: 'Firm: "balanced" indemnity positions',
    content: "## Indemnity\nMutual only.",
  })
  expect(skill.builtin).toBe(false)
  const file = path.join(workspace, ".skills", "indemnity-standards", "SKILL.md")
  expect(existsSync(file)).toBe(true)
  const raw = readFileSync(file, "utf8")
  expect(raw).toContain("name: indemnity-standards")
  expect(raw).toContain(`description: ${JSON.stringify('Firm: "balanced" indemnity positions')}`)
  expect(raw).toContain("## Indemnity")
  const listed = sk.listSkills().find((s) => s.name === "indemnity-standards")!
  expect(listed.description).toBe('Firm: "balanced" indemnity positions')
})

test("updateSkill rewrites description and body", () => {
  sk.updateSkill("indemnity-standards", { description: "Updated", content: "New body" })
  const listed = sk.listSkills().find((s) => s.name === "indemnity-standards")!
  expect(listed.description).toBe("Updated")
  expect(listed.content).toBe("New body")
})

test("deleteSkill removes the directory", () => {
  sk.deleteSkill("indemnity-standards")
  expect(existsSync(path.join(workspace, ".skills", "indemnity-standards"))).toBe(false)
  expect(sk.listSkills().some((s) => s.name === "indemnity-standards")).toBe(false)
})

test("createSkill: collision with builtin name → 409", () => {
  let err: any
  try {
    sk.createSkill({ name: "clause-library", description: "", content: "x" })
  } catch (e) {
    err = e
  }
  expect(err.status).toBe(409)
})

test("createSkill: playbook- prefix rejected", () => {
  let err: any
  try {
    sk.createSkill({ name: "playbook-msa", description: "", content: "x" })
  } catch (e) {
    err = e
  }
  expect(err.status).toBe(400)
})

test("createSkill: invalid name and empty content rejected", () => {
  expect(() => sk.createSkill({ name: "Bad Name!", description: "", content: "x" })).toThrow()
  expect(() => sk.createSkill({ name: "ok-name", description: "", content: "   " })).toThrow()
})

test("updateSkill: builtin → 403, unknown → 404, traversal rejected", () => {
  let builtinErr: any
  try {
    sk.updateSkill("clause-library", { description: "", content: "x" })
  } catch (e) {
    builtinErr = e
  }
  expect(builtinErr.status).toBe(403)

  let missingErr: any
  try {
    sk.updateSkill("nope", { description: "", content: "x" })
  } catch (e) {
    missingErr = e
  }
  expect(missingErr.status).toBe(404)

  expect(() => sk.updateSkill("../escape", { description: "", content: "x" })).toThrow()
})

test("deleteSkill: builtin → 403", () => {
  let err: any
  try {
    sk.deleteSkill("clause-library")
  } catch (e) {
    err = e
  }
  expect(err.status).toBe(403)
  expect(existsSync(path.join(dochaus, "skill", "clause-library", "SKILL.md"))).toBe(true)
})

test("setSkillEnabled: deny rule round-trips through opencode.json, bare-action form preserved", () => {
  // Mimic the real config, where `skill` is the bare action "allow".
  writeFileSync(path.join(dochaus, "opencode.json"), JSON.stringify({ permission: { skill: "allow" } }, null, 2))
  expect(sk.listSkills().find((s) => s.name === "clause-library")!.enabled).toBe(true)

  const off = sk.setSkillEnabled("clause-library", false)
  expect(off.enabled).toBe(false)
  expect(sk.listSkills().find((s) => s.name === "clause-library")!.enabled).toBe(false)
  const config = (awaitConfig())
  expect(config.permission!.skill).toEqual({ "*": "allow", "clause-library": "deny" })

  const on = sk.setSkillEnabled("clause-library", true)
  expect(on.enabled).toBe(true)
  const restored = (awaitConfig())
  expect(restored.permission!.skill).toEqual({ "*": "allow" })
})

test("setSkillEnabled: unknown skill → 404; deleteSkill drops the deny rule", () => {
  let err: any
  try {
    sk.setSkillEnabled("nope", false)
  } catch (e) {
    err = e
  }
  expect(err.status).toBe(404)

  sk.createSkill({ name: "toggle-me", description: "", content: "body" })
  sk.setSkillEnabled("toggle-me", false)
  sk.deleteSkill("toggle-me")
  const config = (awaitConfig())
  expect((config.permission!.skill as Record<string, unknown>)["toggle-me"]).toBeUndefined()
})

test("importSkill: .md keeps its frontmatter name and description", async () => {
  const md = '---\nname: deal-checklist\ndescription: "Closing checklist"\n---\n\n## Steps\nDo things.'
  const skill = await sk.importSkill("anything.md", Buffer.from(md))
  expect(skill.name).toBe("deal-checklist")
  expect(skill.description).toBe("Closing checklist")
  expect(skill.content).toContain("## Steps")
  sk.deleteSkill("deal-checklist")
})

test("importSkill: plain .md falls back to slugified filename", async () => {
  const skill = await sk.importSkill("Firm IP Positions.md", Buffer.from("Raw notes"))
  expect(skill.name).toBe("firm-ip-positions")
  expect(skill.content).toBe("Raw notes")
  sk.deleteSkill("firm-ip-positions")
})
