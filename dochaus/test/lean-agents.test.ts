import { expect, test } from "bun:test"
import { readdirSync, readFileSync } from "node:fs"
import path from "node:path"
import { ConfigMarkdown } from "@opencode/core/config/markdown"
import { ConfigAgentV1 } from "@opencode/core/v1/config/agent"
import { ConfigMigrateV1 } from "@opencode/core/v1/config/migrate"
import { Permission } from "@opencode/core/permission"
import { LEAN_CHILD_PERMISSIONS, LEAN_READ_TOOLS, LEAN_STAGE_AGENTS } from "../workflows/lean/stages"

// The lean stage agents load through the same native markdown + v1 migration
// path the engine uses, and their own frontmatter (not only the session rules)
// allows exactly the read tools listed by the coordinator.

const core = path.resolve(path.dirname(Bun.resolveSync("@opencode/core/session", import.meta.dir)), "..")
const { Schema } = await import(Bun.resolveSync("effect", core))
const root = path.resolve(import.meta.dir, "..")
const agents = Object.values(LEAN_STAGE_AGENTS).filter(id => id !== "lean-research")

function loadAgent(id: string) {
  const markdown = ConfigMarkdown.parse(readFileSync(path.join(root, "agent", `${id}.md`), "utf8"))
  return ConfigMigrateV1.migrateAgent(Schema.decodeUnknownSync(ConfigAgentV1.Info)({ ...markdown.data, prompt: markdown.content },
    { errors: "all", propertyOrder: "original" }))
}

// Every doc.haus tool plus the native built-ins a legacy agent could reach.
const TOOLS = [
  ...readdirSync(path.join(root, "tool")).filter(name => name.endsWith(".ts") && !name.endsWith(".test.ts")).map(name => name.slice(0, -3)),
  "edit", "write", "bash", "task", "skill", "glob", "grep", "list", "patch", "webfetch", "todowrite", "todoread",
]

test("each lean stage agent loads as a hidden subagent whose own rules allow only the coordinator's read tools", () => {
  expect(agents.toSorted()).toEqual(["lean-change-plan", "lean-change-prepare", "lean-change-review", "lean-changes", "lean-coverage", "lean-discrepancy", "lean-draft", "lean-drafting-notes", "lean-files", "lean-intake", "lean-package-review", "lean-research-review", "lean-review", "lean-review-client-disadvantage", "lean-review-independent", "lean-review-playbook", "lean-review-wording", "lean-spine"])
  for (const id of agents) {
    const agent = loadAgent(id)
    expect(agent.mode).toBe("subagent")
    expect(agent.hidden).toBe(true)
    expect(agent.system?.trim()).toStartWith("You ")
    const allowed = [...new Set([...TOOLS, ...LEAN_READ_TOOLS])]
      .filter(action => Permission.evaluate(action, "*", agent.permissions ?? []).effect === "allow")
    expect({ id, allowed: allowed.toSorted() }).toEqual({ id, allowed: [...LEAN_READ_TOOLS].toSorted() })
    // The session rules the coordinator adds restate the same boundary.
    const session = [...(agent.permissions ?? []), ...LEAN_CHILD_PERMISSIONS]
    expect(TOOLS.filter(action => Permission.evaluate(action, "*", session).effect === "allow").toSorted())
      .toEqual(LEAN_READ_TOOLS.filter(tool => TOOLS.includes(tool)).toSorted())
  }
})

// The drafting narrator restates run progress for the lawyer; it reads only the
// excerpt it is sent and may call no tool at all.
test("narrator agent loads as a hidden primary agent with every tool denied", () => {
  const agent = loadAgent("narrator")
  expect(agent.mode).toBe("primary")
  expect(agent.hidden).toBe(true)
  expect(agent.system?.trim()).toStartWith("You ")
  expect([...TOOLS, ...LEAN_READ_TOOLS].filter(action => Permission.evaluate(action, "*", agent.permissions ?? []).effect === "allow"))
    .toEqual([])
})

// The research agent loads through the same path; its grant set includes direct
// official-source retrieval, which the research settings describe as available.
test("research agent's own rules allow exactly its matter, case-law and official-source tools", () => {
  const agent = loadAgent("research")
  expect(TOOLS.filter(action => Permission.evaluate(action, "*", agent.permissions ?? []).effect === "allow").toSorted())
    .toEqual(["case-law", "cite", "glob", "grep", "list", "official-source", "search-document"])
})


test("lean research alone permits discovery and authority retrieval without writing or delegating", () => {
  const agent = loadAgent("lean-research")
  expect(agent.mode).toBe("subagent")
  expect(agent.hidden).toBe(true)
  expect(TOOLS.filter(action => Permission.evaluate(action, "*", agent.permissions ?? []).effect === "allow").toSorted())
    .toEqual(["case-law", "official-source", "read-document", "read-record", "search-document", "verify-quote", "web-search", "webfetch"])
})
