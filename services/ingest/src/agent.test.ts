import { readEngineConfig as awaitConfig } from "./engine-config"
import { test, expect, afterAll } from "bun:test"
import { mkdtempSync, rmSync, existsSync, mkdirSync, writeFileSync, readFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"

// Set both roots before the dynamic import so agent.ts (and the workflow registry
// it consults for delete protection) capture the temp dirs.
const dochaus = mkdtempSync(path.join(tmpdir(), "dochaus-agent-config-"))
const workspace = mkdtempSync(path.join(tmpdir(), "dochaus-agent-workspace-"))
process.env.DOCHAUS_DIR = dochaus
process.env.DOCHAUS_RUNTIME_DIR = path.join(dochaus, "runtime")
process.env.WORKSPACE_ROOT = workspace

// Seed a repo-shipped specialist (mode: subagent), a primary agent that must stay
// out of the list, and a hidden subagent.
mkdirSync(path.join(dochaus, "agent"), { recursive: true })
writeFileSync(
  path.join(dochaus, "agent", "legal-reviewer.md"),
  '---\ndescription: "Reviews documents"\nmode: subagent\n---\n\nReview things.',
)
writeFileSync(path.join(dochaus, "agent", "qa.md"), '---\ndescription: "Q&A"\nmode: primary\n---\n\nAnswer.')
writeFileSync(
  path.join(dochaus, "agent", "secret.md"),
  '---\ndescription: "internal"\nmode: subagent\nhidden: true\n---\n\nShh.',
)

const ag = await import("./agent")
const wf = await import("./workflow")

afterAll(() => {
  rmSync(dochaus, { recursive: true, force: true })
  rmSync(workspace, { recursive: true, force: true })
})

test("listAgents: builtin subagents only, primaries and hidden excluded", () => {
  const agents = ag.listAgents()
  const builtin = agents.find((a) => a.name === "legal-reviewer")
  expect(builtin).toBeDefined()
  expect(builtin!.builtin).toBe(true)
  expect(builtin!.label).toBe("Legal Reviewer")
  expect(builtin!.instructions).toBe("Review things.")
  expect(agents.some((a) => a.name === "qa")).toBe(false)
  expect(agents.some((a) => a.name === "secret")).toBe(false)
})

test("setAgentEnabled: disable flag round-trips through opencode.json, other agent entries kept", () => {
  // Mimic the real config, where stock engine agents are already disabled.
  writeFileSync(
    path.join(dochaus, "opencode.json"),
    JSON.stringify({ agent: { explore: { disable: true } } }, null, 2),
  )
  expect(ag.listAgents().find((a) => a.name === "legal-reviewer")!.enabled).toBe(true)

  const off = ag.setAgentEnabled("legal-reviewer", false)
  expect(off.enabled).toBe(false)
  expect(ag.listAgents().find((a) => a.name === "legal-reviewer")!.enabled).toBe(false)
  const config = (awaitConfig())
  expect(config.agent).toEqual({ explore: { disable: true }, "legal-reviewer": { disable: true } })

  const on = ag.setAgentEnabled("legal-reviewer", true)
  expect(on.enabled).toBe(true)
  const restored = (awaitConfig())
  expect(restored.agent).toEqual({ explore: { disable: true } })
})

test("setAgentEnabled: unknown agent → 404", () => {
  let err: any
  try {
    ag.setAgentEnabled("nope", false)
  } catch (e) {
    err = e
  }
  expect(err.status).toBe(404)
})

test("renderCustomAgentMarkdown: frontmatter, task body, citation discipline", () => {
  const md = ag.renderCustomAgentMarkdown({
    name: "ip-specialist",
    label: "IP Specialist",
    description: "Checks IP ownership chains",
    instructions: "Trace every IP assignment.",
    created_at: 0,
  })
  expect(md).toContain('description: "Checks IP ownership chains"')
  expect(md).toContain("mode: subagent")
  expect(md).toContain('"*": false')
  expect(md).toContain("search-document: true")
  expect(md).toContain("cite: true")
  expect(md).toContain("<task>\nTrace every IP assignment.\n</task>")
  expect(md).toContain("Anchor every quoted excerpt with the `cite` tool")
  // The cross-pasted engineering mantra was deliberately removed from render output.
  expect(md).not.toContain("No edge case handling")
})

test("createAgent writes .md and registry; appears in listAgents once", () => {
  const rec = ag.createAgent({
    label: "IP Specialist",
    description: "Checks IP ownership chains",
    instructions: "Trace every IP assignment.",
  })
  expect(rec.name).toBe("ip-specialist")
  expect(existsSync(path.join(dochaus, "agent", "ip-specialist.md"))).toBe(true)
  const listed = ag.listAgents().filter((a) => a.name === "ip-specialist")
  expect(listed.length).toBe(1)
  expect(listed[0]!.builtin).toBe(false)
})

test("updateAgent rewrites .md and registry, name immutable", () => {
  ag.updateAgent("ip-specialist", {
    label: "IP Specialist v2",
    description: "Updated",
    instructions: "Also check trademarks.",
  })
  const rec = ag.listAgents().find((a) => a.name === "ip-specialist")!
  expect(rec.label).toBe("IP Specialist v2")
  expect(readFileSync(path.join(dochaus, "agent", "ip-specialist.md"), "utf8")).toContain("Also check trademarks.")
})

test("deleteAgent refused while a workflow references it", () => {
  wf.createWorkflow({
    label: "IP Sweep",
    description: "",
    scope: "matter",
    prompt: "",
    steps: [{ agent: "ip-specialist", instructions: "" }],
  })
  let err: any
  try {
    ag.deleteAgent("ip-specialist")
  } catch (e) {
    err = e
  }
  expect(err.status).toBe(409)
  expect(err.message).toContain("ip-sweep")
  expect(existsSync(path.join(dochaus, "agent", "ip-specialist.md"))).toBe(true)

  wf.deleteWorkflow("ip-sweep")
  ag.deleteAgent("ip-specialist")
  expect(existsSync(path.join(dochaus, "agent", "ip-specialist.md"))).toBe(false)
  expect(ag.listAgents().some((a) => a.name === "ip-specialist")).toBe(false)
})

test("collisions: builtin file, reserved name, duplicate registry entry", () => {
  expect(() => ag.createAgent({ label: "Legal Reviewer", description: "", instructions: "x" })).toThrow()
  expect(() => ag.createAgent({ label: "build", description: "", instructions: "x" })).toThrow()
  ag.createAgent({ label: "My Agent", description: "", instructions: "x" })
  expect(() => ag.createAgent({ label: "My Agent", description: "", instructions: "x" })).toThrow()
  ag.deleteAgent("my-agent")
})

test("validation: bad label and empty instructions → 400", () => {
  let err: any
  try {
    ag.createAgent({ label: "---", description: "", instructions: "x" })
  } catch (e) {
    err = e
  }
  expect(err.status).toBe(400)
  try {
    ag.createAgent({ label: "Fine Label", description: "", instructions: "  " })
  } catch (e) {
    err = e
  }
  expect(err.status).toBe(400)
})

test("update/delete unknown agent → 404", () => {
  let err: any
  try {
    ag.updateAgent("nope", { label: "x", description: "", instructions: "y" })
  } catch (e) {
    err = e
  }
  expect(err.status).toBe(404)
  try {
    ag.deleteAgent("legal-reviewer")
  } catch (e) {
    err = e
  }
  expect(err.status).toBe(404)
  expect(existsSync(path.join(dochaus, "agent", "legal-reviewer.md"))).toBe(true)
})
