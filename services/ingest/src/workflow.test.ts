import { test, expect, afterAll } from "bun:test"
import { mkdtempSync, rmSync, existsSync, mkdirSync, writeFileSync, readFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"

// Set DOCHAUS_DIR before the dynamic import so workflow.ts captures the temp dir.
const dir = mkdtempSync(path.join(tmpdir(), "dochaus-workflow-"))
process.env.DOCHAUS_DIR = dir

// Seed hand-authored subagents: validate() only accepts step agents that
// listAgents() reports as enabled, and collision tests need a real existing .md.
const handAuthored = '---\ndescription: "Hand-authored reviewer"\nmode: subagent\n---\n\n# hand-authored'
mkdirSync(path.join(dir, "agent"), { recursive: true })
writeFileSync(path.join(dir, "agent", "legal-reviewer.md"), handAuthored)
writeFileSync(path.join(dir, "agent", "summarizer.md"), '---\ndescription: "Summarizes"\nmode: subagent\n---\n\nSummarize.')

const wf = await import("./workflow")

afterAll(() => rmSync(dir, { recursive: true, force: true }))

// --- renderAgentMarkdown ---

test("renderAgentMarkdown: frontmatter fields", () => {
  const md = wf.renderAgentMarkdown({
    name: "test-wf",
    label: "Test WF",
    description: "A test workflow",
    scope: "matter",
    prompt: "Run it",
    steps: [{ agent: "legal-reviewer", instructions: "" }],
    created_at: 0,
  })
  expect(md).toContain('description: "A test workflow"')
  expect(md).toContain("mode: primary")
  expect(md).toContain("temperature: 0.2")
  expect(md).toContain("color: primary")
  expect(md).toContain('"*": false')
  expect(md).toContain("read: true")
  expect(md).toContain("task: true")
  expect(md).toContain("search-document: true")
})

test("renderAgentMarkdown: falls back description when blank", () => {
  const md = wf.renderAgentMarkdown({
    name: "test-wf",
    label: "My Label",
    description: "",
    scope: "matter",
    prompt: "",
    steps: [{ agent: "legal-reviewer" }],
    created_at: 0,
  })
  expect(md).toContain('description: "Custom workflow: My Label"')
})

test("renderAgentMarkdown: pipeline step 1 seeds from matter", () => {
  const md = wf.renderAgentMarkdown({
    name: "test-wf",
    label: "Test WF",
    description: "desc",
    scope: "matter",
    prompt: "",
    steps: [{ agent: "legal-reviewer", instructions: "" }, { agent: "summarizer", instructions: "" }],
    created_at: 0,
  })
  expect(md).toContain("state the documents/matter and any focus from the user")
  expect(md).toContain("Pass the output of every prior step verbatim in the prompt")
})

test("renderAgentMarkdown: instructions inlined when present, omitted when blank", () => {
  const mdWith = wf.renderAgentMarkdown({
    name: "test-wf",
    label: "Test WF",
    description: "desc",
    scope: "matter",
    prompt: "",
    steps: [{ agent: "legal-reviewer", instructions: "Focus on indemnities." }],
    created_at: 0,
  })
  expect(mdWith).toContain("Focus on indemnities.")

  const mdWithout = wf.renderAgentMarkdown({
    name: "test-wf",
    label: "Test WF",
    description: "desc",
    scope: "matter",
    prompt: "",
    steps: [{ agent: "legal-reviewer", instructions: "" }],
    created_at: 0,
  })
  expect(mdWithout).not.toContain("Focus on indemnities.")
})

test("renderAgentMarkdown: output sections per step plus bottom line", () => {
  const md = wf.renderAgentMarkdown({
    name: "test-wf",
    label: "Test WF",
    description: "desc",
    scope: "matter",
    prompt: "",
    steps: [{ agent: "legal-reviewer" }, { agent: "summarizer" }],
    created_at: 0,
  })
  expect(md).toContain("**Legal Reviewer**")
  expect(md).toContain("**Summarizer**")
  expect(md).toContain("**Bottom line**")
  // The cross-pasted engineering mantra was deliberately removed from render output.
  expect(md).not.toContain("No edge case handling")
})

// --- create / update / delete round-trip ---

test("createWorkflow writes .md and registry entry", () => {
  const rec = wf.createWorkflow({
    label: "Diligence Sweep",
    description: "A diligence workflow",
    scope: "matter",
    prompt: 'Run the "Diligence sweep" workflow.',
    steps: [{ agent: "legal-reviewer", instructions: "Focus on indemnities." }],
  })
  expect(rec.name).toBe("diligence-sweep")
  expect(rec.label).toBe("Diligence Sweep")
  expect(rec.created_at).toBeGreaterThan(0)

  const mdPath = path.join(dir, "agent", "diligence-sweep.md")
  expect(existsSync(mdPath)).toBe(true)
  const md = readFileSync(mdPath, "utf8")
  expect(md).toContain("Diligence Sweep")

  const registry = wf.listWorkflows()
  expect(registry.some((r) => r.name === "diligence-sweep")).toBe(true)
})

test("updateWorkflow updates .md and registry, keeps name immutable", () => {
  wf.updateWorkflow("diligence-sweep", {
    label: "Diligence Sweep v2",
    description: "Updated",
    scope: "document",
    prompt: "New prompt",
    steps: [{ agent: "summarizer", instructions: "" }],
  })
  const registry = wf.listWorkflows()
  const rec = registry.find((r) => r.name === "diligence-sweep")!
  expect(rec.label).toBe("Diligence Sweep v2")
  expect(rec.scope).toBe("document")

  const md = readFileSync(path.join(dir, "agent", "diligence-sweep.md"), "utf8")
  expect(md).toContain("Diligence Sweep v2")
  expect(md).toContain("summarizer")
})

test("deleteWorkflow removes .md and registry entry", () => {
  wf.deleteWorkflow("diligence-sweep")
  expect(existsSync(path.join(dir, "agent", "diligence-sweep.md"))).toBe(false)
  expect(wf.listWorkflows().some((r) => r.name === "diligence-sweep")).toBe(false)
})

// --- collisions ---

test("collision: existing .md file (hand-authored legal-reviewer)", () => {
  expect(() =>
    wf.createWorkflow({
      label: "legal reviewer",
      description: "",
      scope: "matter",
      prompt: "",
      steps: [{ agent: "legal-reviewer" }],
    }),
  ).toThrow()
  // The hand-authored file must survive
  expect(existsSync(path.join(dir, "agent", "legal-reviewer.md"))).toBe(true)
  expect(readFileSync(path.join(dir, "agent", "legal-reviewer.md"), "utf8")).toBe(handAuthored)
})

test("collision: name already in registry", () => {
  wf.createWorkflow({ label: "My Flow", description: "", scope: "matter", prompt: "", steps: [{ agent: "summarizer" }] })
  expect(() =>
    wf.createWorkflow({ label: "My Flow", description: "", scope: "matter", prompt: "", steps: [{ agent: "summarizer" }] }),
  ).toThrow()
  // cleanup
  wf.deleteWorkflow("my-flow")
})

test("collision: reserved name", () => {
  expect(() =>
    wf.createWorkflow({ label: "build", description: "", scope: "matter", prompt: "", steps: [{ agent: "summarizer" }] }),
  ).toThrow()
})

// --- validation ---

test("empty steps → 400-mapped error", () => {
  let err: any
  try {
    wf.createWorkflow({ label: "empty steps", description: "", scope: "matter", prompt: "", steps: [] })
  } catch (e) {
    err = e
  }
  expect(err).toBeDefined()
  expect(err.status).toBe(400)
})

test("label that slugs to empty → 400-mapped error", () => {
  let err: any
  try {
    wf.createWorkflow({ label: "---", description: "", scope: "matter", prompt: "", steps: [{ agent: "summarizer" }] })
  } catch (e) {
    err = e
  }
  expect(err).toBeDefined()
  expect(err.status).toBe(400)
})

test("step agent with invalid name → 400-mapped error", () => {
  let err: any
  try {
    wf.createWorkflow({ label: "valid label", description: "", scope: "matter", prompt: "", steps: [{ agent: "Bad Agent!" }] })
  } catch (e) {
    err = e
  }
  expect(err).toBeDefined()
  expect(err.status).toBe(400)
})

// --- delete unknown ---

test("delete of name not in registry throws, hand-authored .md untouched", () => {
  expect(() => wf.deleteWorkflow("legal-reviewer")).toThrow()
  expect(existsSync(path.join(dir, "agent", "legal-reviewer.md"))).toBe(true)
  expect(readFileSync(path.join(dir, "agent", "legal-reviewer.md"), "utf8")).toBe(handAuthored)
})

test("updateWorkflow throws 404 for unknown name", () => {
  let err: any
  try {
    wf.updateWorkflow("does-not-exist", { label: "x", description: "", scope: "matter", prompt: "", steps: [{ agent: "summarizer" }] })
  } catch (e) {
    err = e
  }
  expect(err).toBeDefined()
  expect(err.status).toBe(404)
})
