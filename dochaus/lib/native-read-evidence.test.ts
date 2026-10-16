import { expect, test } from "bun:test"
import { nativeReadEvidence } from "./native-read-evidence"

test("native complete and partial reads carry only verified range boundaries", () => {
  expect(nativeReadEvidence({ path: "handoff.md" }, "Read file /matter/handoff.md, lines 1-2\n1: First\n2: Second", { truncated: false })).toEqual({ input: { path: "handoff.md", filePath: "handoff.md", offset: 1 }, output: "Read file /matter/handoff.md, lines 1-2\n1: First\n2: Second\n(End of file - total 2 lines)", range: { start: 1, next: null } })
  expect(nativeReadEvidence({ path: "handoff.md", offset: 2 }, "Read file /matter/handoff.md, lines 2-3\n2: Second\n3: Third\n[Output truncated. Continue reading with offset: 4]", { truncated: true }).output).toEndWith("Use offset=4 to continue.)")
  for (const [output, metadata] of [
    ["Read directory /matter, entries 1-2\na\nb", { truncated: false }],
    ["Read file /matter/handoff.md, lines 1-3\n1: First\n3: Third", { truncated: false }],
    ["Read file /matter/handoff.md, lines 1-1\n1: First", {}],
    ["Read file /matter/handoff.md, lines 1-1\n1: Long... (line truncated to 2000 chars)", { truncated: false }],
  ] as const) expect(nativeReadEvidence({ path: "handoff.md" }, output, metadata).output).not.toContain("End of file")
})

test("native stored reader output projects to stable policy input and EOF evidence", async () => {
  const { projectSessionMessages } = await import("./engine-v2-projection")
  const { toModelContent } = await import("../../packages/core/src/tool/plugin/read")
  const content = toModelContent("/matter/handoff.md", 1, { type: "text-page", offset: 1, content: "First\nSecond", truncated: false })
  const messages = projectSessionMessages([{ id: "msg_read", type: "assistant", agent: "document-reviewer", model: { providerID: "fixture", id: "fixture" }, time: { created: 1, completed: 2 }, content: [{ type: "tool", id: "call_read", name: "read", time: { created: 1, completed: 2 }, state: { status: "completed", input: { path: "handoff.md", offset: 1 }, metadata: { truncated: false }, content: [{ type: "text", text: content }] } }] }], { id: "ses_read" })
  expect(messages[0]?.parts[0]?.state).toMatchObject({ input: { filePath: "handoff.md", path: "handoff.md", offset: 1 }, output: expect.stringContaining("(End of file - total 2 lines)") })
})

test("native read adaptation is idempotent for persisted EOF and continuation boundaries", () => {
  for (const truncated of [true, false]) {
    const output = "Read file /matter/handoff.md, lines 1-1\n1: First" + (truncated ? "\n[Output truncated. Continue reading with offset: 2]" : "")
    const first = nativeReadEvidence({ path: "handoff.md" }, output, { truncated })
    expect(nativeReadEvidence(first.input, first.output, { truncated })).toEqual(first)
  }
})
