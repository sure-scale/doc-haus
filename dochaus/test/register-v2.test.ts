import { expect, test } from "bun:test"
import { childSessionID, normalizeToolMetadata, subagentReport } from "../tool/register-v2"

test("native legal tool metadata preserves citation data and omits absent optional fields", () => {
  const input = { citations: [{ document: "source.docx", section: "1", structuralPath: undefined, party: undefined, score: 0.5, superseded_by: undefined }], pending: [], title: "Sources", enabled: false, count: 0, absent: null }
  expect(normalizeToolMetadata(input)).toEqual({ citations: [{ document: "source.docx", section: "1", score: 0.5 }], pending: [], title: "Sources", enabled: false, count: 0, absent: null })
  expect(input.citations[0]).toHaveProperty("structuralPath")
})

test("invalid tool metadata cannot become a successful empty result", () => {
  const circular: Record<string, unknown> = {}
  circular.self = circular
  expect(() => normalizeToolMetadata(circular)).toThrow()
  expect(() => normalizeToolMetadata({ unsupported: 1n })).toThrow()
})

test("foreground task resolves the child session from native metadata or output", () => {
  expect(childSessionID({ metadata: { sessionID: "ses_child" } })).toBe("ses_child")
  expect(childSessionID({ output: { sessionID: "ses_output" } })).toBe("ses_output")
  expect(() => childSessionID({ output: { status: "running" } })).toThrow("session ID")
})

test("foreground task reports the final assistant text of the newest turn", () => {
  expect(subagentReport([
    { type: "idle", outcome: "succeeded" },
    { type: "assistant", content: [{ type: "text", text: "First finding" }, { type: "tool", text: "" }, { type: "text", text: "Second finding" }] },
    { type: "user" },
    { type: "idle", outcome: "succeeded" },
    { type: "assistant", content: [{ type: "text", text: "Previous turn" }] },
  ])).toBe("First finding\n\nSecond finding")
})

test("foreground task reports the placeholder when the child wrote no text", () => {
  expect(subagentReport([
    { type: "idle", outcome: "succeeded" },
    { type: "assistant", content: [{ type: "tool" }] },
  ])).toBe("Subagent completed without a text response.")
  // A follow-up that wrote no text must not report the previous turn's answer.
  expect(subagentReport([
    { type: "idle", outcome: "succeeded" },
    { type: "user" },
    { type: "idle", outcome: "succeeded" },
    { type: "assistant", content: [{ type: "text", text: "Previous turn" }] },
  ])).toBe("Subagent completed without a text response.")
})

test("foreground task turns a failed or interrupted child into an error", () => {
  expect(() => subagentReport([{ type: "idle", outcome: "failed", error: { message: "provider refused" } }])).toThrow("provider refused")
  expect(() => subagentReport([{ type: "idle", outcome: "interrupted" }])).toThrow("Subagent cancelled")
  expect(() => subagentReport([{ type: "idle", outcome: "failed", error: { message: "provider refused" } }], "ses_child")).toThrow("Subagent failed (sessionID: ses_child): provider refused")
  expect(() => subagentReport([{ type: "idle", outcome: "interrupted" }], "ses_child")).toThrow("Subagent cancelled (sessionID: ses_child)")
  // An earlier turn's idle marker does not confirm the current turn.
  expect(() => subagentReport([{ type: "user" }, { type: "idle", outcome: "succeeded" }])).toThrow("could not be confirmed")
})
