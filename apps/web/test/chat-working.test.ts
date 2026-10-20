import { expect, test } from "bun:test"
import completed from "../../../dochaus/workflows/lean/fixtures/status-completed.json"
import { isLeanDraftingStatus } from "../src/api/ingest/lean"
import type { Part } from "../src/api/session-view"
import { readTurn } from "../src/components/chat/turn-model"
import { statusbarLabel } from "../src/components/matter/MatterRail"

// A chat answer shows it is working, and what it is doing when that has a plain
// name, in the live answer and the matter's status bar.

function tool(id: string, name: string, status: "pending" | "running" | "completed", input: Record<string, unknown>): Part {
  const state =
    status === "completed"
      ? { status, input, output: "FOUND", title: name }
      : status === "running"
        ? { status, input, title: name }
        : { status, input }
  return { type: "tool", id, sessionID: "ses", messageID: "msg", tool: name, state }
}

function activity(...parts: Part[]) {
  return readTurn(new Map(parts.map((part) => [part.id, part])), new Map([["msg", "assistant"]]), "matter", "ses").activity
}

test("a live answer names its running step in plain words, and nothing between steps", () => {
  expect(activity(tool("p1", "search-document", "running", { query: "letter date" }))).toBe("Searching matter documents")
  expect(activity(tool("p1", "read", "pending", { filePath: "/m/Invoice 2.docx" }))).toBe("Reading source material")
  expect(activity(tool("p1", "verify-quote", "running", { quote: "Dated 12 October" }))).toBe("Checking quotations")
  // The latest running step is the one named.
  expect(
    activity(tool("p1", "read", "running", { filePath: "/m/Letter.docx" }), tool("p2", "search-document", "running", { query: "due" })),
  ).toBe("Searching matter documents")
  // Between steps every tool has finished; a tool without a plain name is not named.
  expect(activity(tool("p1", "search-document", "completed", { query: "letter date" }))).toBeUndefined()
  expect(activity(tool("p1", "get-section", "running", { section: "4" }))).toBeUndefined()
  expect(activity({ type: "text", id: "t1", sessionID: "ses", messageID: "msg", text: "The letter is dated 12 October." })).toBeUndefined()
})

const done = completed as unknown
if (!isLeanDraftingStatus(done)) throw new Error("fixture is not a lean drafting status")

test("the status bar shows a chat answer while it works, and the drafting label or Ready otherwise", () => {
  expect(statusbarLabel()).toEqual({ text: "Ready", findings: false })
  expect(statusbarLabel(undefined, {})).toEqual({ text: "Answer in progress", status: "running", findings: false })
  expect(statusbarLabel(undefined, { activity: "Searching matter documents" }).text).toBe("Answer in progress · Searching matter documents")
  // Drafting keeps its label when no answer runs, and yields only while one does.
  expect(statusbarLabel(done)).toEqual({ text: "Drafting complete", status: "completed", findings: false })
  expect(statusbarLabel(done, {})).toEqual({ text: "Answer in progress", status: "running", findings: false })
})
