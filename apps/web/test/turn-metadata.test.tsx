import { expect, test } from "bun:test"
import { createElement, Fragment } from "react"
import { renderToStaticMarkup } from "react-dom/server"
import type { Part } from "../src/api/session-view"
import { readTurn, toTurns } from "../src/components/chat/turn-model"
import CitationView from "../src/components/CitationView"
import TodoView from "../src/components/TodoView"
import RedlineView from "../src/components/chat/RedlineView"

const text = 'Literal [source: example] ["array"] \"quote\" — 😀'
const citation = { documentName: "terms.docx", docPath: "terms.docx", section: "4", excerpt: text, charStart: 0, charEnd: text.length }
function part(tool: string, metadata: unknown, input: unknown = {}): Part {
  return { type: "tool", id: tool, sessionID: "ses", messageID: "msg", tool,
    state: { status: "completed", input, metadata, output: "Original tool output remains.", title: tool } } as Part
}
function views(parts: Part[]) {
  const live = readTurn(new Map(parts.map(p => [p.id, p])), new Map([["msg", "assistant"]]), "", "ses")
  const stored = toTurns([{ info: { id: "msg", role: "assistant" }, parts: JSON.parse(JSON.stringify(parts)) }], "")[0]!
  return [live, stored]
}
test("malformed tool metadata cannot become citation, todo or redline cards", () => {
  const parts = [part("cite", { citations: [null, { ...citation, excerpt: {} }, { ...citation, charStart: -1 }] }),
    part("todowrite", { todos: [{ content: {}, status: "completed", priority: "high" }] }, { todos: "completed" }),
    part("redline", { document: { name: "terms.docx" }, oldText: text, replacement: text })]
  for (const view of views(parts)) {
    expect(view.citations).toEqual([])
    expect(view.todos).toEqual([])
    expect(view.redlines).toEqual([])
    expect(() => renderToStaticMarkup(createElement(Fragment, null,
      createElement(CitationView, { citations: view.citations }), createElement(TodoView, { todos: view.todos }),
      createElement(RedlineView, { redlines: view.redlines, onView: () => {} })))).not.toThrow()
    expect(view.steps.map(s => s.kind === "tool" ? s.detail : "").join(" ")).toContain("Original tool output remains.")
  }
})
test("valid tool metadata preserves literal text identically live and after reload", () => {
  const parts = [part("cite", { citations: [citation] }),
    part("todowrite", { todos: [{ content: text, status: "pending", priority: "high" }] }),
    part("redline", { document: "terms.docx", oldText: text, replacement: text, redline: 4, comment: text })]
  const [live, stored] = views(parts)
  expect(live!.citations).toEqual([citation]); expect(stored!.citations).toEqual(live!.citations)
  expect(live!.todos[0]?.content).toBe(text); expect(stored!.todos).toEqual(live!.todos)
  expect(live!.redlines[0]?.newText).toBe(text); expect(stored!.redlines).toEqual(live!.redlines)
})
test("one bad metadata container cannot fabricate array items", () => {
  for (const view of views([part("cite", { citations: citation }), part("todowrite", { todos: {} })])) {
    expect(view.citations).toEqual([]); expect(view.todos).toEqual([])
  }
})

test("malformed todo metadata keeps that tool's own output when no card can render", () => {
  const value = part("todowrite", { todos: [{ content: {} }] })
  for (const view of views([value])) {
    expect(view.todos).toEqual([])
    expect(view.steps).toHaveLength(1)
    expect(view.steps[0]?.kind === "tool" && view.steps[0].detail).toBe("Original tool output remains.")
  }
})
test("comparison metadata cannot change options or disagree after JSON reload", () => {
  const base = { left: "old.docx", right: "new.docx", counts: { changes: 1, untracked: 0 }, options: { ignoreFormatting: false } }
  for (const compare of [{ ...base, counts: { changes: NaN, untracked: 0 } },
    { ...base, counts: { changes: -1, untracked: 0 } }, { ...base, options: { ignoreFormatting: "false" } },
    { ...base, options: [] }]) {
    for (const view of views([part("compare-versions", { compare })])) expect(view.compares).toEqual([])
  }
  const [live, stored] = views([part("compare-versions", { compare: base })])
  expect(live!.compares?.[0]?.ignoreFormatting).toBe(false)
  expect(stored!.compares).toEqual(live!.compares)
})
