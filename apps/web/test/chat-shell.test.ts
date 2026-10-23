import { expect, test } from "bun:test"
import { createElement, createRef } from "react"
import { renderToStaticMarkup } from "react-dom/server"
import ChatEmpty from "../src/components/chat/ChatEmpty"
import Composer from "../src/components/chat/Composer"

const starters = ["What are the key points in these materials?"]

function empty(documentCount?: number) {
  return renderToStaticMarkup(
    createElement(ChatEmpty, { hint: "Ask about the documents.", starters, documentCount, onPick: () => {}, onAddDocuments: () => {} }),
  )
}

test("a matter with no documents leads with adding them instead of document starters", () => {
  const html = empty(0)
  expect(html).toContain("Start with the source documents")
  expect(html).toContain("Add documents")
  expect(html).toContain("Drop Word, PDF, email, Excel or zip files onto the message box below")
  expect(html).not.toContain(starters[0])
})

test("starters return once the matter has documents, and surfaces without a count keep them", () => {
  for (const html of [empty(2), empty(undefined)]) {
    expect(html).toContain(starters[0])
    expect(html).toContain("Ask about the documents.")
    expect(html).not.toContain("Add documents")
  }
})

function composer(toolsLoading: boolean, draftingClarification = false) {
  return renderToStaticMarkup(
    createElement(Composer, {
      available: new Set(["draft", "extract"]),
      agent: "auto",
      setResolvedAgent: () => {},
      onAgentChange: () => {},
      inputRef: createRef<HTMLTextAreaElement>(),
      workflows: [],
      runWorkflow: async () => {},
      playbooks: [{ name: "standard", description: "Standard positions" }],
      onPlaybookChange: () => {},
      toolsLoading,
      draftingUpdate: draftingClarification,
      draftingClarification,
      composerPlaceholder: "Ask",
      input: "",
      setInput: () => {},
      onSend: async () => {},
      busy: false,
    }),
  )
}

test("the composer holds a place for each control while its choices load", () => {
  const loading = composer(true)
  expect(loading.match(/composer-tool-skeleton/g)?.length).toBe(3)
  expect(loading).not.toContain("Playbook")
  const loaded = composer(false)
  expect(loaded).not.toContain("composer-tool-skeleton")
  expect(loaded).toContain("Playbook")
})


test("held checklist composer names clarification without claiming a draft edit", () => {
  const html = composer(false, true)
  expect(html).toContain('aria-label="Send checklist clarification"')
  expect(html).toContain("Clarify checklist")
  expect(html).not.toContain("Update draft")
})
