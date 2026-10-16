import { expect, test } from "bun:test"
import { createElement } from "react"
import { renderToStaticMarkup } from "react-dom/server"
import { templateResult } from "../src/components/chat/template-result"
import { toTurns } from "../src/components/chat/turn-model"
import StepsPanel from "../src/components/chat/StepsPanel"
import type { Part } from "../src/api/session-view"

const placeholders = ["[insert client name]", "[insert city, region]", "[optional: renewal]"]
const output = `Added services.docx to the firm template library. Placeholders: ${placeholders.join(", ")}.`

function steps(tool: string, state: Extract<Part, { type: "tool" }>["state"]) {
  return toTurns([{ info: { id: "msg", role: "assistant" }, parts: [{
    type: "tool", id: "call", sessionID: "session", messageID: "msg", tool, state,
  }] }], "")[0]!.steps
}

test("empty template lookup is a status without an empty expandable output", () => {
  const result = steps("list-templates", { status: "completed", input: {}, title: "0 template(s)", output: "[]" })
  expect(result[0]).toMatchObject({ label: "No saved templates found", status: "done", detail: undefined })
  const html = renderToStaticMarkup(createElement(StepsPanel, { steps: result, busy: false, answered: true }))
  expect(html).not.toContain('class="step-detail"')
  expect(html).not.toContain("[]")
})

test("lookup errors and nonempty results keep their output", () => {
  expect(steps("list-templates", { status: "error", input: {}, title: "Lookup failed", error: "[]" })[0])
    .toMatchObject({ status: "error", detail: "[]" })
  expect(steps("list-templates", { status: "completed", input: {}, title: "1 template(s)", output: '[{"template":"services.docx"}]' })[0])
    .toMatchObject({ detail: '[{"template":"services.docx"}]' })
})

test("creation shows a summary and collapsed fields without splitting commas inside placeholders", () => {
  const result = steps("create-template", { status: "completed", input: {}, title: "Created template services.docx", output, metadata: { placeholders } })
  expect(result[0]).toMatchObject({ templateResult: { summary: "Added services.docx to the firm template library.", placeholders } })
  const html = renderToStaticMarkup(createElement(StepsPanel, { steps: result, busy: false, answered: true }))
  expect(html).toContain('<details class="step-template-fields">')
  expect(html).toContain("3 placeholders")
  expect(html).toContain("<li>[insert city, region]</li>")
  expect(html).not.toContain('<pre class="step-output">')
})

test("updates use the same presentation and unrecognized or warning-bearing outputs remain intact", () => {
  expect(templateResult("update-template", output, { placeholders })?.placeholders).toEqual(placeholders)
  expect(templateResult("create-template", output + " Warning: needs review.", { placeholders })).toBeUndefined()
  expect(templateResult("create-template", output, { placeholders: [1] })).toBeUndefined()
  expect(templateResult("create-template", output)).toBeUndefined()
  expect(templateResult("other-tool", output, { placeholders })).toBeUndefined()
  const warning = "It exposes no placeholders — confirm every variable term is a [insert ...] placeholder."
  expect(templateResult("create-template", warning, { placeholders: [] })).toEqual({ summary: warning, placeholders: [] })
})
