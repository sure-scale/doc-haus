import { describe, expect, test } from "bun:test"
import { createElement } from "react"
import { renderToStaticMarkup } from "react-dom/server"
import type { Part } from "../src/api/session-view"
import Markdown from "../src/components/markdown/Markdown"
import StepsPanel from "../src/components/chat/StepsPanel"
import { toTurns } from "../src/components/chat/turn-model"

const answer = (text: string) => renderToStaticMarkup(createElement(Markdown, { children: text }))
function tool(tool: string, output: string) {
  const part: Part = {
    type: "tool",
    id: "call",
    sessionID: "session",
    messageID: "message",
    tool,
    state: { status: "completed", input: {}, title: tool, output },
  }
  const turns = toTurns([{ info: { id: "message", role: "assistant" }, parts: [part] }], "")
  return renderToStaticMarkup(createElement(StepsPanel, { steps: turns[0]!.steps, busy: false, answered: true }))
}

describe("lossless legacy report rendering", () => {
  test("styles ordinary reports while keeping surrounding qualifications in order", () => {
    const control = answer("Before\n<findings>Thirty days.</findings>")
    expect(control).toContain('class="report-section-title">Findings')
    expect(control).toContain("Thirty days.")
    const html = answer(
      "Before\n<findings>Thirty days.</findings>\nCheck the missing exhibit.\n<notes>A qualification.</notes>\nDo not rely on this yet.",
    )
    expect(html).toContain("Check the missing exhibit.")
    expect(html).toContain("Do not rely on this yet.")
    expect(html.indexOf("Thirty days.")).toBeLessThan(html.indexOf("Check the missing exhibit."))
    expect(html.indexOf("Check the missing exhibit.")).toBeLessThan(html.indexOf("A qualification."))
    expect(html.indexOf("A qualification.")).toBeLessThan(html.indexOf("Do not rely on this yet."))
  })

  test("keeps substantive leading warnings visible even when they mention a gate", () => {
    const warning =
      "Do not rely on this report until the source access gate is resolved. The original exhibit is missing, and the following dates and conclusions need checking against it."
    const html = answer(`${warning}\n<findings>Thirty days.</findings>`)
    expect(html).toContain(warning)
    expect(html).not.toContain('class="report-narration"')
  })

  test("keeps text inside result wrappers in order around nested sections", () => {
    const html = answer("<result>First.<findings>Second.</findings>Third.</result>")
    expect(html.indexOf("First.")).toBeLessThan(html.indexOf("Second."))
    expect(html.indexOf("Second.")).toBeLessThan(html.indexOf("Third."))
  })

  test.each([
    "The source quotes `<task_result>Thirty days.</task_result>`. The actual period is sixty days.",
    "The source quotes `<findings>Thirty days.</findings>`. The actual period is sixty days.",
    "The source quotes ``<findings>`Thirty days.`</findings>``. The actual period is sixty days.",
    "The source quotes &lt;findings&gt;Thirty days.&lt;/findings&gt;. The actual period is sixty days.",
    "The source quotes \\<findings>Thirty days.\\</findings>. The actual period is sixty days.",
    "```xml\n<findings>Thirty days.</findings>\n```\nThe actual period is sixty days.",
    "~~~xml\n<findings>Thirty days.</findings>\n~~~\nThe actual period is sixty days.",
  ])("leaves literal markup out of report recognition: %s", (text) => {
    const html = answer(text)
    expect(html).toContain("The actual period is sixty days.")
    expect(html).toContain("&lt;")
    expect(html).not.toContain("structured-report")
  })

  test("does not recognize report tags inside code within a real report", () => {
    const html = answer(
      "<findings>Use this example:\n\n```xml\n<notes>Literal only.</notes>\n```\n\nActual finding.</findings>",
    )
    expect(html).toContain('class="report-section-title">Findings')
    expect(html).toContain("&lt;notes&gt;Literal only.&lt;/notes&gt;")
    expect(html).toContain("Actual finding.")
    expect(html).not.toContain('class="report-section-title">Notes')
  })
})

describe("tool-origin report and read rendering", () => {
  test("retains task report styling, including escaped persisted envelopes", () => {
    const raw = '<task id="s" state="completed"><task_result><findings>Thirty days.</findings></task_result></task>'
    for (const text of [raw, raw.replaceAll("<", "&lt;").replaceAll(">", "&gt;")]) {
      const html = tool("task", text)
      expect(html).toContain('class="report-section-title">Findings')
      expect(html).toContain("Thirty days.")
    }
  })

  test("keeps a task's final qualifications alongside its child tool timeline", () => {
    const output =
      '<task id="child" state="completed"><task_result><findings>Thirty days.</findings>Do not rely on this until the missing exhibit is checked.</task_result></task>'
    const task: Part = {
      type: "tool",
      id: "task-call",
      sessionID: "session",
      messageID: "message",
      tool: "task",
      state: { status: "completed", input: {}, title: "Review", output, metadata: { sessionId: "child" } },
    }
    const childRead: Part = {
      type: "tool",
      id: "child-read",
      sessionID: "child",
      messageID: "child-message",
      tool: "read",
      state: {
        status: "completed",
        input: {},
        title: "Read exhibit",
        output: "<path>/synthetic/exhibit.txt</path><content>1: Child source evidence.</content>",
      },
    }
    const render = (pool?: Part[]) => {
      const turns = toTurns([{ info: { id: "message", role: "assistant" as const }, parts: [task] }], "", pool)
      return renderToStaticMarkup(createElement(StepsPanel, { steps: turns[0]!.steps, busy: false, answered: true }))
    }
    expect(render()).toContain("Do not rely on this until the missing exhibit is checked.")
    const html = render([childRead])
    expect(html).toContain("Child source evidence.")
    expect(html).toContain("Thirty days.")
    expect(html).toContain("Do not rely on this until the missing exhibit is checked.")
  })

  test("never drops prefix or suffix warnings around a task result", () => {
    const html = tool(
      "task",
      'UNVERIFIED.\n<task id="s" state="completed"><task_result><findings>Thirty days.</findings></task_result></task>\nCheck the exhibit.',
    )
    expect(html).toContain("UNVERIFIED.")
    expect(html).toContain("Check the exhibit.")
    expect(html.indexOf("UNVERIFIED.")).toBeLessThan(html.indexOf("Thirty days."))
    expect(html.indexOf("Thirty days.")).toBeLessThan(html.indexOf("Check the exhibit."))
  })

  test("retains leading read warnings while preserving useful file formatting", () => {
    const raw = "<path>/synthetic/contract.txt</path><content>1: Thirty days.</content>"
    expect(tool("read", raw)).toContain('class="tool-output-label">Content')
    const html = tool("read", "UNVERIFIED EXTRACTION.\n" + raw + "\nCheck the exhibit.")
    expect(html).toContain("UNVERIFIED EXTRACTION.")
    expect(html).toContain("Check the exhibit.")
    expect(html).toContain("Thirty days.")
  })

  test("keeps raw and escaped report markers in read-file content opaque", () => {
    for (const literal of ["<findings>Thirty days.</findings>", "&lt;findings&gt;Thirty days.&lt;/findings&gt;"]) {
      const html = tool(
        "read",
        `<path>/synthetic/contract.txt</path><content>1: ${literal}\n2: The actual period is sixty days.</content>`,
      )
      expect(html).toContain("The actual period is sixty days.")
      expect(html).toContain('class="tool-output-label">Content')
      expect(html).not.toContain("structured-report")
    }
  })

  test("keeps literal content closing markers inside the file body", () => {
    const html = tool(
      "read",
      "<path>/synthetic/contract.txt</path><content>1: Example </content> remains literal.\n2: The actual period is sixty days.</content>",
    )
    expect(html).toContain("Example &lt;/content&gt; remains literal.")
    expect(html).toContain("The actual period is sixty days.")
  })

  test("keeps unrelated tool results out of legacy report recognition", () => {
    const html = tool("search-document", "Literal <findings>Thirty days.</findings>. The actual period is sixty days.")
    expect(html).toContain("The actual period is sixty days.")
    expect(html).toContain("&lt;findings&gt;")
    expect(html).not.toContain("structured-report")
  })
})
