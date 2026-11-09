import { describe, expect, test } from "bun:test"
import { createElement } from "react"
import { renderToStaticMarkup } from "react-dom/server"
import MarkdownContent from "../src/components/markdown/MarkdownContent"

const render = (text: string) => renderToStaticMarkup(createElement(MarkdownContent, { children: text }))
const table = "| Item | Term |\n| :--- | ---: |\n| Notice | 30 days |"

describe("Markdown tables use their surrounding syntax tree", () => {
  test("keeps table styling, alignment and prose in order", () => {
    const html = render(`Before.\n\n${table}\n\nAfter.`)
    expect(html).toContain('class="markdown-table-scroll"')
    expect(html).toContain('class="markdown-table"')
    expect(html).toContain('<th style="text-align:left">Item</th>')
    expect(html).toContain('<td style="text-align:right">30 days</td>')
    expect(html.indexOf("Before.")).toBeLessThan(html.indexOf("<table"))
    expect(html.indexOf("</table>")).toBeLessThan(html.indexOf("After."))
  })

  test.each(["```text", "~~~~text"])("leaves fenced tables literal: %s", (fence) => {
    const close = fence.startsWith("`") ? "```" : "~~~~"
    const html = render(`${fence}\n${table}\n${close}`)
    expect(html).toContain('<pre><code class="language-text">')
    expect(html).toContain("| Notice | 30 days |")
    expect(html).not.toContain("<table")
  })

  test("leaves indented tables and escaped delimiters outside a table literal", () => {
    const html = render(table.split("\n").map(line => `    ${line}`).join("\n") + "\n\n\\| Literal \\| text \\|")
    expect(html).toContain("<pre><code>")
    expect(html).toContain("| Literal | text |")
    expect(html).not.toContain("<table")
  })

  test("preserves escaped pipes, nested inline formatting, Unicode and legal placeholders", () => {
    const html = render("| Literal | Content |\n|---|---|\n| A\\|B | **échéance** [party] [source: missing] [1, 2] |\n| Code | `x\\|y` |")
    expect(html.match(/<tr>/g)).toHaveLength(3)
    expect(html).toContain("A|B")
    expect(html).toContain("<strong>échéance</strong> [party] [source: missing] [1, 2]")
    expect(html).toContain("<code>x|y</code>")
    expect(html).not.toContain("<button")
  })

  test("resolves ordinary Markdown references across table and prose through one parser", () => {
    const html = render("| Item | Source |\n|---|---|\n| Notice | [clause][evidence] |\n\nOutside [also][evidence].\n\n[evidence]: /documents/terms \"Terms\"")
    expect(html.match(/href="\/documents\/terms" title="Terms"/g)).toHaveLength(2)
  })

  test("retains blockquote and list containment", () => {
    const quoted = render(table.split("\n").map(line => `> ${line}`).join("\n"))
    expect(quoted).toContain('<blockquote>\n<div class="markdown-table-scroll">')
    const listed = render("- A term:\n\n" + table.split("\n").map(line => `  ${line}`).join("\n"))
    expect(listed).toContain("<li>")
    expect(listed.indexOf("<li>")).toBeLessThan(listed.indexOf("<table"))
    expect(listed.indexOf("</table>")).toBeLessThan(listed.indexOf("</li>"))
  })

  test("does not enable unrelated GFM syntax", () => {
    const html = render("~~literal~~ [ ] task www.example.invalid person@example.invalid")
    expect(html).toContain("~~literal~~ [ ] task www.example.invalid person@example.invalid")
    expect(html).not.toContain("<del>")
    expect(html).not.toContain("<a ")
    expect(html).not.toContain("checkbox")
  })

  test("streaming fence prefixes and persisted completed text retain one interpretation", () => {
    const text = `\`\`\`text\n${table}\n\`\`\`\n\n${table}`
    const streaming = render(text.slice(0, text.indexOf("\n```", 3)))
    expect(streaming).not.toContain("<table")
    const completed = render(text)
    expect(completed.match(/<table /g)).toHaveLength(1)
    expect(render(JSON.parse(JSON.stringify({ text })).text)).toBe(completed)
  })
})
