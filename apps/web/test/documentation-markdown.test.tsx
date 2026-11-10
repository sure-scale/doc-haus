import { expect, test } from "bun:test"
import { renderToStaticMarkup } from "react-dom/server"
import { DocMarkdown } from "../src/components/markdown/Markdown"

test("instruction tags remain visible without altering literal code bodies", () => {
  const example = '<task>\n\nKeep the `tick` and [source: later].\n\n</task>'
  const html = renderToStaticMarkup(<DocMarkdown>{'````text\n' + example + '\n````'}</DocMarkdown>)
  expect(html).toContain('<pre><code class="language-text">&lt;task&gt;\n\nKeep the `tick` and [source: later].\n\n&lt;/task&gt;\n</code></pre>')
})

test("standalone tags are literal while surrounding instructions retain Markdown", () => {
  const html = renderToStaticMarkup(<DocMarkdown>{'<task>\n\n**Keep qualifications.**\n\n</task>'}</DocMarkdown>)
  expect(html).toContain('&lt;task&gt;')
  expect(html).toContain('&lt;/task&gt;')
  expect(html).toContain('<strong>Keep qualifications.</strong>')
})
