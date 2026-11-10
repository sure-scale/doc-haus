import { expect, test } from "bun:test"
import { scanMarkdownStructure } from "./markdown-structure"
import { visibleLines, visibleSectionLines, visibleTextMatches } from "./markdown-visible-text"
import { inspectDraftMarkdownNamespace } from "./markdown"
import { draftingPlaceholderLocations } from "./placeholder-resolution"

test("container fences end at their actual boundary and later prose resumes inline reading", () => {
  for (const source of [
    "- ```text\n  A\\-B\n  ```\n\nNext *prose*",
    "> ```text\n> A\\-B\n> ```\n\nNext *prose*",
  ]) {
    const lines = visibleSectionLines(source)
    expect(lines.map(line => line.kind)).toEqual(["fence_delimiter", "fenced", "fence_delimiter", "text", "text"])
    expect(lines[1]?.visible).toBe(lines[1]?.source)
    expect(lines[4]?.visible).toBe("Next prose")
  }
})

test("indentation is code only when the document tokenizer owns it as code", () => {
  expect(visibleSectionLines("ordinary paragraph\n    **continued prose**").map(line => [line.kind, line.visible])).toEqual([
    ["text", "ordinary paragraph"], ["text", "    continued prose"],
  ])
  expect(visibleSectionLines("- item\n\n      **literal code**").at(-1)?.kind).toBe("indented_code")
  expect(visibleSectionLines("> paragraph\ncontinued **lazy prose**").map(line => line.kind)).toEqual(["text", "text"])
})

test("contained literal code cannot suppress advisory findings in subsequent real prose", () => {
  const prose = 'The "**Supply Group**" delivers. The "**Buyer Group**" pays. The "**Review Group**" confirms.\nThe Unknown Party pays. The Unknown Party receives.'
  const expected = inspectDraftMarkdownNamespace(prose)
  expect(expected).toHaveLength(4)
  const actual = inspectDraftMarkdownNamespace(`- \`\`\`text\n  literal\n  \`\`\`\n\n${prose}`)
  expect(actual.map(({ line, ...issue }) => ({ ...issue, line: line - 4 }))).toEqual(expected)
})

test("literal placeholder fallback retains exact UTF-16 source offsets in containers and CRLF", () => {
  const source = "🙂 prefix\r\n\r\n> ```text\r\n> [___]\r\n> ```\r\n\r\nPay on \\[\\_\\_\\_\\]."
  const matches = visibleTextMatches(source, [/\[_{2,}\]/g])
  expect(matches.map(match => match.raw)).toEqual(["[___]", "\\[\\_\\_\\_\\]"])
  for (const match of matches) expect(source.slice(match.start, match.end)).toBe(match.raw)
  expect(matches[0]?.start).toBe(source.indexOf("[___]"))
  expect(matches[0]?.line).toBe(4)
  expect(matches[0]?.column).toBe(3)
  expect(draftingPlaceholderLocations(source).map(item => [item.placeholder, item.line, item.column])).toEqual([
    ["[___]", 4, 3], ["\\[\\_\\_\\_\\]", 7, 8],
  ])
  expect(visibleLines(source).map(line => line.raw)).toEqual(source.split("\r\n"))
})

test("named top-level fence envelopes retain their explicit legacy contract", () => {
  const source = "```packet\nline\n```\n\n  ```packet\nindented\n  ```\n\n- ```packet\n  nested\n  ```"
  const result = scanMarkdownStructure(source, ["packet"])
  expect(result.fences).toEqual([{ name: "packet", info: "packet", payload: "line", raw: "```packet\nline\n```", complete: true, line: 1 }])
  const unfinished = "```packet\r\nline\r\n"
  expect(scanMarkdownStructure(unfinished, ["packet"]).fences).toEqual([{ name: "packet", info: "packet", payload: "line\n", raw: "```packet\nline\n", complete: false, line: 1 }])
})

test("shorter or mismatched fences and unclosed container code retain literal ownership", () => {
  const source = "> ````text\n> A\\-B\n> ```\n> ~~~\n> ````\n\nNext *prose*"
  expect(visibleSectionLines(source).map(line => line.kind)).toEqual(["fence_delimiter", "fenced", "fenced", "fenced", "fence_delimiter", "text", "text"])
  expect(visibleSectionLines(source).at(-1)?.visible).toBe("Next prose")
  const unfinished = "> ```text\n> A\\-B\n\nNext *prose*"
  expect(visibleSectionLines(unfinished).at(-1)?.visible).toBe("Next prose")
})
