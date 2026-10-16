import { expect, test } from "bun:test"
import { scanMarkdownStructure } from "./markdown-structure"
import {
  visibleInlineSegments,
  visibleInlineSpans,
  visibleInlineText,
  visibleLines,
  visibleRawSpan,
  visibleSectionLines,
  visibleSectionText,
  visibleTextMatches,
} from "./markdown-visible-text"

const ASCII_PUNCTUATION = Array.from("!\"#$%&'()*+,-./:;<=>?@[\\]^_`{|}~")

test("every escaped ASCII punctuation character is its literal character", () => {
  expect(ASCII_PUNCTUATION).toHaveLength(32)
  for (const character of ASCII_PUNCTUATION) {
    expect(visibleInlineText(`\\${character}`)).toBe(character)
    expect(visibleInlineText(`Rate\\${character}Card`)).toBe(`Rate${character}Card`)
  }
})

const VISIBLE_ROWS = [
  ["a double backslash", "a\\\\b", "a\\b"],
  ["an escaped backslash before an escape", "a\\\\\\-b", "a\\-b"],
  ["a backslash before a letter", "a\\q", "a\\q"],
  ["a backslash before a digit", "a\\3", "a\\3"],
  ["a backslash before a space", "a\\ b", "a\\ b"],
  ["a trailing backslash", "a\\", "a\\"],
  ["a backslash before a non-ASCII letter", `a\\${String.fromCodePoint(966)}`, `a\\${String.fromCodePoint(966)}`],
  ["emphasis", "*x y*", "x y"],
  ["strong emphasis", "**x y**", "x y"],
  ["underscore emphasis", "_x y_", "x y"],
  ["a code span", "`A-B`", "A-B"],
  ["a character reference", "&amp;", "&"],
  ["a numeric character reference", "&#45;", "-"],
  ["a link", "[label](destination \"title\")", "label"],
  ["an autolink", "<https://example.test/a_b>", "https://example.test/a_b"],
  ["a hard line break", "a\\\nb", "a \nb"],
] as const

test.each(VISIBLE_ROWS)("%s shows only its visible text", (_, source, visible) => {
  expect(visibleInlineText(source)).toBe(visible)
})

const LITERAL_ROWS = [
  ["an escaped underscore", "Rate\\_Card", "Rate_Card"],
  ["an intraword underscore", "Rate_Card", "Rate_Card"],
  ["an unmatched asterisk", "Rate*Card", "Rate*Card"],
  ["an escaped asterisk", "Rate\\*Card", "Rate*Card"],
  ["an escaped backtick", "Rate\\`Card", "Rate`Card"],
  ["a lone backtick", "Rate`Card", "Rate`Card"],
  ["a backslash inside a code span", "`A\\-B`", "A\\-B"],
  ["an escape inside a code span", "`Rate\\_Card`", "Rate\\_Card"],
  ["raw inline HTML", "<span a=\"x\\-y\">", "<span a=\"x\\-y\">"],
  ["an autolink with a backslash", "<https://example.test/a\\-b>", "https://example.test/a\\-b"],
] as const

test.each(LITERAL_ROWS)("%s stays literal visible punctuation", (_, source, visible) => {
  expect(visibleInlineText(source)).toBe(visible)
  expect(visibleInlineText(source)).not.toBe(source.replace(/\\/g, "").replace(/[*_`]/g, ""))
})

const INLINE_ROWS = [
  ["a numbered heading text", "1. Title", "1. Title"],
  ["a list-looking scalar", "- Title", "- Title"],
  ["a quote-looking scalar", "> Title", "> Title"],
  ["a literal hash", "\\# Payment", "# Payment"],
] as const

test.each(INLINE_ROWS)("%s is read as inline text, never as block structure", (_, source, visible) => {
  expect(visibleInlineText(source)).toBe(visible)
})

test("segments map each visible run back to its raw span", () => {
  const source = "Pre\\-Agreed *fee* `A\\-B`"
  expect(visibleInlineSegments(source)).toEqual([
    { value: "Pre", start: 0, end: 3 },
    { value: "-", start: 3, end: 5 },
    { value: "Agreed ", start: 5, end: 12 },
    { value: "fee", start: 13, end: 16 },
    { value: " ", start: 17, end: 18 },
    { value: "A\\-B", start: 19, end: 23 },
  ])
})

test("section text classifies structure on raw lines and decodes only visible lines", () => {
  const section = [
    "## 1. Pre\\-Agreed Terms",
    "",
    "\\# Payment is due monthly.",
    "",
    "```text",
    "A\\-B *x*",
    "```",
    "",
    "    indented \\-x",
    "",
    "<div>",
    "raw \\-html",
    "</div>",
    "",
    "Visible \\| pipe <!-- hidden --> end",
  ].join("\n")
  const structure = scanMarkdownStructure(section, [])
  expect(structure.visibleLines.find((line) => line.line === 3)?.value).toBe("\\# Payment is due monthly.")
  expect(structure.visibleLines.filter((line) => /^##\s/.test(line.value)).map((line) => line.line)).toEqual([1])
  const lines = visibleSectionLines(section)
  expect(lines).toHaveLength(section.split("\n").length)
  expect(lines.map((line) => line.source)).toEqual(section.split("\n"))
  expect(lines.map((line) => line.visible)).toEqual([
    "## 1. Pre-Agreed Terms",
    "",
    "# Payment is due monthly.",
    "",
    "```text",
    "A\\-B *x*",
    "```",
    "",
    "    indented \\-x",
    "",
    "<div>",
    "raw \\-html",
    "</div>",
    "",
    "Visible | pipe  end",
  ])
  expect(visibleSectionText(section)).toBe(lines.map((line) => line.visible).join("\n"))
  expect(lines[2]!.segments[0]).toEqual({ value: "#", start: 0, end: 2 })
})

test.each([
  ["a code span continued onto the next line keeps its escape", "The identifier is `A\\-B\nC`.", ["The identifier is A\\-B", "C."]],
  ["a comment inside a code span is literal code", "x `red<!--x-->blue` y", ["x red<!--x-->blue y"]],
  ["an HTML comment shows nothing and adds no space", "red<!--x-->blue", ["redblue"]],
  ["a comment continued onto the next line shows nothing", "Before <!-- a\nb --> after", ["Before ", " after"]],
  ["an unclosed comment opener inside code hides no later prose", "`a <!-- b`\n\nNext \\- prose.", ["a <!-- b", "", "Next - prose."]],
] as const)("section text reads each paragraph as one CommonMark inline run: %s", (_, section, visible) => {
  const lines = visibleSectionLines(section)
  expect(lines.map((line) => line.visible)).toEqual([...visible])
  for (const line of lines) {
    for (const segment of line.segments) expect(line.source.slice(segment.start, segment.end).length).toBeGreaterThan(0)
  }
})

test("a fragment and a section show an HTML comment and literal code the same way", () => {
  expect(visibleInlineText("red<!--x-->blue")).toBe("redblue")
  expect(visibleInlineText("`red<!--x-->blue`")).toBe("red<!--x-->blue")
  expect(visibleSectionText("red<!--x-->blue")).toBe(visibleInlineText("red<!--x-->blue"))
  expect(visibleSectionText("`red<!--x-->blue`")).toBe(visibleInlineText("`red<!--x-->blue`"))
})

test("visible spans agree with visible text and map every character to its raw token", () => {
  const sources = [
    ...ASCII_PUNCTUATION.flatMap((character) => [`\\${character}`, `Rate\\${character}Card`]),
    ...[...VISIBLE_ROWS, ...LITERAL_ROWS, ...INLINE_ROWS].map((row) => row[1]),
    "Pre\\-Agreed *fee* `A\\-B`",
  ]
  for (const source of sources) {
    const spans = visibleInlineSpans(source)
    expect(spans.text).toBe(visibleInlineText(source))
    expect(spans.map).toHaveLength(spans.text.length + 1)
    expect(spans.ends).toHaveLength(spans.text.length)
    expect(spans.map[spans.text.length]).toBe(source.length)
    spans.text.split("").forEach((character, index) => {
      const raw = source.slice(spans.map[index], spans.ends[index])
      // A literal byte is the character itself; any other token is an escape,
      // a character reference or a hard break that shows this one character.
      expect(raw === character || /^(?:\\|&|[ \t]+$)/.test(raw)).toBe(true)
      expect(raw.length).toBeGreaterThan(0)
      expect(spans.map[index + 1]!).toBeGreaterThanOrEqual(spans.map[index]!)
    })
  }
})

test.each([
  ["an escaped blank", "x \\[\\_\\_\\_\\] y", "[___]", "\\[\\_\\_\\_\\]"],
  ["an escaped structural reference", "Section \\[Notice Period\\]", "Section [Notice Period]", "Section \\[Notice Period\\]"],
  ["an emphasized blank", "*[___]*", "[___]", "[___]"],
  ["interior emphasis", "[*Insert* date]", "[Insert date]", "[*Insert* date]"],
] as const)("the raw span of %s covers exactly the bytes that show the match", (_, source, match, raw) => {
  const spans = visibleInlineSpans(source)
  const index = spans.text.indexOf(match)
  expect(index).toBeGreaterThanOrEqual(0)
  const span = visibleRawSpan(spans, index, index + match.length)
  expect(source.slice(span.start, span.end)).toBe(raw)
})

test("code span content is literal visible text: escapes stay, brackets show", () => {
  // Code-span ruling R1: literal (no decoding), not invisible.
  expect(visibleInlineText("`[___]`")).toBe("[___]")
  expect(visibleInlineText("`\\[\\_\\_\\_\\]`")).toBe("\\[\\_\\_\\_\\]")
  expect(visibleInlineSpans("`\\[x\\]`").text).toBe("\\[x\\]")
  expect(visibleTextMatches("`\\[x\\]`", [/\[x\]/g])).toEqual([])
  expect(visibleTextMatches("`[___]`", [/\[_{2,}\]/g]).map((match) => match.raw)).toEqual(["[___]"])
})

test("visible lines classify structure on raw lines and keep code and HTML literal", () => {
  const section = ["Pre\\-Agreed", "```text", "A\\-B", "```", "", "    code \\-x", "", "<div>", "raw \\-x", "</div>"].join("\n")
  const lines = visibleLines(section)
  expect(lines.map((line) => line.kind)).toEqual([
    "text", "fence_delimiter", "fenced", "fence_delimiter", "text", "indented_code", "text", "html_block", "html_block", "html_block",
  ])
  expect(lines.map((line) => line.raw)).toEqual(section.split("\n"))
  expect(lines.filter((line) => line.kind !== "text").every((line) => line.visible === line.raw)).toBe(true)
  expect(lines[0]!.visible).toBe("Pre-Agreed")
  expect(lines.map((line) => line.visible).join("\n")).toBe(visibleSectionText(section))
  expect(visibleSectionLines(section).map((line) => line.kind)).toEqual(lines.map((line) => line.kind))
})

const BLANK = /\[\s*_{2,}[^\]\n]*\]/g
const REFERENCE = /\bsection\s+\[[A-Za-z][^[\]\n]{1,80}\]/gi
const TAG = /<stage_packet>/g

test.each([
  ["paired underscore blanks the CommonMark reading turns into emphasis", "dated [___] between [___] and [___]", [BLANK], ["[___]", "[___]", "[___]"]],
  ["a blank inside an HTML comment", "<!-- [___] --> visible text", [BLANK], ["[___]"]],
  ["a reference split across a line break", "see Section\n[Notice Period] here", [REFERENCE], ["Section\n[Notice Period]"]],
  ["an escaped Word blank", "x \\[\\_\\_\\_\\] y", [BLANK], ["\\[\\_\\_\\_\\]"]],
  ["an escaped tag the structure scanner reads as HTML", "<stage\\_packet\\>", [TAG], ["<stage\\_packet\\>"]],
  ["one blank found by two patterns", "x [___] y", [BLANK, /\[_+\]/g], ["[___]"]],
] as const)("detection never sees less than the raw text: %s", (_, text, patterns, raws) => {
  const matches = visibleTextMatches(text, patterns)
  expect(matches.map((match) => match.raw)).toEqual([...raws])
  for (const match of matches) {
    expect(text.slice(match.start, match.end)).toBe(match.raw)
    const lineStart = text.split("\n").slice(0, match.line - 1).join("\n").length + (match.line > 1 ? 1 : 0)
    expect(match.start - lineStart + 1).toBe(match.column)
  }
})

test("an escaped Word blank is found once at its visible span, never again at its raw tail", () => {
  const text = "Before.\n\nPay on \\[\\_\\_ notice timing\\] and \\[Insert date\\].\n"
  const matches = visibleTextMatches(text, [BLANK, /\[(?:insert)\b[^\]\n]*\]/gi])
  expect(matches.map((match) => [match.raw, match.visible, match.line, match.column])).toEqual([
    ["\\[\\_\\_ notice timing\\]", "[__ notice timing]", 3, 8],
    ["\\[Insert date\\]", "[Insert date]", 3, 35],
  ])
})
