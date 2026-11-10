import { afterEach, describe, expect, test } from "bun:test"
import { DOMParser } from "@xmldom/xmldom"
import { strFromU8, unzipSync } from "fflate"
import { splitBlocks } from "./markdown"
import { renderMarkdownToDocx } from "./render"
import { fillTemplate } from "./template-fill"
import createTemplate from "../tool/create-template"
import type { LegalToolContext } from "../tool/legacy-definition"

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const originalFetch = globalThis.fetch
afterEach(() => { globalThis.fetch = originalFetch })

// Delimiters that look like a closing fence are literal unless the maintained
// CommonMark parser says they close this particular block.
const cases = [
  { name: "long backtick fence with short and mixed interior fences",
    source: "````text\nfirst\\\n```\n~~~\nlast\\\n`````", value: "first\\\n```\n~~~\nlast\\" },
  { name: "tilde fence with backticks and a too-short tilde fence",
    source: "~~~~\n```\n~~~\nlast\\\n~~~~~", value: "```\n~~~\nlast\\" },
  { name: "literal alignment wrappers, blank lines and tabs",
    source: "```\n<center>\n\nA & B\tC\\\n\n</center>\n```", value: "<center>\n\nA & B\tC\\\n\n</center>" },
  { name: "indented pipe code", source: "    | A | B |\n    |---|---|\n\n    | C | D |", value: "| A | B |\n|---|---|\n\n| C | D |" },
  { name: "quoted pipe code", source: "> ~~~~\n> | A | B |\n> |---|---|\n>\n> | C | D |\n> ~~~~", value: "| A | B |\n|---|---|\n\n| C | D |" },
  { name: "list-contained pipe code", source: "- ````\n  | A | B |\n  |---|---|\n  | C | D |\n  ````", value: "| A | B |\n|---|---|\n| C | D |" },
] as const

function body(code: string) { return `Before\\\n\n${code}\n\n<center>\nAfter\\\n</center>` }

function assertLiteral(bytes: Uint8Array, expected: string, label: string) {
  const xml = strFromU8(unzipSync(bytes)["word/document.xml"]!)
  const document = new DOMParser().parseFromString(xml, "application/xml")
  expect(document.getElementsByTagNameNS(W, "tbl").length, label).toBe(0)
  const paragraphs = Array.from(document.getElementsByTagNameNS(W, "p")).map(paragraph =>
    Array.from(paragraph.getElementsByTagNameNS(W, "*")).map(element =>
      element.localName === "t" ? element.textContent : element.localName === "br" ? "\n" : element.localName === "tab" ? "\t" : "").join(""))
  expect(paragraphs, label).toEqual(["Before", expected, "After"])
  expect(xml, label).not.toContain("DocHausText")
}

const context: LegalToolContext = {
  directory: "/fixture/workspace", worktree: "/fixture/workspace", agent: "template-builder",
  sessionID: "literal-code-template", messageID: "message", callID: "call",
  abort: new AbortController().signal, metadata() {}, async ask() {},
}

async function create(content: string) {
  let bytes: Uint8Array | undefined
  globalThis.fetch = (async (_input, options) => {
    if (options?.method !== "POST") return Response.json({ templates: [] })
    const file = (options.body as FormData).get("file") as File
    bytes = new Uint8Array(await file.arrayBuffer())
    return Response.json({})
  }) as typeof fetch
  const result = await createTemplate.execute({ name: "literal-code.docx", description: "Literal syntax example", content }, context)
  if (typeof result === "string" || !bytes) throw new Error(`Template creation failed: ${JSON.stringify(result)}`)
  return { bytes, result }
}

describe("literal code through drafting consumers", () => {
  test("splitBlocks preserves code source while keeping prose boundaries and alignment behavior", () => {
    for (const fixture of cases)
      expect(splitBlocks(body(fixture.source)), fixture.name).toBe(`Before\n\n${fixture.source}\n\nAfter`)
  })

  test("actual render keeps code literal and never creates a table from its pipes", async () => {
    for (const fixture of cases) {
      const { bytes } = await renderMarkdownToDocx(body(fixture.source), undefined, undefined, { numbering: false, styleTemplate: false })
      assertLiteral(bytes, fixture.value, fixture.name)
    }
  })

  test("actual template creation preserves code in the uploaded DOCX", async () => {
    for (const fixture of cases) assertLiteral((await create(body(fixture.source))).bytes, fixture.value, fixture.name)
  })

  test("template fill content and clause replacement preserve literal code", async () => {
    const seed = await Bun.file(new URL("../templates/_base.docx", import.meta.url)).bytes()
    const template = (await create("Replacement target")).bytes
    for (const fixture of cases) {
      assertLiteral((await fillTemplate(seed, { content: body(fixture.source) })).bytes, fixture.value, `content: ${fixture.name}`)
      assertLiteral((await fillTemplate(template, { replaces: [{ clause: "Replacement target", replacement: body(fixture.source) }] })).bytes,
        fixture.value, `replacement: ${fixture.name}`)
    }
  })

  test("code placeholders stay discoverable and fillable after literal restoration", async () => {
    const content = body("~~~~\n[insert command]\n\nend\\\n~~~~")
    const created = await create(content)
    expect(created.result.metadata.placeholders).toContain("[insert command]")
    const fills = { "[insert command]": "run | exact" }
    assertLiteral((await fillTemplate(created.bytes, { fills })).bytes, "run | exact\n\nend\\", "existing template fill")
    const seed = await Bun.file(new URL("../templates/_base.docx", import.meta.url)).bytes()
    assertLiteral((await fillTemplate(seed, { content, fills })).bytes, "run | exact\n\nend\\", "content fill")
  })
})
