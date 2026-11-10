import { expect, test } from "bun:test"
import { isMalformedToolOutput } from "./malformed-tool-output"

const classify = (text: string) => isMalformedToolOutput([{ type: "text", text }])
const call = 'call:any-future-tool{"input":{"value":"a } and \\\"quoted\\\" value"}}'

test("bare supported tool transports recover, including incomplete object payloads", () => {
  for (const text of ["<|tool_call|>", call, `<|tool_call|>${call}`, `\n   ${call}\n`, 'call:task{"description":', '<|tool_call|>call:task{'])
    expect(classify(text)).toBe(true)
})

test("literal Markdown and explanatory prose never admit another model turn", () => {
  for (const text of [
    `\`\`\`json\n${call}\n\`\`\``, `~~~\n${call}\n~~~`, `    ${call}`, `\t${call}`,
    `\`${call}\``, `> ${call}`, `- ${call}`, `Explain ${call}`, `${call}\nThis is an example.`,
    `Here is an example: \`${call}\`.`, `<pre>${call}</pre>`, '<|tool_call|> is a literal marker.',
    `${call}\n\n\`\`\`json\n{}\n\`\`\``,
  ]) expect(classify(text)).toBe(false)
})

test("native chunking cannot change protocol or literal context", () => {
  for (const text of [call, `<|tool_call|>${call}`, `\`\`\`json\n${call}\n\`\`\``, `Here is an example: \`${call}\`.`, `${call}\nThis is an example.`]) {
    const expected = classify(text)
    for (let at = 0; at <= text.length; at++) {
      expect(isMalformedToolOutput([{ type: "text", text: text.slice(0, at) }, { type: "text", text: text.slice(at) }])).toBe(expected)
      expect(isMalformedToolOutput([{ type: "reasoning", text: text.slice(0, at) }, { type: "reasoning", text: text.slice(at) }])).toBe(expected)
    }
  }
})

test("reasoning protocol cannot replace an ordinary or literal text answer", () => {
  for (const text of ["Here is the completed answer.", `Here is an example: \`${call}\`.`, `\`\`\`\n${call}\n\`\`\``])
    expect(isMalformedToolOutput([{ type: "reasoning", text: call }, { type: "text", text }])).toBe(false)
  expect(isMalformedToolOutput([{ type: "reasoning", text: call }, { type: "text", text: "generated context", synthetic: true }])).toBe(true)
  expect(isMalformedToolOutput([{ type: "text", text: call, synthetic: true }])).toBe(false)
  expect(isMalformedToolOutput([{ type: "tool", text: call }])).toBe(false)
})
