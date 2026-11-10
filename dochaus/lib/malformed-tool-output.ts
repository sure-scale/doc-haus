type OutputPart = { type: string; text?: unknown; synthetic?: unknown }

/** Recover only a bare printed transport, never syntax mentioned in an answer.
 * These are the existing supported spellings: a tool_call marker, call:name
 * with an object payload, or both. A truncated payload is still failed protocol;
 * no JSON is repaired or executed. Markdown wrappers and surrounding prose do
 * not belong to this grammar and remain ordinary model output.
 */
function bareToolProtocol(source: string): boolean {
  // Preserve leading indentation: trim() would turn an indented code example
  // into a bare call. Blank lines and CommonMark's up-to-three-space prefix do
  // not create a code block.
  const prefix = /^(?:[ \t]*\r?\n)* {0,3}(?=<\|tool_call\|>|call:)/i.exec(source)
  if (!prefix) return false
  let body = source.slice(prefix[0].length).trimEnd()
  const marker = /^<\|tool_call\|>/i.exec(body)
  if (marker) {
    body = body.slice(marker[0].length).trimStart()
    if (!body) return true
  }
  const call = /^call:[a-z][\w-]*\s*\{/i.exec(body)
  if (!call) return false
  let depth = 1, quoted = false, escaped = false
  for (let index = call[0].length; index < body.length; index++) {
    const character = body[index]!
    if (quoted) {
      if (escaped) escaped = false
      else if (character === "\\") escaped = true
      else if (character === '"') quoted = false
    } else if (character === '"') quoted = true
    else if (character === "{") depth++
    else if (character === "}" && --depth === 0) return !body.slice(index + 1).trim()
  }
  return true
}

/** Native text parts are chunks of one answer, not semantic boundaries. Join
 * without inventing separators so splitting a delimiter, payload or sentence
 * cannot change admission. Reasoning is separate: a real text answer always
 * prevents recovery, even when reasoning printed bare protocol.
 */
export function isMalformedToolOutput(parts: readonly OutputPart[]): boolean {
  const text = parts.filter(part => part.type === "text" && part.synthetic !== true && typeof part.text === "string")
    .map(part => part.text as string).join("")
  if (text.trim()) return bareToolProtocol(text)
  const reasoning = parts.filter(part => part.type === "reasoning" && typeof part.text === "string")
    .map(part => part.text as string).join("")
  return bareToolProtocol(reasoning)
}
