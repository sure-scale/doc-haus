// The complex fields of a paragraph (w:fldChar begin/separate/end with a
// w:instrText instruction and a cached result), read and unwrapped as XML. Shared
// by the cross-reference linker and the writers that edit paragraphs holding
// live cross-references.

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const NUMBER_SWITCHES = new Set(["\\r", "\\n", "\\w"])

/** A complex field of a paragraph at the paragraph's own level: its control
 * elements, its instruction, and where its cached result is in the
 * paragraph's text. `nested` marks a field inside another's instruction. */
export type ParagraphField = { controls: Element[]; instruction: string; from: number; to: number; result: string; texts: Element[]; separate?: Element; nested: boolean }

/** The complex fields of a paragraph, outermost only. Offsets count w:t text,
 * as the paragraph's text does. */
export function paragraphFields(paragraph: Element): ParagraphField[] {
  const found: ParagraphField[] = []
  const stack: (ParagraphField & { phase: "instruction" | "result" })[] = []
  let offset = 0
  for (const node of elements(paragraph, "*")) {
    const top = stack.at(-1)
    if (node.localName === "fldChar") {
      const type = node.getAttributeNS(W, "fldCharType")
      if (type === "begin") {
        if (top?.phase === "instruction") top.nested = true
        stack.push({ controls: [node], instruction: "", from: offset, to: offset, result: "", texts: [], nested: false, phase: "instruction" })
      } else if (top && stack.length === 1) {
        top.controls.push(node)
        if (type === "separate") { top.phase = "result"; top.separate = node; top.from = offset }
        if (type === "end") { const { phase: _phase, ...field } = stack.pop()!; found.push({ ...field, from: field.separate ? field.from : offset, to: offset }) }
      } else if (type === "end") stack.pop()
    } else if (node.localName === "instrText") {
      if (stack.length === 1 && top!.phase === "instruction") { top!.instruction += node.textContent ?? ""; top!.controls.push(node) }
    } else if (node.localName === "t") {
      const text = node.textContent ?? ""
      if (stack[0]?.phase === "result") { stack[0].result += text; stack[0].texts.push(node) }
      offset += text.length
    }
  }
  return found
}

/** A REF field's bookmark and number switch; undefined for any other field,
 * or a REF with switches this module does not keep. */
export function refInstruction(field: ParagraphField) {
  if (field.nested) return
  const tokens = field.instruction.trim().split(/\s+/)
  if (tokens[0]?.toUpperCase() !== "REF" || !tokens[1]) return
  let number: string | undefined
  for (let at = 2; at < tokens.length; at += 1) {
    const token = tokens[at]!
    if (token === "\\h") continue
    if (NUMBER_SWITCHES.has(token)) { number = token; continue }
    if (token === "\\*" && /^(?:MERGEFORMAT|CHARFORMAT)$/i.test(tokens[at + 1] ?? "")) { at += 1; continue }
    return
  }
  return { name: tokens[1], number }
}

/** The field as the plain text it shows, in the result's own formatting. */
export function unlink(field: ParagraphField) {
  for (const node of field.controls) {
    const run = node.parentNode as Element
    run.removeChild(node)
    if (run.localName === "r" && !Array.from(run.childNodes).some(child => child.nodeType === 1 && (child as Element).localName !== "rPr")) run.parentNode?.removeChild(run)
  }
}

/** A paragraph's text as Docxodus reads it: its w:t text, field results included. */
export function runText(paragraph: Element) {
  return elements(paragraph, "t").map(node => node.textContent ?? "").join("")
}

function elements(node: Document | Element, name: string) {
  return Array.from(node.getElementsByTagNameNS(W, name))
}
