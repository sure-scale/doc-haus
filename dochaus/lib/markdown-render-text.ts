import { randomUUID } from "node:crypto"
import { parse, postprocess, preprocess } from "micromark"
import { decodeString } from "micromark-util-decode-string"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

// Restore literal text after the native Markdown parser and table polish have
// finished interpreting grammar. The SDK supports triple-backtick code, but
// not every CommonMark fence or indented code block. Let micromark own those
// boundaries and pass their content through one inline-code token per block.
// Restoration writes real Word breaks/tabs, preserving the code run's style.
export function prepareMarkdownRenderText(content: string, options: { codeOnly?: true } = {}) {
  const prefix = `DocHausText${randomUUID().replaceAll("-", "")}Z`
  if (content.includes(prefix)) throw new Error("Markdown render text token collision.")
  const scopes: string[] = []
  const events = postprocess(parse().document().write(preprocess()(content, undefined, true)))
  type Reference = { start: number; end: number; value: string; replacement?: string }
  const literals: Reference[] = []
  let code: { start: number; end: number; first: number; last: number; fenced: boolean; closing?: number; values: Map<number, string> } | undefined
  for (const [kind, token, context] of events) {
    if (kind === "enter" && ["codeFenced", "codeIndented"].includes(token.type)) {
      code = { start: token.start.offset!, end: token.end.offset!, first: token.start.line,
        last: token.end.line - (token.end.column === 1 ? 1 : 0), fenced: token.type === "codeFenced", values: new Map() }
    } else if (code && kind === "enter" && token.type === "codeFencedFence" && token.start.line !== code.first) {
      code.closing = token.start.line
    } else if (code && kind === "enter" && token.type === "codeFlowValue") {
      code.values.set(token.start.line, context.sliceSerialize(token))
    } else if (code && kind === "exit" && ["codeFenced", "codeIndented"].includes(token.type)) {
      const first = code.first + (code.fenced ? 1 : 0)
      const last = code.closing === undefined ? code.last : code.closing - 1
      const value = Array.from({ length: Math.max(0, last - first + 1) }, (_, i) => code!.values.get(first + i) ?? "").join("\n")
      literals.push({ start: code.start, end: code.end, value,
        replacement: "`" + `${prefix}${literals.length}Z` + "`" })
      code = undefined
    }
  }
  if (!options.codeOnly) literals.push(...events.flatMap(([kind, token]) => {
      if (kind === "enter") {
        scopes.push(token.type)
        return []
      }
      const text = !scopes.some(scope => [
        "codeFencedFenceInfo", "codeFencedFenceMeta", "resourceDestinationString",
        "resourceTitleString", "definitionDestinationString", "definitionTitleString",
      ].includes(scope))
      scopes.pop()
      return ["characterReference", "characterEscape"].includes(token.type) && text
        ? [{ start: token.start.offset!, end: token.end.offset!,
            value: decodeString(content.slice(token.start.offset!, token.end.offset!)) }]
        : []
    }))
  const references = literals.map((reference, index) => ({ ...reference, key: `${prefix}${index}Z` }))
  const ordered = references.toSorted((a, b) => b.start - a.start)
  return {
    hasCode: references.some(reference => reference.replacement !== undefined),
    markdown: ordered.reduce((text, reference) =>
      text.slice(0, reference.start) + (reference.replacement ?? reference.key) + text.slice(reference.end), content),
    restore(bytes: Uint8Array) {
      if (!references.length) return bytes
      const parts = unzipSync(bytes)
      if (!parts["word/document.xml"])
        throw new Error("Markdown render text requires word/document.xml.")
      const counts = new Map(references.map(reference => [reference.key, 0]))
      const document = new DOMParser().parseFromString(strFromU8(parts["word/document.xml"]), "application/xml")
      const byKey = new Map(references.map(reference => [reference.key, reference]))
      const pattern = new RegExp(`${prefix}\\d+Z`, "g")
      for (const text of Array.from(document.getElementsByTagNameNS(W, "t"))) {
        const value = text.textContent ?? ""
        const matches = [...value.matchAll(pattern)]
        if (!matches.length) continue
        const parent = text.parentNode!
        const append = (value: string) => {
          const literal = text.cloneNode(false)
          literal.appendChild(document.createTextNode(value))
          parent.insertBefore(literal, text)
        }
        let cursor = 0
        for (const match of matches) {
          const reference = byKey.get(match[0])
          if (!reference) throw new Error("Markdown render text contains an unknown literal token.")
          append(value.slice(cursor, match.index))
          counts.set(reference.key, counts.get(reference.key)! + 1)
          if (reference.replacement === undefined) append(reference.value)
          else {
            if (parent.nodeType !== 1 || (parent as Element).namespaceURI !== W || (parent as Element).localName !== "r")
              throw new Error("Markdown code text requires a Word run.")
            for (const part of reference.value.split(/([\t\n])/)) {
              if (part === "\t" || part === "\n") parent.insertBefore(document.createElementNS(W,
                `${text.prefix || "w"}:${part === "\t" ? "tab" : "br"}`), text)
              else append(part)
            }
          }
          cursor = match.index! + match[0].length
        }
        append(value.slice(cursor))
        parent.removeChild(text)
      }
      const xml = new XMLSerializer().serializeToString(document)
      if (xml.includes(prefix) || [...counts.values()].some(count => count !== 1))
        throw new Error("Markdown render text did not preserve every exact literal text token.")
      parts["word/document.xml"] = strToU8(xml)
      return zipSync(parts, { mtime: new Date(1980, 0, 1) })
    },
  }
}
