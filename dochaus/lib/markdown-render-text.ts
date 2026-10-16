import { randomUUID } from "node:crypto"
import { parse, postprocess, preprocess } from "micromark"
import { decodeString } from "micromark-util-decode-string"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"

// Decode only CommonMark text tokens, after both the native Markdown parser
// and table polish have finished interpreting grammar. A decoded asterisk or
// pipe is literal text, never a new emphasis delimiter or table column.
export function prepareMarkdownRenderText(content: string) {
  const prefix = `DocHausText${randomUUID().replaceAll("-", "")}Z`
  if (content.includes(prefix)) throw new Error("Markdown render text token collision.")
  const scopes: string[] = []
  const references = postprocess(parse().document().write(preprocess()(content, undefined, true)))
    .flatMap(([kind, token]) => {
      if (kind === "enter") {
        scopes.push(token.type)
        return []
      }
      const text = !scopes.some(scope => [
        "codeFencedFenceInfo", "codeFencedFenceMeta", "resourceDestinationString",
        "resourceTitleString", "definitionDestinationString", "definitionTitleString",
      ].includes(scope))
      scopes.pop()
      return token.type === "characterReference" && text
        ? [{ start: token.start.offset!, end: token.end.offset!,
            value: decodeString(content.slice(token.start.offset!, token.end.offset!)) }]
        : []
    }).map((reference, index) => ({ ...reference, key: `${prefix}${index}Z` }))
  return {
    markdown: references.toReversed().reduce((text, reference) =>
      text.slice(0, reference.start) + reference.key + text.slice(reference.end), content),
    restore(bytes: Uint8Array) {
      if (!references.length) return bytes
      const parts = unzipSync(bytes)
      if (!parts["word/document.xml"])
        throw new Error("Markdown render text requires word/document.xml.")
      const counts = new Map(references.map(reference => [reference.key, 0]))
      const xml = strFromU8(parts["word/document.xml"])
        .replace(/(<w:t\b[^>]*>)([\s\S]*?)(<\/w:t>)/g, (_, open: string, text: string, close: string) =>
          open + references.reduce((value, reference) => value.replaceAll(reference.key, () => {
            counts.set(reference.key, counts.get(reference.key)! + 1)
            return reference.value.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;")
          }), text) + close)
      if (xml.includes(prefix) || [...counts.values()].some(count => count !== 1))
        throw new Error("Markdown render text did not preserve every exact character reference token.")
      parts["word/document.xml"] = strToU8(xml)
      return zipSync(parts)
    },
  }
}
