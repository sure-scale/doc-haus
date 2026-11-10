import { parse, postprocess, preprocess } from "micromark"

export type MarkdownFenceBlock = {
  name: string
  info: string
  payload: string
  raw: string
  complete: boolean
  line: number
}

export type MarkdownVisibleLine = {
  line: number
  value: string
  source: string
}

export type MarkdownFenceContent = {
  payload: string
  complete: boolean
  line: number
}

/** Classify original lines by CommonMark token ownership. Container prefixes,
 * lazy continuations, fence identity and indentation belong to the maintained
 * document tokenizer, not an independent line-state machine. Source strings and
 * line numbers stay exact; no quoted/literal evidence is discarded here.
 */
export function scanMarkdownStructure(content: string, targetNames: readonly string[]) {
  const lines = content.split(/\r?\n/)
  const starts = [0, ...[...content.matchAll(/\r?\n/g)].map(match => match.index + match[0].length)]
  const events = postprocess(parse().document().write(preprocess()(content, undefined, true)))
  const entered = events.filter(([kind]) => kind === "enter").map(([, token]) => token)
  type Token = (typeof entered)[number]
  const fenceDelimiters = new Map<Token, Token[]>()
  let activeFence: Token | undefined
  for (const [kind, token] of events) {
    if (token.type === "codeFenced") {
      activeFence = kind === "enter" ? token : undefined
      if (activeFence) fenceDelimiters.set(activeFence, [])
    } else if (kind === "enter" && token.type === "codeFencedFence" && activeFence) fenceDelimiters.get(activeFence)!.push(token)
  }
  const lastLine = (token: Token) => token.end.line - (token.end.column === 1 ? 1 : 0)
  const kinds = new Map<number, "fence_delimiter" | "fenced" | "indented_code" | "html_block">()
  const mark = (token: Token, kind: NonNullable<ReturnType<typeof kinds.get>>) => {
    for (let line = token.start.line; line <= lastLine(token); line++) kinds.set(line, kind)
  }
  const fences: MarkdownFenceBlock[] = [], fencedContent: MarkdownFenceContent[] = []
  const targets = new Set(targetNames.map(name => name.toLowerCase()))
  for (const token of entered) {
    if (token.type === "codeFenced") {
      mark(token, "fenced")
      const delimiters = fenceDelimiters.get(token)!
      for (const delimiter of delimiters) mark(delimiter, "fence_delimiter")
      const opening = delimiters[0]!, closing = delimiters[1]
      const complete = Boolean(closing)
      const raw = content.slice(token.start.offset, token.end.offset).replace(/\r\n/g, "\n")
      const payload = raw.split("\n").slice(1, complete ? -1 : undefined).join("\n")
      fencedContent.push({ payload, complete, line: token.start.line })
      const rawOpening = content.slice(opening.start.offset, opening.end.offset)
      const marker = /^(`+|~+)/.exec(rawOpening)?.[1]!
      const info = rawOpening.slice(marker.length)
      // Explicit legacy envelope: only a column-one, three-backtick opener
      // admits a named block. Container examples cannot become task metadata.
      const name = token.start.column === 1 && marker === "```" ? targetFenceName(info, targets) : undefined
      if (name) fences.push({ name, info, payload, raw, complete, line: token.start.line })
    } else if (token.type === "codeIndented") mark(token, "indented_code")
    else if (token.type === "htmlFlow" && !content.slice(token.start.offset, token.end.offset).startsWith("<!--")) mark(token, "html_block")
  }

  // Comments affect ordinary visible-line values only when CommonMark actually
  // parsed HTML. A literal comment inside code cannot hide later prose.
  const masked = content.split("")
  for (const token of entered.filter(token => token.type === "htmlText" || token.type === "htmlFlow")) {
    const raw = content.slice(token.start.offset, token.end.offset)
    for (const comment of raw.matchAll(/<!--[\s\S]*?(?:-->|$)/g)) {
      for (let index = token.start.offset + comment.index; index < token.start.offset + comment.index + comment[0].length; index++) {
        if (masked[index] !== "\r" && masked[index] !== "\n") masked[index] = " "
      }
    }
  }
  const visible = masked.join("")
  const indentedCodeLines: MarkdownVisibleLine[] = [], visibleLines: MarkdownVisibleLine[] = []
  const fenceLines: Array<{ line: number; delimiter: boolean }> = []
  lines.forEach((source, index) => {
    const line = index + 1, kind = kinds.get(line)
    if (kind === "fenced" || kind === "fence_delimiter") fenceLines.push({ line, delimiter: kind === "fence_delimiter" })
    else if (kind === "indented_code") indentedCodeLines.push({ line, source, value: source.replace(/^(?: {4}|\t)/, "") })
    else if (kind !== "html_block") visibleLines.push({ line, source, value: visible.slice(starts[index], starts[index]! + source.length) })
  })
  return { fences, fencedContent, indentedCodeLines, visibleLines, fenceLines }
}

function targetFenceName(info: string, targets: Set<string>) {
  const candidate = info.replace(/^\s+/, "").toLowerCase()
  return [...targets].find(target => {
    if (!candidate.startsWith(target)) return false
    const suffix = candidate[target.length]
    return !suffix || !/[a-z0-9_-]/.test(suffix)
  })
}
