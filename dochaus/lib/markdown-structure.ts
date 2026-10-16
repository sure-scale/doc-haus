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

export function scanMarkdownStructure(content: string, targetNames: readonly string[]) {
  const lines = content.split(/\r?\n/)
  const targets = new Set(targetNames.map((name) => name.toLowerCase()))
  const fences: MarkdownFenceBlock[] = []
  const fencedContent: MarkdownFenceContent[] = []
  const indentedCodeLines: MarkdownVisibleLine[] = []
  const visibleLines: MarkdownVisibleLine[] = []
  // Every line a fence owns: its delimiters and its literal payload.
  const fenceLines: Array<{ line: number; delimiter: boolean }> = []
  let commentOpen = false
  let htmlBlock:
    | {
        close?: RegExp
        untilBlank?: true
      }
    | undefined
  let fence:
    | {
        marker: string
        name?: string
        info: string
        line: number
        index: number
        payload: string[]
      }
    | undefined

  lines.forEach((line, index) => {
    if (fence) {
      if (!isMarkdownFenceClose(line, fence.marker)) {
        fence.payload.push(line)
        fenceLines.push({ line: index + 1, delimiter: false })
        return
      }
      fenceLines.push({ line: index + 1, delimiter: true })
      fencedContent.push({
        payload: fence.payload.join("\n"),
        complete: true,
        line: fence.line,
      })
      if (fence.name)
        fences.push({
          name: fence.name,
          info: fence.info,
          payload: fence.payload.join("\n"),
          raw: lines.slice(fence.index, index + 1).join("\n"),
          complete: true,
          line: fence.line,
        })
      fence = undefined
      return
    }

    if (!commentOpen && !htmlBlock && /^(?: {4}|\t)/.test(line)) {
      indentedCodeLines.push({
        line: index + 1,
        value: line.replace(/^(?: {4}|\t)/, ""),
        source: line,
      })
      return
    }
    const uncommented = maskHtmlComments(line, commentOpen)
    commentOpen = uncommented.commentOpen
    if (htmlBlock) {
      if (htmlBlock.untilBlank && /^[ \t]*$/.test(line)) {
        htmlBlock = undefined
        visibleLines.push({ line: index + 1, value: uncommented.value, source: line })
        return
      }
      if (htmlBlock.close?.test(uncommented.value)) htmlBlock = undefined
      return
    }
    const htmlOpening = markdownHtmlBlockOpen(uncommented.value)
    if (htmlOpening) {
      if (!htmlOpening.close?.test(uncommented.value)) htmlBlock = htmlOpening
      return
    }
    const opening = markdownFenceOpen(uncommented.value)
    if (opening) {
      const name = opening.indent || opening.marker !== "```"
        ? undefined
        : targetFenceName(opening.info, targets)
      fenceLines.push({ line: index + 1, delimiter: true })
      fence = {
        marker: opening.marker,
        name: name && targets.has(name) ? name : undefined,
        info: line.slice(opening.indent.length + opening.marker.length),
        line: index + 1,
        index,
        payload: [],
      }
      return
    }
    visibleLines.push({ line: index + 1, value: uncommented.value, source: line })
  })

  if (fence) {
    fencedContent.push({
      payload: fence.payload.join("\n"),
      complete: false,
      line: fence.line,
    })
    if (fence.name)
      fences.push({
        name: fence.name,
        info: fence.info,
        payload: fence.payload.join("\n"),
        raw: lines.slice(fence.index).join("\n"),
        complete: false,
        line: fence.line,
      })
  }
  return { fences, fencedContent, indentedCodeLines, visibleLines, fenceLines }
}

function markdownFenceOpen(line: string) {
  const match = /^( {0,3})(`{3,}|~{3,})(.*)$/.exec(line)
  if (!match) return
  if (match[2]![0] === "`" && match[3]!.includes("`")) return
  return { indent: match[1]!, marker: match[2]!, info: match[3]! }
}

function targetFenceName(info: string, targets: Set<string>) {
  const candidate = info.replace(/^\s+/, "").toLowerCase()
  return [...targets].find((target) => {
    if (!candidate.startsWith(target)) return false
    const suffix = candidate[target.length]
    return !suffix || !/[a-z0-9_-]/.test(suffix)
  })
}

function isMarkdownFenceClose(line: string, opening: string) {
  const match = /^( {0,3})(`+|~+)[ \t]*$/.exec(line)
  return Boolean(
    match &&
    match[2]![0] === opening[0] &&
    match[2]!.length >= opening.length,
  )
}

function markdownHtmlBlockOpen(line: string) {
  const closedTag = /^ {0,3}<(script|pre|style|textarea)(?:[ \t]|>|$)/i.exec(line)?.[1]
  if (closedTag) return { close: new RegExp(`</${closedTag}[ \\t]*>`, "i") }
  if (/^ {0,3}<\?/.test(line)) return { close: /\?>/ }
  if (/^ {0,3}<![A-Z]/.test(line)) return { close: />/ }
  if (/^ {0,3}<!\[CDATA\[/.test(line)) return { close: /\]\]>/ }
  if (
    /^ {0,3}<\/?(?:address|article|aside|base|basefont|blockquote|body|caption|center|col|colgroup|dd|details|dialog|dir|div|dl|dt|fieldset|figcaption|figure|footer|form|frame|frameset|h[1-6]|head|header|hr|html|iframe|legend|li|link|main|menu|menuitem|nav|noframes|ol|optgroup|option|p|param|search|section|summary|table|tbody|td|tfoot|th|thead|title|tr|track|ul)(?:[ \t]|\/?>|$)/i.test(line)
  )
    return { untilBlank: true as const }
  if (isCompleteMarkdownHtmlTag(line))
    return { untilBlank: true as const }
}

function isCompleteMarkdownHtmlTag(line: string) {
  const value = /^ {0,3}(\S.*)$/.exec(line)?.[1]?.trimEnd()
  if (!value || !/^<\/?[A-Za-z]/.test(value)) return false
  let quote: "'" | '"' | undefined
  for (const [index, character] of value.split("").entries()) {
    if (quote) {
      if (character === quote) quote = undefined
      continue
    }
    if (character === "'" || character === '"') {
      quote = character
      continue
    }
    if (character === ">") return index === value.length - 1
  }
  return false
}

function maskHtmlComments(line: string, initiallyOpen: boolean) {
  let commentOpen = initiallyOpen
  let cursor = 0
  let value = ""
  while (cursor < line.length) {
    if (commentOpen) {
      const end = line.indexOf("-->", cursor)
      if (end === -1)
        return {
          value: value + " ".repeat(line.length - cursor),
          commentOpen: true,
        }
      value += " ".repeat(end + 3 - cursor)
      cursor = end + 3
      commentOpen = false
      continue
    }
    const start = line.indexOf("<!--", cursor)
    if (start === -1)
      return {
        value: value + line.slice(cursor),
        commentOpen: false,
      }
    value += line.slice(cursor, start) + " ".repeat(4)
    cursor = start + 4
    commentOpen = true
  }
  return { value, commentOpen }
}
