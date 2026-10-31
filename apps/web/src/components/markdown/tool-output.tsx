import { decodeMarkupEntities, parseStructuredReportMarkdown, unwrapTaskResult } from "./structured-report"
import StructuredReportView from "./StructuredReportView"

const TOOL_OUTPUT_TAGS = {
  path: "Path",
  type: "Type",
  entries: "Entries",
  content: "Content",
  warning: "Warning",
  error: "Error",
} as const

type ToolOutputTag = keyof typeof TOOL_OUTPUT_TAGS

export type StructuredToolOutputField = {
  tag: string
  title: string
  text: string
}

export type StructuredToolOutput = {
  structured: boolean
  fields: StructuredToolOutputField[]
}

export function parseStructuredToolOutput(text: string): StructuredToolOutput {
  // Old transcripts sometimes stored a rendered/escaped read envelope. Decode
  // that transport only; entities in an ordinary file body remain file data.
  const decoded = /^\s*&lt;path&gt;/i.test(text) ? decodeMarkupEntities(text) : text
  const fields: StructuredToolOutputField[] = []
  const pattern = /<(path|type|entries|content|warning|error)\b[^>]*>/gi
  let cursor = 0
  let sawKnownTag = false
  let filePath = ""
  let match: RegExpExecArray | null
  while ((match = pattern.exec(decoded))) {
    const start = match.index
    const tag = match[1].toLowerCase()
    if (!isToolOutputTag(tag)) continue
    const bodyStart = start + match[0].length
    const closing = new RegExp(`</${tag}\\s*>`, "gi")
    closing.lastIndex = bodyStart
    let end = closing.exec(decoded)
    // The content body is opaque. A literal closing marker in a file must not
    // cause its later text to be parsed as another tool/report envelope.
    if (tag === "content" && end) {
      let next: RegExpExecArray | null
      while ((next = closing.exec(decoded))) end = next
    }
    if (!end && tag !== "content") continue
    appendToolOutputNote(fields, decoded.slice(cursor, start))
    cursor = end ? end.index + end[0].length : decoded.length
    pattern.lastIndex = cursor
    sawKnownTag = true

    const fieldText = decoded.slice(bodyStart, end?.index).trim()
    if (tag === "path") filePath = fieldText
    appendToolOutputField(fields, tag, fieldText, filePath)
  }
  if (!sawKnownTag) return { structured: false, fields: [] }
  appendTrailingToolOutput(fields, decoded.slice(cursor), filePath)
  return { structured: true, fields }
}

export function ToolOutput({ children, tool }: { children: string; tool?: string }) {
  if (tool === "task") {
    const text = unwrapTaskResult(children, true)
    // Mixed or clipped transport envelopes have no safe unwrapping boundary.
    // Keep them raw, including any surrounding warning or partial result.
    if (text === children && /(?:<|&lt;)\/?task(?:\s|>|&gt;|_result)/i.test(children)) return <pre className="step-output">{children}</pre>
    const report = parseStructuredReportMarkdown(text)
    return <div className="step-output step-output-report">
      {report.structured ? <StructuredReportView report={report} /> : <pre className="step-output">{text}</pre>}
    </div>
  }

  // Read-style fields contain literal source content, not agent reports.
  // Unknown tool origins retain the old field formatter without report guessing.
  if (tool && tool !== "read") return <pre className="step-output">{children}</pre>

  const output = parseStructuredToolOutput(children)
  if (!output.structured) return <pre className="step-output">{children}</pre>

  return (
    <div className="step-output tool-output-structured">
      {output.fields.map((field, index) => (
        <div key={`${field.tag}-${index}`} className={`tool-output-field tool-output-${field.tag}`}>
          <div className="tool-output-label">{field.title}</div>
          {field.tag === "path" ? (
            <code className="tool-output-path">{field.text}</code>
          ) : field.tag === "type" || field.tag === "lines" ? (
            <span className="tool-output-chip">{field.text}</span>
          ) : (
            <pre className="tool-output-pre">{field.text}</pre>
          )}
        </div>
      ))}
    </div>
  )
}

function isToolOutputTag(tag: string): tag is ToolOutputTag {
  return tag in TOOL_OUTPUT_TAGS
}

function appendToolOutputField(fields: StructuredToolOutputField[], tag: ToolOutputTag, text: string, filePath: string) {
  if (tag !== "content") {
    fields.push({ tag, title: TOOL_OUTPUT_TAGS[tag], text })
    return
  }

  const content = formatToolContent(text, filePath)
  if (content.lineRange) fields.push({ tag: "lines", title: "Lines", text: content.lineRange })
  if (content.text) fields.push({ tag, title: TOOL_OUTPUT_TAGS[tag], text: content.text })
  fields.push(...content.notes)
}

function appendTrailingToolOutput(fields: StructuredToolOutputField[], text: string, filePath: string) {
  const trimmed = text.trim()
  if (!trimmed) return

  const unclosed = trimmed.match(/^<([a-z][a-z0-9_-]*)\b[^>]*>([\s\S]*)$/i)
  if (unclosed) {
    const tag = unclosed[1].toLowerCase()
    if (isToolOutputTag(tag)) {
      appendToolOutputField(fields, tag, unclosed[2].trim(), filePath)
      return
    }
  }

  appendToolOutputNote(fields, trimmed)
}

function appendToolOutputNote(fields: StructuredToolOutputField[], text: string) {
  const trimmed = text.trim()
  if (!trimmed) return

  const preview = trimmed.match(/^\[tool-output-preview\]\s*([\s\S]*)$/)
  if (preview) {
    fields.push({ tag: "tool-output-preview", title: "Output Preview", text: preview[1].trim() })
    return
  }

  fields.push({ tag: "note", title: "Note", text: trimmed })
}

function formatToolContent(text: string, filePath: string) {
  const notes: StructuredToolOutputField[] = []
  const preview = text.match(/\n*\[tool-output-preview\]\s*([\s\S]*)$/)
  const withoutPreview = preview ? text.slice(0, preview.index).trimEnd() : text
  if (preview) notes.push({ tag: "tool-output-preview", title: "Output Preview", text: preview[1].trim() })

  const fileNote = withoutPreview.match(/\n*\((End of file - total \d+ lines)\)\s*$/)
  const body = (fileNote ? withoutPreview.slice(0, fileNote.index) : withoutPreview).trim()
  if (fileNote) notes.push({ tag: "file-note", title: "File Note", text: fileNote[1] })

  const unnumbered = stripLinePrefixes(body)
  const formatted = filePath.endsWith(".json") ? prettyJson(unnumbered) : unnumbered.trimEnd()
  return { lineRange: contentLineRange(body), text: formatted, notes }
}

function stripLinePrefixes(content: string) {
  const lines = content.split("\n")
  const nonEmpty = lines.filter((line) => line.trim())
  if (!nonEmpty.length || !nonEmpty.every((line) => /^\d+:\s?/.test(line))) return content
  return lines.map((line) => line.replace(/^\d+:\s?/, "")).join("\n").trimEnd()
}

function prettyJson(content: string) {
  try {
    return JSON.stringify(JSON.parse(content), null, 2)
  } catch {
    return content.trimEnd()
  }
}

function contentLineRange(content: string) {
  const lines = content
    .split("\n")
    .map((line) => line.match(/^(\d+):/))
    .filter((match): match is RegExpMatchArray => Boolean(match))
    .map((match) => Number(match[1]))

  if (!lines.length) return ""
  const first = lines[0]
  const last = lines.at(-1)!
  return first === last ? String(first) : `${first}-${last}`
}
