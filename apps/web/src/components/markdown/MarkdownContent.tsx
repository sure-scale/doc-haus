import ReactMarkdown from "react-markdown"

type MarkdownBlock =
  | { kind: "markdown"; markdown: string }
  | { kind: "table"; header: string[]; rows: string[][] }

export default function MarkdownContent({ children }: { children: string }) {
  return (
    <>
      {splitMarkdownTableBlocks(children).map((block, index) =>
        block.kind === "markdown" ? (
          <ReactMarkdown key={`markdown-${index}`}>{block.markdown}</ReactMarkdown>
        ) : (
          <div key={`table-${index}`} className="markdown-table-scroll">
            <table className="markdown-table">
              <thead>
                <tr>
                  {block.header.map((cell, cellIndex) => (
                    <th key={cellIndex}>
                      <ReactMarkdown>{cell}</ReactMarkdown>
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {block.rows.map((row, rowIndex) => (
                  <tr key={rowIndex}>
                    {block.header.map((_, cellIndex) => (
                      <td key={cellIndex}>
                        <ReactMarkdown>{row[cellIndex] ?? ""}</ReactMarkdown>
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        ),
      )}
    </>
  )
}

export function splitMarkdownTableBlocks(markdown: string): MarkdownBlock[] {
  const blocks: MarkdownBlock[] = []
  const pending: string[] = []
  const lines = markdown.replace(/\r\n/g, "\n").split("\n")

  for (let index = 0; index < lines.length; index += 1) {
    if (!isMarkdownTableStart(lines, index)) {
      pending.push(lines[index])
      continue
    }

    appendMarkdownBlock(blocks, pending)
    const rows: string[][] = []
    const header = splitMarkdownTableRow(lines[index])
    for (index += 2; index < lines.length && isMarkdownTableRow(lines[index]); index += 1) {
      rows.push(splitMarkdownTableRow(lines[index]))
    }
    blocks.push({ kind: "table", header, rows })
    index -= 1
  }

  appendMarkdownBlock(blocks, pending)
  return blocks
}

function appendMarkdownBlock(blocks: MarkdownBlock[], pending: string[]) {
  const markdown = pending.join("\n").trim()
  pending.length = 0
  if (markdown) blocks.push({ kind: "markdown", markdown })
}

function isMarkdownTableStart(lines: string[], index: number) {
  return isMarkdownTableRow(lines[index]) && isMarkdownTableSeparator(lines[index + 1] ?? "")
}

function isMarkdownTableRow(line: string) {
  const trimmed = line.trim()
  return trimmed.startsWith("|") && trimmed.endsWith("|") && splitMarkdownTableRow(trimmed).length > 1
}

function isMarkdownTableSeparator(line: string) {
  const cells = splitMarkdownTableRow(line)
  return cells.length > 1 && cells.every((cell) => /^:?-{3,}:?$/.test(cell.trim()))
}

function splitMarkdownTableRow(line: string) {
  const cells: string[] = []
  let cell = ""
  let inCode = false
  const trimmed = line.trim().replace(/^\|/, "").replace(/\|$/, "")

  for (let index = 0; index < trimmed.length; index += 1) {
    const char = trimmed[index]
    if (char === "`") inCode = !inCode
    if (char === "|" && !inCode && trimmed[index - 1] !== "\\") {
      cells.push(cell.trim())
      cell = ""
      continue
    }
    cell += char
  }

  cells.push(cell.trim())
  return cells
}
