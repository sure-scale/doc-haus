// Display budget for tool results in model context. A result over the budget is
// replaced by its prefix plus a preview notice; a clipped read keeps an exact
// continuation (or EOF) boundary so the next read resumes without repeating text.
export function boundLargeToolOutput(
  input: { tool: string },
  output: { output?: unknown },
) {
  if (typeof output.output !== "string") return
  // These readers bound their canonical data and return an exact cursor.
  // Reclipping would discard text while advertising a later continuation, and
  // could remove the untrusted-data notice or an injection warning.
  if (["read-document", "read-record", "read-template-source", "compare-versions"].includes(input.tool)) return
  const max = toolOutputDisplayChars()
  if (output.output.length <= max) return
  const preview = output.output.slice(0, max)
  const readBoundary = input.tool === "read"
    ? output.output.match(
        /\((?:(?:Output capped at|Showing lines)[^\n]*Use offset=\d+ to continue\.|End of file - total \d+ lines)\)/,
      )?.[0].trim()
    : undefined
  const boundaryVisible = readBoundary && output.output.indexOf(readBoundary) + readBoundary.length <= max
    ? readBoundary
    : undefined
  const lastVisibleLine = input.tool === "read" && !boundaryVisible
    ? [...preview.matchAll(/(?:^|\n)(\d+):[^\n]*(?=\n)/g)].at(-1)?.[1]
    : undefined
  const visibleBoundary = lastVisibleLine
    ? `(Tool preview showing through line ${lastVisibleLine}. Use offset=${Number(lastVisibleLine) + 1} to continue.)`
    : boundaryVisible
  output.output =
    preview.trimEnd() +
    `\n\n[tool-output-preview] ${input.tool} output clipped at ${max} characters for context budget. ` +
    (visibleBoundary
      ? visibleBoundary.includes("Use offset=")
        ? `The next read of this file must use the exact offset preserved below. ` +
          `Changing only the limit repeats discarded prefix text and will be rejected.\n\n${visibleBoundary}`
        : `The read reached EOF; do not reread this file.\n\n${visibleBoundary}`
      : `Use a narrower query, a smaller file range, source-index.json, source-map.md, or a specific source-chunk file for the next step; do not repeat the same broad call.`)
}

function toolOutputDisplayChars() {
  const value = Number(process.env.DOCHAUS_TOOL_OUTPUT_CHARS ?? "6000")
  if (!Number.isFinite(value) || value < 1000) return 6000
  return Math.floor(value)
}
