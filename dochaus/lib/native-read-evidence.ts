// Adapt the native reader's verified line range to the legal policy's stable
// read-boundary representation. Directory/media output provides no text EOF.
export function nativeReadEvidence(input: Record<string, unknown>, output: string, metadata?: Record<string, unknown>) {
  const filePath = typeof input.path === "string" ? input.path : input.filePath
  const header = output.split("\n", 1)[0] ?? ""
  const range = header.match(/^Read file .+, lines (\d+)-(\d+)$/)
  const empty = /^Read file .+, 0 lines$/.test(header)
  const normalized = { ...input, ...(typeof filePath === "string" ? { filePath } : {}) }
  if (!range && !empty) return { input: normalized, output }
  const start = range ? Number(range[1]) : 1
  const end = range ? Number(range[2]) : 0
  const lines = output.split("\n").slice(1)
  const contiguous = end >= start && lines.slice(0, end - start + 1).length === end - start + 1 &&
    lines.slice(0, end - start + 1).every((line, index) => line.startsWith(`${start + index}: `))
  const next = output.match(/\[Output truncated\. Continue reading with offset: (\d+)\](?:\n\(Showing lines \d+-\d+\. Use offset=\d+ to continue\.\))?$/)?.[1]
  const clippedLine = lines.some(line => /\.\.\. \(line truncated to \d+ chars\)$/.test(line))
  const preview = output.match(/\(Tool preview showing through line (\d+)\. Use offset=(\d+) to continue\.\)$/)
  const previewEnd = preview ? Number(preview[1]) : undefined
  const completePreview = previewEnd !== undefined && previewEnd >= start && previewEnd <= end &&
    Number(preview![2]) === previewEnd + 1 &&
    lines.slice(0, previewEnd - start + 1).length === previewEnd - start + 1 &&
    lines.slice(0, previewEnd - start + 1).every((line, index) => line.startsWith(`${start + index}: `))
  const boundary = metadata?.truncated === false && !clippedLine && (empty || contiguous)
    ? `\n(End of file - total ${end} lines)`
    : metadata?.truncated === true && contiguous && next && Number(next) === end + 1
      ? `\n(Showing lines ${start}-${end}. Use offset=${next} to continue.)`
      : ""
  const verified = !clippedLine && (boundary || completePreview && typeof metadata?.truncated === "boolean")
  return {
    input: verified ? { ...normalized, offset: start } : normalized,
    output: boundary && !output.endsWith(boundary) ? output + boundary : output,
    ...(verified ? { range: { start, next: completePreview ? previewEnd! + 1 : next ? Number(next) : null } } : {}),
  }
}
