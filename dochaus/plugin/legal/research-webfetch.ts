import { projectUntrustedHTML } from "../../lib/untrusted-html"

// Keep HTML intact until the legal projection has handled explicit hidden
// content. Upstream's markdown/text conversion otherwise loses that distinction.
export function prepareLegalWebFetch(args: Record<string, unknown>) {
  args.format = "html"
}

export function projectLegalWebFetch(
  args: Record<string, unknown>,
  output: { title: string; output: string; metadata: Record<string, unknown> },
) {
  // The core tool binds the HTTP content type to this exact title shape. Do not
  // sniff literal angle brackets in plain text or trust a model-supplied MIME.
  const prefix = `${String(args.url)} (`
  const contentType = typeof output.title === "string" && output.title.startsWith(prefix) && output.title.endsWith(")")
    ? output.title.slice(prefix.length, -1).split(";")[0]!.trim().toLowerCase()
    : ""
  if (args.format !== "html" || !["text/html", "application/xhtml+xml"].includes(contentType)) {
    output.metadata = { ...output.metadata, sourceProjection: { kind: "unmodified_text", visibility: "unassessed" } }
    return
  }
  const projected = projectUntrustedHTML(output.output)
  const completeness = output.metadata.truncated === true ? "truncated" : "tool_returned_body"
  output.output = projected.text + (completeness === "truncated"
    ? "\n\n[Source excerpt only: the engine truncated this response before projection. Unseen content has not been assessed.]"
    : "")
  output.metadata = {
    ...output.metadata,
    sourceProjection: { ...projected.projection, completeness },
  }
  // A decoded tool body is not an original-byte snapshot. This projection never
  // manufactures quotation verification or the official-source fetch receipt.
}
