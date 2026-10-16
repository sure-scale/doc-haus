import path from "node:path"

const reservedMarkdownBasenames = new Set([
  "draft-document",
  "read-document",
  "redline",
  "search-document",
  "tracked-changes",
  "word-integration",
  "write-matter-file",
])

// Agent-written artifacts are markdown; no mechanical reader consumes an
// agent-written JSON artifact.
export function matterArtifactName(
  input: string,
): { ok: true; name: string } | { ok: false; message: string } {
  if (input.endsWith(".json"))
    return { ok: false, message: `JSON artifact "${input}" is not agent-writable. Use a .md basename.` }
  const name = path.basename(input.endsWith(".md") ? input : `${input}.md`)
  if (!name || name === ".md") return { ok: false, message: "File name must include a non-empty basename." }
  if (name !== input && name !== `${input}.md`)
    return { ok: false, message: `Refusing path-like name "${input}". Pass only a markdown basename.` }
  if (reservedMarkdownBasenames.has(name.replace(/\.md$/, "")))
    return {
      ok: false,
      message: `Refusing tool-name file "${input}". Pass the intended markdown artifact filename, not a tool name.`,
    }
  return { ok: true, name }
}
