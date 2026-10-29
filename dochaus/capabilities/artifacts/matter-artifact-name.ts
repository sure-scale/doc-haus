import path from "node:path"
import { safeFileName } from "../../lib/safe-file-name"

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
  // A name every platform can hold; the write reports the name it used.
  const stem = safeFileName(name.slice(0, -".md".length))
  if (!stem) return { ok: false, message: "File name must include a non-empty basename." }
  if (stem.toLowerCase() === "chronology")
    return { ok: false, message: "Chronology is managed in the matter's Chronology workbench. Use read-chronology or propose-chronology-changes; existing chronology.md source files are preserved." }
  if (reservedMarkdownBasenames.has(stem))
    return {
      ok: false,
      message: `Refusing tool-name file "${input}". Pass the intended markdown artifact filename, not a tool name.`,
    }
  return { ok: true, name: `${stem}.md` }
}
