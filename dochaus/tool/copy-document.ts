import { tool } from "./legacy-definition"
import { hasTrackedRevisions } from "../lib/docx-amendments"
import { readFileSync } from "node:fs"
import path from "node:path"

const ingestUrl = process.env.INGEST_URL ?? `http://127.0.0.1:${process.env.DOC_HAUS_INGEST_PORT ?? 4501}`

export default tool({
  description: "Create a separate Word copy of an existing registered matter document, preserving the canonical file bytes, styles, native comments and native revisions. Never overwrites. App pending proposals are NOT copied or accepted. Use before editing when the lawyer requests a new copy; then direct tracked-changes/redline only at the returned new document name. A copy that holds tracked changes cannot be edited until they are accepted or rejected in Word; edits answering the other side's markup go on the uploaded markup itself (Request edits).",
  args: {
    document: tool.schema.string().describe("Existing Word document filename in this matter, without a path"),
    name: tool.schema.string().describe("New Word document filename ending in .docx, without a path"),
  },
  async execute(args, ctx) {
    if ([args.document, args.name].some(name => name !== name.trim() || /[\\/\0]/.test(name) || path.extname(name).toLowerCase() !== ".docx"))
      return "Use Word document filenames within this matter, without paths."
    const target = path.join(ctx.directory, args.name)
    await ctx.ask({
      permission: "copy-document",
      patterns: [target],
      metadata: { source: args.document, document: args.name, pendingProposals: "Excluded; only canonical Word bytes are copied" },
    })
    const matter = JSON.parse(readFileSync(path.join(ctx.directory, "matter.json"), "utf8")) as { id: string }
    const response = await fetch(`${ingestUrl}/matters/${encodeURIComponent(matter.id)}/documents/copy`, {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ source: args.document, name: args.name }),
      signal: ctx.abort,
    })
    if (!response.ok) return `Copy failed (${response.status}): ${await response.text()}`
    const copied = await response.json() as { name: string; docPath: string; pendingProposalsExcluded: number }
    const next = hasTrackedRevisions(await Bun.file(copied.docPath).bytes())
      ? `${copied.name} holds tracked changes, so it cannot be edited until they are accepted or rejected in Word. Edits answering the other side's markup go on the uploaded markup itself (Request edits).`
      : `Use ${copied.name} for subsequent edits.`
    return {
      title: `Created ${copied.name}`,
      output: `Created and registered ${copied.name} as a separate Word copy of ${args.document}. Original unchanged. Canonical Word bytes preserved exactly, including styles, native comments and native revisions. ${copied.pendingProposalsExcluded} pending app proposals excluded; they remain on the original. ${next}`,
      metadata: { document: copied.docPath, name: copied.name, source: args.document, pendingProposalsExcluded: copied.pendingProposalsExcluded },
    }
  },
})
