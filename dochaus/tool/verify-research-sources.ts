import { tool } from "./legacy-definition"
import path from "node:path"
import { auditResearchFiles, renderResearchSourceAudit } from "../lib/research-source-audit"

export default tool({
  description:
    "Verify that source URLs in a research handoff match the claimed source titles and identifiers. Writes a compact research-source-audit.md file and returns PASS, PARTIAL, or BLOCKED.",
  args: {
    researchFile: tool.schema.string().optional().describe('Research handoff markdown file (default "research.md")'),
    outputFile: tool.schema.string().optional().describe('Audit markdown file to write (default "research-source-audit.md")'),
    maxFetches: tool.schema.number().min(1).max(50).optional().describe("Maximum distinct source URLs to verify"),
  },
  async execute(args, ctx) {
    const outputFile = safeMarkdownName(args.outputFile ?? "research-source-audit.md")
    const audit = await auditResearchFiles({
      directory: ctx.directory,
      researchFile: args.researchFile,
      maxFetches: args.maxFetches,
    })

    await Bun.write(path.join(ctx.directory, outputFile), renderResearchSourceAudit(audit))

    return {
      title: `Research source audit: ${audit.status}`,
      output: `${audit.status}: checked ${audit.checked} source(s), ${audit.failures.length} blocking failure(s), ${audit.warnings.length} warning(s). Wrote ${outputFile}.`,
      metadata: {
        status: audit.status,
        checked: audit.checked,
        failures: audit.failures.length,
        warnings: audit.warnings.length,
        outputFile,
      },
    }
  },
})

function safeMarkdownName(name: string) {
  const basename = path.basename(name.endsWith(".md") ? name : `${name}.md`)
  if (basename !== name && basename !== `${name}.md`) throw new Error("Pass only a markdown basename.")
  return basename
}
