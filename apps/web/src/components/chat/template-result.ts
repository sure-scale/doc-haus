export type TemplateResult = { summary: string; placeholders: string[] }

// Format the template tools' existing output without dropping unfamiliar text.
// Only replace the exact placeholder suffix described by their metadata.
export function templateResult(tool: string, output: string, metadata?: Record<string, unknown>): TemplateResult | undefined {
  if (tool !== "create-template" && tool !== "update-template") return
  const placeholders = metadata?.placeholders
  if (!Array.isArray(placeholders) || !placeholders.every((value): value is string => typeof value === "string")) return
  if (placeholders.length === 0) return { summary: output, placeholders }
  const suffix = `Placeholders: ${placeholders.join(", ")}.`
  if (!output.endsWith(suffix)) return
  return { summary: output.slice(0, -suffix.length).trim(), placeholders }
}
