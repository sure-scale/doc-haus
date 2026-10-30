export function parseFormatting(value: string): Record<string, unknown> {
  try {
    const parsed: unknown = JSON.parse(value)
    return parsed && typeof parsed === "object" && !Array.isArray(parsed) ? parsed as Record<string, unknown> : {}
  } catch { return {} }
}

export function formattingDescription(format: Record<string, unknown>) {
  const parts: string[] = []
  if (typeof format.style === "string") parts.push(`Style: ${format.style.replace(/^Heading(\d)$/, "Heading $1")}`)
  if (typeof format.spacingBefore === "number") parts.push(`Spacing before: ${format.spacingBefore} pt`)
  if (typeof format.spacingAfter === "number") parts.push(`Spacing after: ${format.spacingAfter} pt`)
  if (typeof format.lineSpacing === "number") parts.push(`Line spacing: ${format.lineSpacing}`)
  return parts.length ? `Paragraph formatting — ${parts.join("; ")}` : "Paragraph formatting could not be displayed."
}
