// Shared value-extraction helper. Used by lib/verification-loop.ts
// (value-grounding gate). Pure.

export function extractValues(text: string): string[] {
  const values: string[] = []
  for (const m of text.matchAll(/\$[\d,]+(?:\.\d+)?/g)) values.push(m[0])
  for (const m of text.matchAll(/\d+(?:\.\d+)?%/g)) values.push(m[0])
  return [...new Set(values)]
}
