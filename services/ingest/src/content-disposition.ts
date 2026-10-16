/** HTTP headers need an ASCII fallback; filename* preserves the full UTF-8 name. */
export function inlineDocumentDisposition(name: string) {
  const fallback = name.replace(/[^\x20-\x7e]|["\\]/g, "_")
  const encoded = encodeURIComponent(name).replace(/['()*]/g, value => `%${value.charCodeAt(0).toString(16).toUpperCase()}`)
  return `inline; filename="${fallback}"; filename*=UTF-8''${encoded}`
}
