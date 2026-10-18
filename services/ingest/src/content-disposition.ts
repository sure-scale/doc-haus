/** HTTP headers need an ASCII fallback; filename* preserves the full UTF-8 name. */
export function inlineDocumentDisposition(name: string) {
  return documentDisposition("inline", name)
}

/** The same header for a file the browser saves rather than shows. */
export function attachmentDocumentDisposition(name: string) {
  return documentDisposition("attachment", name)
}

function documentDisposition(type: "inline" | "attachment", name: string) {
  const fallback = name.replace(/[^\x20-\x7e]|["\\]/g, "_")
  const encoded = encodeURIComponent(name).replace(/['()*]/g, value => `%${value.charCodeAt(0).toString(16).toUpperCase()}`)
  return `${type}; filename="${fallback}"; filename*=UTF-8''${encoded}`
}
