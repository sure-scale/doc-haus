import { Parser } from "htmlparser2"

export const UNTRUSTED_HTML_MAX_BYTES = 5 * 1024 * 1024

type VisibilityWarning =
  | "stylesheet_visibility_unresolved"
  | "inline_visibility_unresolved"
  | "interactive_visibility_unresolved"

const NON_RENDERED = new Set(["head", "script", "style", "template"])
const BLOCKS = new Set([
  "address", "article", "aside", "blockquote", "br", "caption", "dd", "div", "dl", "dt", "fieldset",
  "figcaption", "figure", "footer", "form", "h1", "h2", "h3", "h4", "h5", "h6", "header", "hr",
  "li", "main", "nav", "ol", "p", "pre", "section", "table", "td", "th", "tr", "ul",
])

// This is a static projection, not a browser/CSS visibility certificate. It does
// not execute scripts, resolve entities from a DTD, fetch styles, or render assets.
// Callers retain the original bytes and record warnings alongside this exact text.
export function projectUntrustedHTML(html: string) {
  if (Buffer.byteLength(html, "utf8") > UNTRUSTED_HTML_MAX_BYTES)
    throw new Error("HTML visibility projection exceeds its byte limit.")
  const warnings = new Set<VisibilityWarning>()
  const stack: { omitted: boolean; invisible: boolean; pre: boolean; block: boolean }[] = []
  const parts: { text: string; boundary?: true }[] = []
  let hiddenElements = 0
  let hiddenTextChars = 0
  const parser = new Parser({
    onopentag(name, attributes) {
      if (stack.length >= 512) throw new Error("HTML visibility projection exceeds its nesting limit.")
      const parent = stack.at(-1)
      const style = visibilityStyle(attributes.style, warnings)
      if (name === "style" || name === "link" && /(?:^|\s)stylesheet(?:\s|$)/i.test(attributes.rel ?? ""))
        warnings.add("stylesheet_visibility_unresolved")
      // Collapsed or browser-dependent fallback text is retained, not silently
      // mistaken for absent legal content. Its visibility remains unproven.
      if (name === "details" && !("open" in attributes) || ["noscript", "iframe", "object", "embed", "canvas", "input", "textarea", "select"].includes(name))
        warnings.add("interactive_visibility_unresolved")
      if (attributes.hidden?.toLowerCase() === "until-found") warnings.add("interactive_visibility_unresolved")
      const hidden = "hidden" in attributes && style.display === undefined
      if ("hidden" in attributes && style.display !== undefined && style.display !== "none")
        warnings.add("inline_visibility_unresolved")
      const omitted = Boolean(parent?.omitted || NON_RENDERED.has(name) || hidden ||
        style.display === "none" || style["content-visibility"] === "hidden" ||
        style.opacity !== undefined && /^[+-]?(?:\d+\.?\d*|\.\d+)%?$/.test(style.opacity) && Number(style.opacity.replace(/%$/, "")) <= 0)
      // Unlike display/opacity, inherited visibility can be restored by a child.
      const invisible = style.visibility === "visible" ? false
        : ["hidden", "collapse"].includes(style.visibility ?? "") ? true : parent?.invisible ?? false
      const block = BLOCKS.has(name)
      if (!omitted && block) parts.push({ text: "\n", boundary: true })
      if (!parent?.omitted && (omitted || invisible && !parent?.invisible)) hiddenElements++
      stack.push({ omitted, invisible, pre: name === "pre" || parent?.pre === true, block })
    },
    ontext(text) {
      const current = stack.at(-1)
      if (current?.omitted || current?.invisible) {
        hiddenTextChars += text.length
        return
      }
      parts.push({ text: current?.pre ? text.replace(/\r\n?/g, "\n") : text.replace(/[\t\n\r\f ]+/g, " ") })
    },
    onclosetag() {
      const current = stack.pop()
      if (current?.block && !current.omitted) parts.push({ text: "\n", boundary: true })
    },
  }, { decodeEntities: true })
  parser.end(html)
  return {
    text: parts.filter((part, index) => !part.boundary ||
      index > 0 && index < parts.length - 1 && !parts[index - 1]!.boundary).map((part) => part.text).join(""),
    projection: {
      method: "static-html-visible-text" as const,
      version: 1 as const,
      hiddenElements,
      hiddenTextChars,
      warnings: [...warnings].sort(),
    },
  }
}

function visibilityStyle(raw: string | undefined, warnings: Set<VisibilityWarning>) {
  const values: Record<string, string> = {}
  if (!raw) return values
  // Complex CSS cannot be evaluated safely with a declaration splitter. Preserve
  // its text and surface uncertainty rather than guessing what a browser hid.
  if (/[\\()'"/]/.test(raw)) {
    warnings.add("inline_visibility_unresolved")
    return values
  }
  const important = new Set<string>()
  for (const declaration of raw.split(";")) {
    if (!declaration.trim()) continue
    const parsed = /^\s*([a-z-]+)\s*:\s*([^:!]+?)\s*(!important\s*)?$/i.exec(declaration)
    if (!parsed) {
      warnings.add("inline_visibility_unresolved")
      continue
    }
    const property = parsed[1]!.toLowerCase()
    const value = parsed[2]!.trim().toLowerCase()
    if (!value) {
      warnings.add("inline_visibility_unresolved")
      continue
    }
    // Typography can affect what is visible; declarations outside this small
    // non-layout evaluator remain explicit uncertainty, not silently certified.
    if (!["display", "visibility", "opacity", "content-visibility"].includes(property)) {
      warnings.add("inline_visibility_unresolved")
      continue
    }
    if (important.has(property) && !parsed[3]) continue
    if (parsed[3]) important.add(property)
    values[property] = value
    if (["inherit", "revert", "revert-layer", "unset"].includes(value)) warnings.add("inline_visibility_unresolved")
  }
  return values
}
