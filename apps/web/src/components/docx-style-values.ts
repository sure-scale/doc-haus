import { unzipSync, zipSync } from "fflate"

// The converter copies some values from a Word file into the page's stylesheet
// without escaping them: font names inside single quotes (font-family: 'Name',
// serif), colours bare (color: #RRGGBB; background: #RRGGBB) and style ids as
// selector text (p.docx-<styleId> {). A value carrying a quote, semicolon,
// brace or line break could end its string, declaration or selector and add
// CSS rules to the whole page: hide doc.haus's own flags, restyle the tracked
// changes, or fetch a remote image when the file is opened. So every viewer
// renders a copy in which those values keep only the characters they are made
// of. A font name keeps letters, marks and digits in any script, spaces and
// - _ . , + & @ ( ); a colour, letters and digits (hex digits or a named
// colour); a style id, letters, marks and digits in any script, - and _. The
// file, its download and its versions are untouched, and a file whose values
// are all ordinary renders from its own bytes.

// The elements that name a font, by local name, and the attributes holding the
// name: run fonts, the font table, symbol characters and DrawingML (theme and
// drawing) fonts. Matched under any prefix, as the converter reads namespaces.
const FONT_ATTRIBUTES: Record<string, readonly string[]> = {
  rFonts: ["ascii", "hAnsi", "eastAsia", "cs"],
  font: ["name", "typeface"],
  altName: ["val"],
  sym: ["font", "typeface"],
  latin: ["typeface"],
  ea: ["typeface"],
  cs: ["typeface"],
  buFont: ["typeface"],
}
// Colours: WordprocessingML's color and fill attributes on any element (runs,
// shading, borders, the page background), which are always prefixed; the val
// of a color or highlight element; and DrawingML colour values.
const COLOUR_ATTRIBUTES = ["color", "fill"]
const COLOUR_VALUES: Record<string, readonly string[]> = { color: ["val"], highlight: ["val"], srgbClr: ["val"], sysClr: ["lastClr"] }
// Style ids: a style's own id and every reference to one, cleaned alike so a
// paragraph, run, table or numbering level keeps the style it names.
const STYLE_ID_VALUES: Record<string, readonly string[]> = {
  style: ["styleId"],
  pStyle: ["val"],
  rStyle: ["val"],
  tblStyle: ["val"],
  basedOn: ["val"],
  next: ["val"],
  link: ["val"],
  numStyleLink: ["val"],
  styleLink: ["val"],
}
const NOT_FONT = /[^\p{L}\p{M}\p{N} \-_.,+&@()]/gu
const NOT_COLOUR = /[^0-9A-Za-z]/g
const NOT_STYLE_ID = /[^\p{L}\p{M}\p{N}\-_]/gu

// A start tag. A prefix or local name may use letters, digits and marks in any
// script, as the converter accepts them; quoted values are matched whole, as a
// value may hold a raw ">".
const TAG = /<([^\s:=/>!?<"']+(?::[^\s:=/><"']+)?)((?:\s+[^\s=/>]+\s*=\s*(?:"[^"]*"|'[^']*'))*)(\s*\/?>)/gu
const ATTRIBUTE = /([^\s=/>]+)(\s*=\s*)(?:"([^"]*)"|'([^']*)')/g
const ENTITY = /&(?:#(\d+)|#x([0-9a-fA-F]+)|(lt|gt|amp|quot|apos));/g
const NAMED: Record<string, string> = { lt: "<", gt: ">", amp: "&", quot: '"', apos: "'" }

function decode(value: string) {
  return value.replace(ENTITY, (whole, decimal?: string, hex?: string, name?: string) => {
    if (name) return NAMED[name]!
    const code = decimal ? Number(decimal) : parseInt(hex!, 16)
    return code <= 0x10ffff ? String.fromCodePoint(code) : whole
  })
}

function unsafeCharacters(element: string, key: string) {
  if (key.startsWith("xmlns")) return
  const prefixed = key.includes(":")
  const local = key.slice(key.indexOf(":") + 1)
  if (FONT_ATTRIBUTES[element]?.includes(local)) return NOT_FONT
  if (COLOUR_VALUES[element]?.includes(local) || (prefixed && COLOUR_ATTRIBUTES.includes(local))) return NOT_COLOUR
  if (STYLE_ID_VALUES[element]?.includes(local)) return NOT_STYLE_ID
}

// One part's text with its unsafe style values cleaned, or undefined when
// every value is already safe.
export function safeStyleValuesInXml(xml: string) {
  let changed = false
  const out = xml.replace(TAG, (tag, name: string, attributes: string, end: string) => {
    if (!attributes) return tag
    const element = name.slice(name.indexOf(":") + 1)
    const cleaned = attributes.replace(ATTRIBUTE, (attribute, key: string, equals: string, double?: string, single?: string) => {
      const unsafe = unsafeCharacters(element, key)
      if (!unsafe) return attribute
      const value = decode(double ?? single ?? "")
      const safe = value.replace(unsafe, "")
      if (safe === value) return attribute
      changed = true
      return `${key}${equals}"${safe.replaceAll("&", "&amp;")}"`
    })
    return `<${name}${cleaned}${end}`
  })
  return changed ? out : undefined
}

// A part is read as text when it decodes as XML; any part may carry these
// values, whatever its name, since Word finds parts by relationship. The decoder
// drops a byte-order mark, which a UTF-16 part gets back when it is rewritten.
// Media is told apart by its first byte, so it is never decoded.
function partText(bytes: Uint8Array) {
  if (bytes[0] === 0xff && bytes[1] === 0xfe) return { text: new TextDecoder("utf-16le").decode(bytes), encoding: "utf-16le" as const }
  if (bytes[0] === 0xfe && bytes[1] === 0xff) return { text: new TextDecoder("utf-16be").decode(bytes), encoding: "utf-16be" as const }
  let start = bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf ? 3 : 0
  while (bytes[start] === 0x20 || bytes[start] === 0x09 || bytes[start] === 0x0a || bytes[start] === 0x0d) start++
  return bytes[start] === 0x3c ? { text: new TextDecoder().decode(bytes), encoding: "utf-8" as const } : undefined
}

function encode(text: string, encoding: "utf-8" | "utf-16le" | "utf-16be") {
  if (encoding === "utf-8") return new TextEncoder().encode(text)
  text = "\ufeff" + text
  const out = new Uint8Array(text.length * 2)
  const view = new DataView(out.buffer)
  for (let i = 0; i < text.length; i++) view.setUint16(i * 2, text.charCodeAt(i), encoding === "utf-16le")
  return out
}

// The bytes a viewer renders: the file itself, or a copy with its unsafe style
// values cleaned. Bytes that are not a zip package go to the converter as they
// are, so it reports its own error.
export function withSafeStyleValues(bytes: Uint8Array): Uint8Array {
  let parts: Record<string, Uint8Array>
  try {
    parts = unzipSync(bytes)
  } catch {
    return bytes
  }
  let changed = false
  for (const [name, data] of Object.entries(parts)) {
    const part = partText(data)
    const safe = part && safeStyleValuesInXml(part.text)
    if (safe === undefined) continue
    parts[name] = encode(safe, part!.encoding)
    changed = true
  }
  return changed ? zipSync(parts) : bytes
}
