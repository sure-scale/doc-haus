import path from "node:path"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { readFirmPreferences } from "./preferences"

const CORE = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
const DC = "http://purl.org/dc/elements/1.1/"
const APP = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
const REL = "http://schemas.openxmlformats.org/package/2006/relationships"
const CT = "http://schemas.openxmlformats.org/package/2006/content-types"

// XML 1.0 permits tabs, line breaks and Unicode scalar values in these ranges.
// Remove forbidden controls, lone surrogates and the excluded BMP noncharacters
// before names reach metadata, revisions or comments; XML escaping is separate.
function identityText(value: unknown) {
  return typeof value === "string"
    ? value.replace(/[^\u0009\u000A\u000D\u0020-\uD7FF\uE000-\uFFFD\u{10000}-\u{10FFFF}]/gu, "").trim()
    : ""
}

// Read only the identity fields; research credentials and other settings never
// enter document metadata. Explicit editing-tool authors take precedence.
export function draftingAttribution() {
  const preferences = readFirmPreferences()
  const attorney = identityText(preferences.attorney)
  const firm = identityText(preferences.firm)
  return { author: attorney || firm || "doc.haus", company: firm || "doc.haus" }
}

/** Stamp a newly generated file, never an existing source document. Only package
 * properties change: historical revision and comment authors stay untouched. */
export function withGeneratedDocxAttribution(bytes: Uint8Array, identity = draftingAttribution()) {
  const company = identityText(identity.company)
  const author = identityText(identity.author) || company || "doc.haus"
  const parts = unzipSync(bytes)
  const parse = (xml: string) => new DOMParser().parseFromString(xml, "application/xml")
  const relationships = parse(parts["_rels/.rels"] ? strFromU8(parts["_rels/.rels"]) : `<Relationships xmlns="${REL}"/>`)
  const contentTypes = parse(strFromU8(parts["[Content_Types].xml"]))
  const properties = [
    { file: "docProps/core.xml", namespace: CORE, root: "cp:coreProperties", relationship: `${REL}/metadata/core-properties`, type: "application/vnd.openxmlformats-package.core-properties+xml",
      values: [{ namespace: DC, name: "dc:creator", value: author }, { namespace: CORE, name: "cp:lastModifiedBy", value: author }] },
    { file: "docProps/app.xml", namespace: APP, root: "Properties", relationship: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties", type: "application/vnd.openxmlformats-officedocument.extended-properties+xml",
      values: [{ namespace: APP, name: "Company", value: company || "doc.haus" }] },
  ]
  for (const property of properties) {
    const existing = Array.from(relationships.getElementsByTagNameNS(REL, "Relationship")).find(row => row.getAttribute("Type") === property.relationship)
    const file = existing ? path.posix.resolve("/", existing.getAttribute("Target")!).slice(1) : property.file
    const document = parse(parts[file] ? strFromU8(parts[file]) : `<${property.root} xmlns${property.root.includes(":") ? ":cp" : ""}="${property.namespace}"/>`)
    for (const field of property.values) {
      const local = field.name.split(":").at(-1)!
      const element = document.getElementsByTagNameNS(field.namespace, local)[0] ?? document.createElementNS(field.namespace, field.name)
      element.textContent = field.value
      if (!element.parentNode) document.documentElement.appendChild(element)
    }
    parts[file] = strToU8(new XMLSerializer().serializeToString(document))
    if (!existing) {
      const relationship = relationships.createElementNS(REL, "Relationship")
      const ids = new Set(Array.from(relationships.getElementsByTagNameNS(REL, "Relationship")).map(row => row.getAttribute("Id")))
      let index = 1
      while (ids.has(`rId${index}`)) index += 1
      relationship.setAttribute("Id", `rId${index}`)
      relationship.setAttribute("Type", property.relationship)
      relationship.setAttribute("Target", file)
      relationships.documentElement.appendChild(relationship)
    }
    if (!Array.from(contentTypes.getElementsByTagNameNS(CT, "Override")).some(row => row.getAttribute("PartName") === `/${file}`)) {
      const override = contentTypes.createElementNS(CT, "Override")
      override.setAttribute("PartName", `/${file}`)
      override.setAttribute("ContentType", property.type)
      contentTypes.documentElement.appendChild(override)
    }
  }
  parts["_rels/.rels"] = strToU8(new XMLSerializer().serializeToString(relationships))
  parts["[Content_Types].xml"] = strToU8(new XMLSerializer().serializeToString(contentTypes))
  return zipSync(parts)
}
