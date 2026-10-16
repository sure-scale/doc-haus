import { realpathSync } from "node:fs"
import path from "node:path"
import { readRegularFile } from "../lib/regular-file"
import { UNTRUSTED_RECORD_NOTICE } from "../lib/untrusted"
import { tool } from "./legacy-definition"

// This is a view of the current record, not a saved projection or read session.
// Bound the complete rendered response, including escaped text and navigation.
const MAX_OUTPUT_CHARS = 6000
// Bun supports source-token JSON parsing; the current TS lib omits these APIs.
const sourceJSON = JSON as typeof JSON & {
  rawJSON(source: string): object
  isRawJSON(value: unknown): boolean
}

export default tool({
  description:
    "Read current matter-relative JSON data at a JSON Pointer. Small selected values remain complete; large objects/arrays return navigable direct children, and strings return exact decoded Unicode code-point pages. Follow next_offset exactly until null; text is never normalized or joined with added separators. Pointer coordinates are not stable IDs, and the record must remain unchanged while paging. No files are written and no semantic conclusions are supplied.",
  args: {
    path: tool.schema.string().min(1).describe("Path of the existing JSON record, relative to this matter."),
    pointer: tool.schema.string().optional().describe("JSON Pointer; omit or use an empty string for the root. Only own object properties and zero-based array indices are addressable."),
    offset: tool.schema.number().int().min(0).optional().describe("Zero-based Unicode code-point offset for strings, or direct-child offset for objects/arrays; default 0. Continue with the returned next_offset exactly."),
    limit: tool.schema.number().int().min(1).max(4000).optional().describe("Maximum text code points or direct children in a paged response (default 4000). The rendered output bound can return fewer. Complete small objects/arrays are returned whole at offset 0."),
  },
  async execute(args, ctx) {
    const content = recordBytes(ctx.directory, args.path)
    if (typeof content === "string") return content
    let record: unknown
    try {
      // Keep numeric JSON tokens verbatim, including large identifiers that a
      // JavaScript number would round. Raw JSON values remain scalar lookups.
      record = JSON.parse(new TextDecoder("utf-8", { fatal: true }).decode(content), (_key: string, value: unknown, context?: { source: string }) =>
        typeof value === "number" && context ? sourceJSON.rawJSON(context.source) : value,
      )
    } catch {
      return failure("the selected file is not valid JSON with UTF-8 text. Check the current record before retrying.")
    }
    const pointer = args.pointer ?? ""
    const selected = lookup(record, pointer)
    if (!selected.ok) return failure(selected.reason)
    const value = selected.value
    const type = valueType(value)
    const offset = args.offset ?? 0
    const limit = args.limit ?? 4000
    const location = { path: args.path, pointer, type }
    const render = (page: Record<string, unknown>) => JSON.stringify({ ...location, ...page, notice: UNTRUSTED_RECORD_NOTICE })
    const result = (page: Record<string, unknown>) => ({
      title: "Read JSON record",
      output: render(page),
      // Like read-document, the tool owns the complete bound and continuation.
      metadata: { path: args.path, pointer, page: { kind: page.kind, offset: page.offset ?? 0, returned: page.returned ?? null, total: page.total ?? null, next_offset: page.next_offset }, truncated: page.next_offset !== null },
    })
    if (typeof value === "string") {
      const characters = Array.from(value)
      if (offset > characters.length) return failure(`offset ${offset} is beyond the selected string's ${characters.length} Unicode code points. Use the preceding next_offset.`)
      const textPage = (count: number) => ({ kind: "text", unit: "unicode_code_points", offset, returned: count, total: characters.length, next_offset: offset + count < characters.length ? offset + count : null, text: characters.slice(offset, offset + count).join("") })
      if (render(textPage(0)).length > MAX_OUTPUT_CHARS) return failure("the path and pointer are too long for a bounded response. Select a shorter parent coordinate.")
      let low = 0
      let high = Math.min(limit, characters.length - offset)
      while (low < high) {
        const count = Math.ceil((low + high) / 2)
        if (render(textPage(count)).length <= MAX_OUTPUT_CHARS) low = count
        else high = count - 1
      }
      if (low === 0 && offset < characters.length) return failure("the path and pointer leave no room for text in a bounded response. Select a shorter parent coordinate.")
      return result(textPage(low))
    }
    const complete = { kind: "value", value, next_offset: null }
    if (offset === 0 && render(complete).length <= MAX_OUTPUT_CHARS) return result(complete)
    if (value === null || typeof value !== "object" || sourceJSON.isRawJSON(value)) return failure("offset applies only to text or direct-child pages. Read this scalar at offset 0.")
    const keys = Object.keys(value)
    if (offset > keys.length) return failure(`offset ${offset} is beyond the selected ${type}'s ${keys.length} direct children. Use the preceding next_offset.`)
    const children: Record<string, unknown>[] = []
    const childPage = () => ({ kind: "children", unit: "direct_children", offset, returned: children.length, total: keys.length, next_offset: offset + children.length < keys.length ? offset + children.length : null, children })
    if (render(childPage()).length > MAX_OUTPUT_CHARS) return failure("the path and pointer are too long for a bounded response. Select a shorter parent coordinate.")
    for (const key of keys.slice(offset, offset + limit)) {
      const child = (value as Record<string, unknown>)[key]
      const childPointer = `${pointer}/${key.replaceAll("~", "~0").replaceAll("/", "~1")}`
      children.push(childDescription(child, childPointer))
      if (render(childPage()).length <= MAX_OUTPUT_CHARS) continue
      children.pop()
      break
    }
    if (!children.length && offset < keys.length) return failure("the next child pointer is too long for a bounded response. Its full coordinate cannot be displayed without clipping.")
    return result(childPage())
  },
})

function recordBytes(directory: string, artifact: string): Buffer | string {
  // Reuse the matter containment pattern and the regular-file no-follow reader.
  // The canonical parent must be inside this matter, even through a symlink.
  if (path.isAbsolute(artifact)) return failure("path must be relative to this matter.")
  const target = path.resolve(directory, artifact)
  const relative = path.relative(path.resolve(directory), target)
  if (!relative || relative === ".." || relative.startsWith(`..${path.sep}`) || path.isAbsolute(relative)) return failure("path resolves outside this matter.")
  try {
    const root = realpathSync.native(directory)
    const parent = path.relative(root, realpathSync.native(path.dirname(target)))
    if (parent === ".." || parent.startsWith(`..${path.sep}`) || path.isAbsolute(parent)) return failure("the record's parent resolves outside this matter.")
    return readRegularFile(target) ?? failure("record not found. Check its matter-relative path.")
  } catch {
    return failure("record not found or unavailable as a safe regular file within this matter. Check its path and parent directory.")
  }
}

function lookup(record: unknown, pointer: string): { ok: true; value: unknown } | { ok: false; reason: string } {
  if (pointer === "") return { ok: true, value: record }
  if (!pointer.startsWith("/") || /~(?:[^01]|$)/.test(pointer)) return { ok: false, reason: "invalid JSON Pointer. Use an empty root pointer or /-separated property names, escaping ~ as ~0 and / as ~1." }
  let value = record
  for (const encoded of pointer.slice(1).split("/")) {
    const key = encoded.replaceAll("~1", "/").replaceAll("~0", "~")
    if (value === null || typeof value !== "object" || sourceJSON.isRawJSON(value)) return { ok: false, reason: "the pointer traverses a scalar. Select its parent to inspect the current value." }
    if (Array.isArray(value) && !/^(0|[1-9]\d*)$/.test(key)) return { ok: false, reason: "array coordinates must be zero-based indices without leading zeros; '-' is not a readable item." }
    if (!Object.hasOwn(value, key)) return { ok: false, reason: "pointer not found in the current record. Read its parent to obtain current child pointers." }
    value = (value as Record<string, unknown>)[key]
  }
  return { ok: true, value }
}

function valueType(value: unknown) {
  if (value === null) return "null"
  if (sourceJSON.isRawJSON(value)) return "number"
  if (Array.isArray(value)) return "array"
  return typeof value
}

function childDescription(value: unknown, pointer: string) {
  const descriptor: Record<string, unknown> = { pointer, type: valueType(value) }
  if (typeof value === "string") descriptor.total = Array.from(value).length
  if (value !== null && typeof value === "object" && !sourceJSON.isRawJSON(value)) {
    descriptor.total = Object.keys(value).length
    if (!Array.isArray(value) && Object.hasOwn(value, "id")) {
      const id = (value as Record<string, unknown>).id
      // Identity is an existing value, never a synthesized semantic label.
      if (id === null || sourceJSON.isRawJSON(id) || ["string", "number", "boolean"].includes(typeof id)) descriptor.identity = JSON.stringify(id).length <= 256 ? { pointer: `${pointer}/id`, value: id } : { pointer: `${pointer}/id`, type: valueType(id) }
    }
  }
  return descriptor
}

function failure(reason: string) {
  return `Could not read JSON record: ${reason}`
}
