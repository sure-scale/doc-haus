import { inflateRawSync } from "node:zlib"
import { normalizeUntrustedText, type InjectionFinding } from "../../../dochaus/lib/untrusted-text"

export { detectInjection, type InjectionFinding } from "../../../dochaus/lib/untrusted-text"

// Keep the ingest API while sharing its exact normalization with research and
// live document extraction. Chunk offsets refer to this normalized text.
export function normalizeExtractedText(raw: string) {
  return normalizeUntrustedText(raw)
}

// Formatting tricks that hide text from the rendered page while leaving it in the
// extracted (and therefore indexed) text: vanish ("hidden") runs, white-on-white
// runs, and near-invisible font sizes (w:sz is half-points, so 4 = 2pt). Counts
// only — the hidden text itself is already in the normalized body, where the
// pattern rules above see it.
export function scanDocxHiddenContent(buffer: Buffer): InjectionFinding[] {
  const xml = readZipEntry(buffer, "word/document.xml")?.toString("utf8")
  if (!xml) return []
  const hidden = [...xml.matchAll(/<w:vanish\b(?![^>]*w:val="(?:false|0)")[^>]*\/?>/g)].length
  const white = [...xml.matchAll(/<w:color\b[^>]*w:val="FFFFFF"/gi)].length
  const tiny = [...xml.matchAll(/<w:sz\b[^>]*w:val="(\d+)"/g)].filter((m) => Number(m[1]) <= 4).length
  return [
    ...(hidden
      ? [
          {
            rule: "hidden-text",
            detail: `${hidden} run(s) of hidden ("vanish") text — invisible in Word but present in the indexed text`,
          },
        ]
      : []),
    ...(white ? [{ rule: "white-text", detail: `${white} run(s) of white-on-white text` }] : []),
    ...(tiny ? [{ rule: "tiny-text", detail: `${tiny} run(s) of near-invisible text (2pt or smaller)` }] : []),
  ]
}

// Minimal ZIP reader for one named entry: walk the central directory (whose sizes
// are reliable even when a local header deferred them to a data descriptor) to the
// entry's local header, then inflate. Enough for word/document.xml in any real
// DOCX; avoids pulling in a zip dependency for a single read.
function readZipEntry(buffer: Buffer, name: string) {
  const eocd = buffer.lastIndexOf(Buffer.from("PK\x05\x06"))
  if (eocd < 0) return undefined
  const count = buffer.readUInt16LE(eocd + 10)
  let offset = buffer.readUInt32LE(eocd + 16)
  for (let i = 0; i < count; i++) {
    if (buffer.readUInt32LE(offset) !== 0x02014b50) return undefined
    const method = buffer.readUInt16LE(offset + 10)
    const compressedSize = buffer.readUInt32LE(offset + 20)
    const nameLength = buffer.readUInt16LE(offset + 28)
    const extraLength = buffer.readUInt16LE(offset + 30)
    const commentLength = buffer.readUInt16LE(offset + 32)
    const localOffset = buffer.readUInt32LE(offset + 42)
    if (buffer.toString("utf8", offset + 46, offset + 46 + nameLength) === name) {
      const dataStart =
        localOffset + 30 + buffer.readUInt16LE(localOffset + 26) + buffer.readUInt16LE(localOffset + 28)
      const data = buffer.subarray(dataStart, dataStart + compressedSize)
      return method === 0 ? data : inflateRawSync(data)
    }
    offset += 46 + nameLength + extraLength + commentLength
  }
  return undefined
}
