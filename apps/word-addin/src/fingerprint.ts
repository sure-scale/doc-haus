import { unzipSync } from "fflate"

/** Compare package contents, not ZIP timestamps, compression, or entry order. */
export async function documentFingerprint(bytes: Uint8Array): Promise<string> {
  let count = 0
  let unpacked = 0
  const entries = unzipSync(bytes, {
    filter(file) {
      const path = file.name.split("/")
      if (file.name.endsWith("/") || path.includes("__MACOSX") || path.at(-1) === ".DS_Store") return false
      count++
      unpacked += file.originalSize
      // Match the receiver's package budget before fflate allocates each part.
      if (count > 1000 || unpacked > 256 * 1024 * 1024) {
        throw new Error("The Word document exceeds the package size limit. Keep the file and return to doc.haus.")
      }
      return true
    },
  })
  const encoder = new TextEncoder()
  const parts: Uint8Array[] = []
  for (const name of Object.keys(entries).sort()) {
    const content = entries[name]!
    const label = encoder.encode(name)
    const lengths = new Uint8Array(8)
    const view = new DataView(lengths.buffer)
    view.setUint32(0, label.length)
    view.setUint32(4, content.length)
    parts.push(lengths, label, content)
  }
  const canonical = new Uint8Array(parts.reduce((sum, part) => sum + part.length, 0))
  let offset = 0
  for (const part of parts) {
    canonical.set(part, offset)
    offset += part.length
  }
  const hash = await crypto.subtle.digest("SHA-256", canonical)
  return Array.from(new Uint8Array(hash), (byte) => byte.toString(16).padStart(2, "0")).join("")
}
