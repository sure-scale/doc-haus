import { fileURLToPath } from "node:url"
import path from "node:path"

// Shared Docxodus engine bootstrap for the doc.haus tools (word-integration,
// tracked-changes). Docxodus is a .NET 8 OOXML engine compiled to WASM; it loads
// headless under Bun via initialize(wasmBasePath).
//
// Docxodus ships its WASM runtime alongside its entrypoint (dist/index.js →
// dist/wasm). The package's exports map hides package.json, and Bun can't resolve
// bare specifiers from nested tool dirs via createRequire, so resolve the main
// entry with import.meta.resolve and walk to its sibling wasm dir.
const wasmBase = path.join(path.dirname(fileURLToPath(import.meta.resolve("docxodus"))), "wasm")

let engine: Promise<typeof import("docxodus")> | undefined

export function docxodus() {
  if (!engine)
    engine = import("docxodus").then(async (dx) => {
      if (!dx.isInitialized()) await dx.initialize(wasmBase)
      return dx
    })
  return engine
}
