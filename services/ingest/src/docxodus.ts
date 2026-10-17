import { fileURLToPath } from "node:url"
import path from "node:path"

// Docxodus engine bootstrap for the ingest service — the same .NET 8 OOXML engine
// (compiled to WASM) the dochaus tools use, loaded headless under Bun. Ingest owns
// all document I/O, so it owns baking accepted redlines into the canonical .docx and
// building the redlined comparison the viewer renders.
//
// Docxodus ships its WASM runtime alongside its entrypoint (dist/index.js →
// dist/wasm). The package's exports map hides package.json, so resolve the main
// entry with import.meta.resolve and walk to its sibling wasm dir.
const wasmBase = path.join(path.dirname(fileURLToPath(import.meta.resolve("docxodus"))), "wasm")

let engine: Promise<typeof import("docxodus")> | undefined

// The .NET 10 runtime in Docxodus 13 prints "MONO_WASM: Jiterpreter table N is not
// yet initialized" through console.log (stdout) while it starts, which corrupts
// JSON read from a child's stdout. Docxodus passes no runtime options through, so
// drop exactly those lines during initialize. The runtime binds its print sink to
// the console.log in effect at startup, so later notices of that kind stay dropped;
// console.log itself is restored and all other output passes.
async function withoutJiterpreterNotices(start: () => Promise<void>) {
  const log = console.log
  const quiet = (...args: unknown[]) => {
    if (typeof args[0] === "string" && args[0].startsWith("MONO_WASM: Jiterpreter")) return
    log.apply(console, args)
  }
  console.log = quiet
  try {
    await start()
  } finally {
    if (console.log === quiet) console.log = log
  }
}

export function docxodus() {
  if (!engine)
    engine = import("docxodus").then(async (dx) => {
      if (!dx.isInitialized()) await withoutJiterpreterNotices(() => dx.initialize(wasmBase))
      return dx
    })
  return engine
}
