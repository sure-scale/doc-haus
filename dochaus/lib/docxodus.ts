import { fileURLToPath } from "node:url"
import path from "node:path"

// Shared Docxodus engine bootstrap for the doc.haus tools (word-integration,
// tracked-changes). Docxodus is a .NET 10 OOXML engine compiled to WASM; it loads
// headless under Bun via initialize(wasmBasePath).
//
// Docxodus ships its WASM runtime alongside its entrypoint (dist/index.js →
// dist/wasm). The package's exports map hides package.json, and Bun can't resolve
// bare specifiers from nested tool dirs via createRequire, so resolve the main
// entry with import.meta.resolve and walk to its sibling wasm dir.
const wasmBase = path.join(path.dirname(fileURLToPath(import.meta.resolve("docxodus"))), "wasm")

let engine: Promise<typeof import("docxodus")> | undefined

// The .NET 10 runtime in Docxodus 13 prints "MONO_WASM: Jiterpreter table N is not
// yet initialized" through console.log (stdout) while it starts, which corrupts
// JSON read from a child's stdout. Docxodus passes no runtime options through, so
// drop exactly those lines during initialize. The runtime binds its print sink to
// the console.log in effect at startup, so later notices of that kind stay dropped;
// all other output passes, including the runtime's "is out of space" warning.
// console.log is restored unless something replaced it meanwhile (ingest loads
// both loaders, so a concurrent start can leave the other's filter installed).
async function withoutJiterpreterNotices(start: () => Promise<void>) {
  const log = console.log
  const quiet = (...args: unknown[]) => {
    if (typeof args[0] === "string" && /^MONO_WASM: Jiterpreter table \d+ is not yet initialized/.test(args[0])) return
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
