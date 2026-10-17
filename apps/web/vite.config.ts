import { defineConfig, type Plugin } from "vite"
import react from "@vitejs/plugin-react"
import path from "node:path"
import fs from "node:fs"
import { fileURLToPath } from "node:url"
import { backendProxy } from "./backend-proxy"

// Docxodus ships its ~30MB WASM runtime (13.x; ~25MB of .wasm plus optional .br copies) under the package's dist/wasm. The React
// hooks fetch it from a base path, so we host that directory at /wasm/ — streamed
// from node_modules in dev, copied into the build output for production. This keeps
// the runtime out of the repo (installed with the package) while the viewer still
// converts and renders DOCX entirely client-side.
function docxodusWasm(): Plugin {
  const wasmDir = path.join(path.dirname(fileURLToPath(import.meta.resolve("docxodus"))), "wasm")
  const mime: Record<string, string> = { ".wasm": "application/wasm", ".js": "text/javascript", ".json": "application/json", ".html": "text/html" }
  let outDir = "dist"
  return {
    name: "docxodus-wasm",
    configResolved(config) {
      outDir = config.build.outDir
    },
    configureServer(server) {
      server.middlewares.use("/wasm", (req, res, next) => {
        const file = path.join(wasmDir, decodeURIComponent((req.url ?? "").split("?")[0]))
        if (!file.startsWith(wasmDir) || !fs.existsSync(file) || !fs.statSync(file).isFile()) return next()
        res.setHeader("Content-Type", mime[path.extname(file)] ?? "application/octet-stream")
        fs.createReadStream(file).pipe(res)
      })
    },
    writeBundle() {
      fs.cpSync(wasmDir, path.join(outDir, "wasm"), { recursive: true })
    },
  }
}

export default defineConfig({
  plugins: [react(), docxodusWasm()],
  server: {
    // Bind localhost by default; set WEB_HOST=0.0.0.0 to expose the dev server
    // (Docker, reverse proxy), matching the engine and ingest loopback default.
    host: process.env.WEB_HOST || undefined,
    port: 5173,
    proxy: backendProxy(),
  },
})
