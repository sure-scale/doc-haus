import { defineConfig, type Plugin } from "vite"
import react from "@vitejs/plugin-react"
import path from "node:path"
import fs from "node:fs"
import { fileURLToPath } from "node:url"

// Docxodus ships its ~16MB WASM runtime under the package's dist/wasm. The React
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

// The OpenCode SDK client is pure, self-contained TypeScript (no runtime npm
// deps on the client path), so we alias straight to its source instead of
// installing it. This keeps apps/web a standalone package — no root workspace
// edit, no `file:` dep that would drag in the SDK's server-side deps — which
// keeps upstream merges clean.
export default defineConfig({
  plugins: [react(), docxodusWasm()],
  resolve: {
    alias: {
      "@opencode-ai/sdk": path.resolve(__dirname, "../../packages/sdk/js/src/client.ts"),
    },
  },
  server: {
    port: 5173,
  },
})
