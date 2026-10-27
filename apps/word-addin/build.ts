import { mkdir, copyFile, writeFile } from "node:fs/promises"
import { join } from "node:path"
import { renderManifest } from "./src/manifest"

const root = import.meta.dir
const output = join(root, "dist")
await mkdir(join(output, "assets"), { recursive: true })
const result = await Bun.build({
  entrypoints: [join(root, "src/pane.ts")],
  outdir: output,
  target: "browser",
  minify: true,
  naming: "pane.js",
})
if (!result.success) throw new AggregateError(result.logs, "Word add-in build failed")
await Promise.all([
  copyFile(join(root, "pane.html"), join(output, "pane.html")),
  copyFile(join(root, "pane.css"), join(output, "pane.css")),
  copyFile(join(root, "../../scripts/icons/doc-haus.png"), join(output, "assets/icon.png")),
])
if (process.env.WORD_ADDIN_ORIGIN) {
  await writeFile(join(output, "manifest.xml"), renderManifest(process.env.WORD_ADDIN_ORIGIN))
}
console.log("Built Word add-in static files.")
