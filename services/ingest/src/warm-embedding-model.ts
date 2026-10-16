import { existsSync, readdirSync, rmSync } from "node:fs"
import path from "node:path"
import { DIM, MODEL, embed } from "./embed"
import { embeddingModelCacheDir } from "../../../dochaus/lib/search/model-cache"

// start.sh runs this before it reports the stack ready, so a missing or broken
// model download stops the launch with a clear message instead of failing the
// first upload. Each run is a fresh process; start.sh owns the bounded retry.
// After one success the model is in the shared cache and later loads are local.
const cacheDir = embeddingModelCacheDir()
removeAbandonedDownloads(cacheDir)
const vector = await embed("doc.haus embedding model check").catch((error: unknown) => {
  console.error(`doc.haus: ${error instanceof Error ? error.message : String(error)}`)
  process.exit(1)
})
if (vector.length !== DIM) {
  console.error(`doc.haus: the embedding model ${MODEL} returned ${vector.length} dimensions, expected ${DIM}`)
  process.exit(1)
}
console.log(`doc.haus: embedding model ready (${cacheDir})`)
process.exit(0)

// transformers.js streams each download to "<file>.tmp.<pid>.<random>" and
// renames it when complete. A process killed mid-download leaves that partial
// file behind; remove it once its writer is gone so partials never accumulate.
function removeAbandonedDownloads(directory: string) {
  if (!existsSync(directory)) return
  readdirSync(directory, { recursive: true, withFileTypes: true })
    .filter((entry) => entry.isFile())
    .flatMap((entry) => {
      const pid = /\.tmp\.(\d+)\.[a-z0-9]+$/.exec(entry.name)?.[1]
      return pid && !processRunning(Number(pid)) ? [path.join(entry.parentPath, entry.name)] : []
    })
    .forEach((file) => rmSync(file, { force: true }))
}

function processRunning(pid: number) {
  try {
    process.kill(pid, 0)
    return true
  } catch (error) {
    return (error as NodeJS.ErrnoException).code === "EPERM"
  }
}
