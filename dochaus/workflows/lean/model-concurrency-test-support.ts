import { mkdtempSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"

/** Give a workflow test an isolated installation limit and restore the environment. */
export async function withWorkflowLimit<T>(limit: number, run: () => Promise<T>): Promise<T> {
  const root = mkdtempSync(path.join(tmpdir(), "dochaus-limit-test-"))
  const previousLegal = process.env.DOCHAUS_DIR
  const previousRuntime = process.env.DOCHAUS_RUNTIME_DIR
  try {
    writeFileSync(path.join(root, "opencode.json"), JSON.stringify({ model_concurrency: limit }))
    process.env.DOCHAUS_DIR = root
    process.env.DOCHAUS_RUNTIME_DIR = path.join(root, "runtime")
    return await run()
  } finally {
    if (previousLegal === undefined) delete process.env.DOCHAUS_DIR
    else process.env.DOCHAUS_DIR = previousLegal
    if (previousRuntime === undefined) delete process.env.DOCHAUS_RUNTIME_DIR
    else process.env.DOCHAUS_RUNTIME_DIR = previousRuntime
    rmSync(root, { recursive: true, force: true })
  }
}
