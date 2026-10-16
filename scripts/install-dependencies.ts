import { createHash } from "node:crypto"
import { existsSync, readFileSync, rmSync } from "node:fs"
import path from "node:path"

// Cache successful installs beside the installed tree. Changes to dependency
// inputs or Bun invalidate it; deleting node_modules removes the cache too.
export function dependencyKey(directory: string): string {
  const files = new Set(["package.json", "bun.lock", "bun.lockb", "bunfig.toml"])
  const manifest = JSON.parse(readFileSync(path.join(directory, "package.json"), "utf8"))
  const workspaces: string[] = Array.isArray(manifest.workspaces)
    ? manifest.workspaces
    : manifest.workspaces?.packages ?? []
  for (const workspace of workspaces) {
    for (const file of new Bun.Glob(`${workspace}/package.json`).scanSync({ cwd: directory })) files.add(file)
  }
  for (const file of [...files].filter(file => file.endsWith("package.json"))) {
    const pkg = JSON.parse(readFileSync(path.join(directory, file), "utf8"))
    for (const patch of Object.values(pkg.patchedDependencies ?? {}) as string[]) {
      files.add(path.join(path.dirname(file), patch))
    }
  }
  const hash = createHash("sha256").update(`${Bun.version}:${process.platform}:${process.arch}\0`)
  for (const file of [...files].sort()) {
    const absolute = path.join(directory, file)
    hash.update(`${file}\0`)
    hash.update(existsSync(absolute) ? readFileSync(absolute) : "missing")
    hash.update("\0")
  }
  return hash.digest("hex")
}

export async function installDependencies(directory: string, force = false): Promise<number> {
  const key = dependencyKey(directory)
  const stamp = Bun.file(path.join(directory, "node_modules", ".dochaus-install-key"))
  if (!force && await stamp.exists() && (await stamp.text()).trim() === key) {
    console.log(`doc.haus: dependencies ready (${directory})`)
    return 0
  }
  console.log(`doc.haus: installing dependencies (${directory})`)
  rmSync(stamp.name!, { force: true })
  const child = Bun.spawn([process.execPath, "install", "--frozen-lockfile"], {
    cwd: directory, stdin: "inherit", stdout: "inherit", stderr: "inherit",
  })
  const status = await child.exited
  if (status === 0) await Bun.write(stamp, key)
  return status
}

if (import.meta.main) {
  process.exit(await installDependencies(process.argv[2] ?? ".", process.env.FORCE_INSTALL === "1"))
}
