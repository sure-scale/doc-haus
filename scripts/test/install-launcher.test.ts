import { afterEach, expect, test } from "bun:test"
import { chmod, copyFile, mkdir, mkdtemp, readFile, rm, writeFile } from "node:fs/promises"
import os from "node:os"
import path from "node:path"

const repository = path.resolve(import.meta.dir, "../..")
const roots: string[] = []
afterEach(async () => { for (const root of roots.splice(0)) await rm(root, { recursive: true, force: true }) })

async function fixture() {
  const root = await mkdtemp(path.join(os.tmpdir(), "dochaus-install-launcher-"))
  roots.push(root)
  const home = path.join(root, "home"), bin = path.join(root, "bin")
  // Spaces and shell syntax in a selected installation must remain literal.
  const checkout = path.join(root, "doc haus $(touch unexpected-checkout)")
  for (const directory of [home, bin, path.join(checkout, "scripts")]) await mkdir(directory, { recursive: true })
  await copyFile(path.join(repository, "scripts/data-paths.sh"), path.join(checkout, "scripts/data-paths.sh"))
  for (const [file, source] of [
    ["git", "#!/bin/sh\ncase \"$1\" in --version) echo 'git version 2.50.0';; *) echo 'https://github.com/sure-scale/doc-haus.git';; esac\n"],
    ["uname", "#!/bin/sh\necho Darwin\n"],
  ] as const) {
    await writeFile(path.join(bin, file), source)
    await chmod(path.join(bin, file), 0o755)
  }
  await writeFile(path.join(checkout, "start.sh"), `#!/bin/bash
set -eu
cd "$(dirname "$0")"
. scripts/data-paths.sh
configure_data_paths "$PWD"
printf '%s\\0' "$PWD" "\${DOCHAUS_DATA_ROOT:-}" "$WORKSPACE_ROOT" "$DOCHAUS_RUNTIME_DIR" "\${DOC_HAUS_ENGINE_PORT:-4097}" "\${DOC_HAUS_INGEST_PORT:-4501}" "\${DOC_HAUS_WEB_PORT:-5174}" "\${OPENCODE_SERVER_PASSWORD:-}" > profile
`)
  await chmod(path.join(checkout, "start.sh"), 0o755)
  const environment = { HOME: home, PATH: `${bin}:/usr/bin:/bin` }
  const command = path.join(home, "Applications/doc.haus.app/Contents/Resources/doc.haus.command")
  async function run(script: string, env: Record<string, string>) {
    const child = Bun.spawn(["/bin/bash", script], { cwd: root, env, stdin: "ignore", stdout: "pipe", stderr: "pipe" })
    const [status, stderr] = await Promise.all([child.exited, new Response(child.stderr).text()])
    if (status !== 0) throw new Error(`Launcher exited ${status}: ${stderr}`)
    return (await readFile(path.join(checkout, "profile"), "utf8")).split("\0").slice(0, -1)
  }
  const install = (overrides: Record<string, string> = {}) => run(path.join(repository, "scripts/install.sh"), {
    ...environment, DOCHAUS_DIR: checkout, ...overrides,
  })
  const reopen = () => run(command, environment)
  return { root, checkout, command, install, reopen }
}

test("the Mac app reopens the selected data paths and ports without retaining credentials", async () => {
  const f = await fixture()
  const selected = {
    DOCHAUS_DATA_ROOT: path.join(f.root, "data $(touch unexpected-data) 'quoted'"),
    WORKSPACE_ROOT: "selected matters",
    DOCHAUS_RUNTIME_DIR: path.join(f.root, "runtime\nwith newline"),
    DOC_HAUS_ENGINE_PORT: "49197", DOC_HAUS_INGEST_PORT: "49201", DOC_HAUS_WEB_PORT: "49274",
  }
  const initial = await f.install({ ...selected, OPENCODE_SERVER_PASSWORD: "synthetic-install-only" })
  const reopened = await f.reopen()
  expect(reopened.slice(0, -1)).toEqual(initial.slice(0, -1))
  expect(reopened).toEqual([f.checkout, selected.DOCHAUS_DATA_ROOT, path.join(f.checkout, selected.WORKSPACE_ROOT),
    selected.DOCHAUS_RUNTIME_DIR, "49197", "49201", "49274", ""])
  const source = await readFile(f.command, "utf8")
  expect(source).not.toContain("OPENCODE_SERVER_PASSWORD")
  expect(source).not.toContain("synthetic-install-only")
  expect(await Bun.file(path.join(f.root, "unexpected-data")).exists()).toBe(false)
  expect(await Bun.file(path.join(f.root, "unexpected-checkout")).exists()).toBe(false)
})

test("the Mac app keeps repository defaults when no installation overrides are selected", async () => {
  const f = await fixture()
  const initial = await f.install()
  expect(await f.reopen()).toEqual(initial)
  expect(initial).toEqual([f.checkout, "", path.join(f.checkout, "workspace"), path.join(f.checkout, ".dochaus-runtime"), "4097", "4501", "5174", ""])
})
