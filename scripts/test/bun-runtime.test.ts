import { afterEach, expect, test } from "bun:test"
import { chmod, copyFile, mkdir, mkdtemp, readdir, realpath, rm, writeFile } from "node:fs/promises"
import os from "node:os"
import path from "node:path"

const repository = path.resolve(import.meta.dir, "../..")
const roots: string[] = []
afterEach(async () => { for (const root of roots.splice(0)) await rm(root, { recursive: true, force: true }) })

// Run the real launcher and ownership helpers, stopping at the first install
// invocation. No dependency install, model download or server is started.
async function fixture() {
  const root = await mkdtemp(path.join(os.tmpdir(), "dochaus-bun-runtime-"))
  roots.push(root)
  const bin = path.join(root, "bin"), home = path.join(root, "home"), runtime = path.join(root, "data", "runtime")
  for (const directory of [bin, home, path.join(root, "scripts")]) await mkdir(directory, { recursive: true })
  for (const file of ["start.sh", "scripts/start-runtime.sh", "scripts/data-paths.sh"]) await copyFile(path.join(repository, file), path.join(root, file))
  await writeFile(path.join(bin, "lsof"), "#!/bin/sh\nexit 0\n")
  await chmod(path.join(bin, "lsof"), 0o755)
  async function executable(file: string, version: string, valid = true) {
    await mkdir(path.dirname(file), { recursive: true })
    await writeFile(file, `#!/bin/bash
if [ "$1" = "--version" ]; then
  echo '${version}'
  exit ${valid ? "0" : "7"}
fi
printf 'SELECTED:%s\\n' "$BUN_BIN"
printf 'CHILD:%s\\n' "$(bun --version)"
exit 23
`)
    await chmod(file, 0o755)
    return file
  }
  async function run(override?: string) {
    const child = Bun.spawn(["/bin/bash", path.join(root, "start.sh")], {
      cwd: root, stdin: "ignore", stdout: "pipe", stderr: "pipe",
      env: { HOME: home, PATH: `${bin}:/usr/bin:/bin`, DOCHAUS_DATA_ROOT: path.join(root, "data"), ...(override !== undefined ? { BUN_BIN: override } : {}) },
    })
    const [status, stdout, stderr] = await Promise.all([child.exited, new Response(child.stdout).text(), new Response(child.stderr).text()])
    return { status, output: stdout + stderr }
  }
  return { root, bin, home, runtime, executable, run }
}

test("PATH Bun 1.4.3 is used even when the old baseline is installed at home", async () => {
  const f = await fixture()
  const selected = await f.executable(path.join(f.bin, "bun"), "1.4.3")
  await f.executable(path.join(f.home, ".bun/bin/bun"), "1.4.2")
  const result = await f.run()
  expect(result.status).toBe(23)
  expect(result.output).toContain(`SELECTED:${selected}`)
  expect(result.output).toContain("CHILD:1.4.3")
  expect(await readdir(f.runtime)).not.toContain("engine.lock")
})

test("Bun 1.4.0 and newer numeric versions are accepted without exact-version gating", async () => {
  const f = await fixture()
  for (const version of ["1.4.0", "1.4.0-canary.1", "1.5.0", "2.0.0", "1.99.0-canary.123+build.4"]) {
    await f.executable(path.join(f.bin, "bun"), version)
    const result = await f.run()
    expect(result.status).toBe(23)
    expect(result.output).toContain(`CHILD:${version}`)
  }
})

test("Bun below 1.4.0 fails clearly before creating runtime data", async () => {
  const f = await fixture()
  for (const version of ["0.99.0", "1.3.99"]) {
    await f.executable(path.join(f.bin, "bun"), version)
    const result = await f.run()
    expect(result.status).toBe(1)
    expect(result.output).toContain(`requires Bun 1.4.0 or newer (found: ${version})`)
    expect(result.output).not.toContain("SELECTED:")
    expect(await Bun.file(path.join(f.runtime, "engine.lock/pid")).exists()).toBe(false)
  }
})

test("the official home install is a fallback when PATH has no Bun", async () => {
  const f = await fixture()
  const selected = await f.executable(path.join(f.home, ".bun/bin/bun"), "1.4.3")
  const result = await f.run()
  expect(result.status).toBe(23)
  expect(result.output).toContain(`SELECTED:${selected}`)
  expect(result.output).toContain("CHILD:1.4.3")
})

for (const kind of ["absolute", "relative", "command"] as const) {
  test(`explicit ${kind} BUN_BIN preserves the selected custom binary for child bun commands`, async () => {
    const f = await fixture()
    await f.executable(path.join(f.bin, "bun"), "1.4.3")
    const selected = await f.executable(kind === "command" ? path.join(f.bin, "custom-bun") : path.join(f.root, "custom runtime", "chosen-bun"), "2.0.0")
    const override = kind === "absolute" ? selected : kind === "relative" ? "./custom runtime/chosen-bun" : "custom-bun"
    const result = await f.run(override)
    expect(result.status).toBe(23)
    expect(result.output).toContain(`SELECTED:${kind === "relative" ? await realpath(selected) : selected}`)
    expect(result.output).toContain("CHILD:2.0.0")
    expect((await readdir(f.runtime)).some(name => name.startsWith("bun-path."))).toBe(false)
    expect(await readdir(f.runtime)).not.toContain("engine.lock")
  })
}

test("missing Bun fails clearly before creating runtime data", async () => {
  const f = await fixture()
  const result = await f.run()
  expect(result.status).toBe(1)
  expect(result.output).toContain("doc.haus could not run Bun")
  expect(result.output).toContain("set BUN_BIN to a working Bun executable")
  expect(await Bun.file(path.join(f.runtime, "engine.lock/pid")).exists()).toBe(false)
})

test("an unusable explicit override fails instead of choosing a different installed Bun", async () => {
  const f = await fixture()
  await f.executable(path.join(f.bin, "bun"), "1.4.3")
  const invalid = await f.executable(path.join(f.root, "broken-bun"), "not-a-bun-version")
  const nonzero = await f.executable(path.join(f.root, "failing-bun"), "1.4.3", false)
  for (const override of [path.join(f.root, "missing-bun"), invalid, nonzero]) {
    const result = await f.run(override)
    expect(result.status).toBe(1)
    expect(result.output).toContain("doc.haus could not run Bun")
    expect(result.output).not.toContain("SELECTED:")
  }
})
