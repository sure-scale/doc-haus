import { afterEach, expect, test } from "bun:test"
import { mkdir, mkdtemp, rm, writeFile } from "node:fs/promises"
import os from "node:os"
import path from "node:path"
import { dependencyKey } from "../install-dependencies"

const roots: string[] = []
const installer = path.resolve(import.meta.dir, "../install-dependencies.ts")
afterEach(async () => {
  for (const root of roots.splice(0)) await rm(root, { recursive: true, force: true })
})

async function fixture() {
  const root = await mkdtemp(path.join(os.tmpdir(), "dochaus-dependencies-"))
  roots.push(root)
  await writeFile(path.join(root, "package.json"), JSON.stringify({ name: "launcher-install-test", private: true }))
  const initial = Bun.spawn([process.execPath, "install", "--lockfile-only"], { cwd: root, stdout: "ignore", stderr: "pipe" })
  expect(await initial.exited).toBe(0)
  return root
}

async function run(root: string, force = false) {
  const child = Bun.spawn([process.execPath, installer, root], {
    cwd: root, env: { ...process.env, FORCE_INSTALL: force ? "1" : "" }, stdout: "pipe", stderr: "pipe",
  })
  const [output, error, status] = await Promise.all([new Response(child.stdout).text(), new Response(child.stderr).text(), child.exited])
  return { output, error, status }
}

test("unchanged installs are skipped, while force and a removed installed tree reinstall", async () => {
  const root = await fixture()
  const first = await run(root)
  expect(first.status).toBe(0)
  expect(first.output).toContain("installing dependencies")
  const cached = await run(root)
  expect(cached.status).toBe(0)
  expect(cached.output).toContain("dependencies ready")
  expect(cached.output).not.toContain("bun install")
  const forced = await run(root, true)
  expect(forced.status).toBe(0)
  expect(forced.output).toContain("installing dependencies")
  await rm(path.join(root, "node_modules"), { recursive: true, force: true })
  expect((await run(root)).output).toContain("installing dependencies")
})

test("dependency inputs invalidate the cache, including workspace manifests and patches", async () => {
  const root = await fixture()
  await mkdir(path.join(root, "packages", "child"), { recursive: true })
  await mkdir(path.join(root, "patches"))
  await writeFile(path.join(root, "package.json"), JSON.stringify({
    name: "launcher-install-test", workspaces: { packages: ["packages/*"] }, patchedDependencies: { "example@1": "patches/example.patch" },
  }))
  await writeFile(path.join(root, "packages", "child", "package.json"), '{"name":"child","version":"1.0.0"}')
  await writeFile(path.join(root, "patches", "example.patch"), "patch one")
  let key = dependencyKey(root)
  for (const [file, content] of [
    ["packages/child/package.json", '{"name":"child","version":"1.0.1"}'],
    ["patches/example.patch", "patch two"],
    ["bun.lock", "changed lockfile"],
    ["bunfig.toml", "[install]\nexact = true\n"],
    ["package.json", '{"name":"changed-root"}'],
  ]) {
    await writeFile(path.join(root, file!), content!)
    const next = dependencyKey(root)
    expect(next).not.toBe(key)
    key = next
  }
})

test("a failed frozen install cannot leave a successful cache record", async () => {
  const root = await fixture()
  expect((await run(root)).status).toBe(0)
  // Deliberately incompatible lock input causes Bun to reject the install.
  await writeFile(path.join(root, "bun.lock"), "invalid lockfile")
  const failed = await run(root)
  expect(failed.status).not.toBe(0)
  expect(await Bun.file(path.join(root, "node_modules", ".dochaus-install-key")).exists()).toBe(false)
})
