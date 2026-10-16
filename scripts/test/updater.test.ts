import { afterEach, expect, test } from "bun:test"
import { mkdtemp, mkdir, rm } from "node:fs/promises"
import os from "node:os"
import path from "node:path"
import { RepoUpdater } from "../updater"
import { updateRoutes, managedUpdateOwner } from "../../services/ingest/src/updater"

const roots: string[] = []
afterEach(async () => { for (const root of roots.splice(0)) await rm(root, { recursive: true, force: true }) })

async function fixture(launcher = false) {
  const root = await mkdtemp(path.join(os.tmpdir(), "dochaus-updater-"))
  roots.push(root)
  const source = path.join(root, "source")
  const clone = path.join(root, "clone")
  const runtime = path.join(root, "runtime")
  await mkdir(source)
  const git = async (cwd: string, ...args: string[]) => {
    const child = Bun.spawn(["git", ...args], { cwd, stdout: "pipe", stderr: "pipe", env: { ...process.env, GIT_AUTHOR_NAME: "Test", GIT_AUTHOR_EMAIL: "test@example.invalid", GIT_COMMITTER_NAME: "Test", GIT_COMMITTER_EMAIL: "test@example.invalid" } })
    const output = await new Response(child.stdout).text()
    if (await child.exited !== 0) throw new Error(await new Response(child.stderr).text())
    return output.trim()
  }
  await git(source, "init", "-b", "dochaus")
  await Bun.write(path.join(source, "dochaus/opencode.json"), '{"model":"default/model"}')
  if (launcher) {
    await Bun.write(path.join(source, "apps/web/package.json"), '{"name":"web-test","private":true}')
    for (const file of ["start.sh", "scripts/start-runtime.sh", "scripts/data-paths.sh", "scripts/updater.ts", "scripts/install-dependencies.ts", "services/engine/src/data-paths.ts"])
      await Bun.write(path.join(source, file), Bun.file(path.resolve(import.meta.dir, "../..", file)))
    await Bun.write(path.join(source, ".gitignore"), ".dochaus-runtime/\nworkspace/\nnode_modules/\n")
    for (const directory of [source, path.join(source, "services/ingest")]) {
      await Bun.write(path.join(directory, "package.json"), '{"name":"updater-test","private":true}')
      const install = Bun.spawn([process.execPath, "install"], { cwd: directory, stdout: "ignore", stderr: "ignore" })
      expect(await install.exited).toBe(0)
    }
  }
  const publish = async (version: number) => {
    await Bun.write(path.join(source, "release.json"), JSON.stringify({ version, name: `Preview ${version}`, notes: "Release notes" }))
    await git(source, "add", ".")
    await git(source, "commit", "-m", `Release ${version}`)
    return git(source, "rev-parse", "HEAD")
  }
  const original = await publish(1)
  await git(root, "clone", source, clone)
  return { source, clone, runtime, publish, git, original, updater: new RepoUpdater(clone, runtime) }
}

test("polling only offers bumped releases and leaves HEAD unchanged", async () => {
  const f = await fixture()
  expect((await f.updater.check()).available).toBe(false)
  await Bun.write(path.join(f.source, "code.txt"), "ordinary development")
  await f.git(f.source, "add", ".")
  await f.git(f.source, "commit", "-m", "Development without a release")
  expect((await f.updater.check(true)).available).toBe(false)
  const target = await f.publish(2)
  await Bun.write(path.join(f.source, "unreleased.txt"), "not released yet")
  await f.git(f.source, "add", ".")
  await f.git(f.source, "commit", "-m", "Unreleased development after version 2")
  const status = await f.updater.check(true)
  expect(status.available).toBe(true)
  expect(status.target).toBe(target)
  expect(await f.git(f.clone, "rev-parse", "HEAD")).toBe(f.original)
})

test("fast-forward applies the exact offered version and installs both packages", async () => {
  const f = await fixture()
  const target = await f.publish(2)
  await f.updater.check(true)
  await f.publish(3)
  const installed: string[] = []
  await f.updater.apply(target, async directory => { installed.push(directory); return 0 })
  expect(installed).toEqual([f.clone, path.join(f.clone, "services/ingest")])
  expect(await f.git(f.clone, "rev-parse", "HEAD")).toBe(target)
  expect(new RepoUpdater(f.clone, f.runtime).current.version).toBe(2)
})

test("tracked and untracked local changes and divergent branches are preserved", async () => {
  const f = await fixture()
  const target = await f.publish(2)
  await f.updater.check()
  await f.updater.requireApplicable(target)
  await Bun.write(path.join(f.clone, "release.json"), "local edits")
  await expect(f.updater.requireApplicable(target)).rejects.toThrow("local changes")
  await f.git(f.clone, "restore", "release.json")
  await Bun.write(path.join(f.clone, "local.txt"), "keep me")
  await expect(f.updater.requireApplicable(target)).rejects.toThrow("local changes")
  await f.git(f.clone, "add", ".")
  await f.git(f.clone, "commit", "-m", "Local development")
  await expect(f.updater.requireApplicable(target)).rejects.toThrow("separate code changes")
  expect(await Bun.file(path.join(f.clone, "local.txt")).text()).toBe("keep me")
})

test("failed dependency installation restores the previous version", async () => {
  const f = await fixture()
  const target = await f.publish(2)
  await f.updater.check()
  let attempts = 0
  await expect(f.updater.apply(target, async () => ++attempts === 1 ? 1 : 0)).rejects.toThrow("previous version has been restored")
  expect(await f.git(f.clone, "rev-parse", "HEAD")).toBe(f.original)
  expect((await f.updater.check()).phase).toBe("failed")
  await f.updater.writeResult({ phase: "restarting" })
  expect((await f.updater.check()).phase).toBeUndefined()
})

test("runtime overrides and data survive updates and rollback; tracked config edits require migration", async () => {
  const f = await fixture()
  const file = path.join(f.runtime, "config/opencode/opencode.json")
  const settings = '{"model":"firm/model","small_model":"firm/fast"}'
  await Bun.write(file, settings)
  await Bun.write(path.join(f.runtime, "data/session.db"), "durable-session-fixture")
  await Bun.write(path.join(f.source, "dochaus/opencode.json"), '{}')
  const target = await f.publish(2)
  await f.updater.check()
  await f.updater.apply(target, async () => 0)
  expect(await Bun.file(file).text()).toBe(settings)
  expect(await Bun.file(path.join(f.runtime, "data/session.db")).text()).toBe("durable-session-fixture")
  const next = await f.publish(3)
  const updater = new RepoUpdater(f.clone, f.runtime)
  await updater.check()
  let attempts = 0
  await expect(updater.apply(next, async () => ++attempts === 1 ? 1 : 0)).rejects.toThrow("previous version has been restored")
  expect(await f.git(f.clone, "rev-parse", "HEAD")).toBe(target)
  expect(await Bun.file(file).text()).toBe(settings)
  expect(await Bun.file(path.join(f.runtime, "data/session.db")).text()).toBe("durable-session-fixture")
  await Bun.write(path.join(f.clone, "dochaus/opencode.json"), settings)
  await expect(updater.requireApplicable(next)).rejects.toThrow("local changes")
  expect(await Bun.file(path.join(f.clone, "dochaus/opencode.json")).text()).toBe(settings)
})

test("disconnected and unconfigured repos return a non-blocking check error", async () => {
  const f = await fixture()
  await f.git(f.clone, "config", "--unset", "branch.dochaus.remote")
  const result = await f.updater.check()
  expect(result.available).toBe(false)
  expect(result.error).toBeTruthy()
})

test("HTTP rejects foreign origins, active work, stale targets and unmanaged launchers", async () => {
  const f = await fixture()
  const target = await f.publish(2)
  let busy = false
  const app = updateRoutes(f.updater, async () => busy, async () => { throw new Error("Start Doc.Haus with start.sh") })
  const request = (body = target, origin = "http://localhost:5174", header = "1") => app.request("/updates/apply", { method: "POST", headers: { origin, "Content-Type": "application/json", "x-dochaus-update": header }, body: JSON.stringify({ target: body }) })
  expect((await app.request("/updates")).status).toBe(200)
  expect((await request(target, "https://foreign.invalid")).status).toBe(403)
  expect((await request(target, "http://localhost:5174", "")).status).toBe(403)
  busy = true
  const blocked = await request()
  expect(blocked.status).toBe(409)
  expect(await blocked.text()).toContain("Work is still in progress")
  busy = false
  expect(await (await request("stale")).text()).toContain("available update changed")
  expect(await (await request()).text()).toContain("start.sh")
  expect(await f.git(f.clone, "rev-parse", "HEAD")).toBe(f.original)
})

test("Update now signals the real supervisor, installs and restarts its child processes", async () => {
  const f = await fixture(true)
  const runtime = path.join(f.clone, ".dochaus-runtime")
  const shim = path.join(path.dirname(f.clone), "bun-shim")
  await Bun.write(shim, `#!/bin/bash
set -euo pipefail
if [ "$1" = "--version" ]; then echo 1.4.2; exit 0; fi
case "$*" in
  "run services/engine/src/server.ts")
    echo "$$" > "$DOCHAUS_RUNTIME_DIR/engine-process.lock"
    trap ':' USR2
    trap 'rm -f "$DOCHAUS_RUNTIME_DIR/engine-process.lock"; exit 0' TERM
    while true; do sleep 0.1 || true; done ;;
  "run src/server.ts"|"run dev "*)
    trap 'exit 0' TERM
    while true; do sleep 0.1; done ;;
  "run scripts/updater.ts --apply") exec "$TEST_REAL_BUN" "$@" ;;
  *) exit 0 ;;
esac
`)
  await Bun.$`chmod +x ${shim}`.quiet()
  const owner = Bun.spawn(["bash", "./start.sh"], { cwd: f.clone, env: { ...process.env, BUN_BIN: shim, TEST_REAL_BUN: process.execPath, DOCHAUS_RUNTIME_DIR: runtime, WORKSPACE_ROOT: path.join(f.clone, "workspace"), DOCHAUS_DATA_ROOT: "", SKIP_INSTALL: "1", NO_OPEN: "1", DOC_HAUS_ENGINE_PORT: "49497", DOC_HAUS_INGEST_PORT: "49501", DOC_HAUS_WEB_PORT: "49574" }, stdout: "pipe", stderr: "pipe" })
  const saved = process.env.DOCHAUS_LAUNCHER_PID
  try {
    const record = path.join(runtime, "engine-process.lock")
    await until(async () => await Bun.file(record).exists())
    const oldEngine = await Bun.file(record).text()
    expect((await Bun.file(path.join(runtime, "engine.lock/pid")).text()).trim()).toBe(String(owner.pid))
    const target = await f.publish(2)
    const updater = new RepoUpdater(f.clone, runtime)
    await updater.check()
    process.env.DOCHAUS_LAUNCHER_PID = String(owner.pid)
    expect(owner.exitCode).toBeNull()
    expect((await Bun.file(path.join(runtime, "engine.lock/pid")).text()).trim()).toBe(String(owner.pid))
    const app = updateRoutes(updater, async () => false, () => managedUpdateOwner(f.clone, runtime), () => { process.kill(owner.pid, "SIGUSR1") })
    app.post("/work", c => c.json({ ok: true }))
    const accepted = await app.request("/updates/apply", { method: "POST", headers: { origin: "http://localhost:5174", "Content-Type": "application/json", "x-dochaus-update": "1" }, body: JSON.stringify({ target }) })
    if (accepted.status !== 202) throw new Error(await accepted.text())
    expect(accepted.status).toBe(202)
    expect((await app.request("/updates/apply", { method: "POST", headers: { origin: "http://localhost:5174", "x-dochaus-update": "1" } })).status).toBe(409)
    expect((await app.request("/work", { method: "POST" })).status).toBe(503)
    await until(async () => (await Bun.file(record).text().catch(() => oldEngine)) !== oldEngine && (await Bun.file(path.join(f.clone, "release.json")).json()).version === 2)
    expect(await f.git(f.clone, "rev-parse", "HEAD")).toBe(target)
    expect(owner.exitCode).toBeNull()
    expect(new RepoUpdater(f.clone, runtime).current.version).toBe(2)
  } catch (error) {
    owner.kill("SIGTERM")
    await owner.exited
    throw new Error(`${error}\n${await new Response(owner.stdout).text()}\n${await new Response(owner.stderr).text()}`)
  } finally {
    if (saved === undefined) delete process.env.DOCHAUS_LAUNCHER_PID
    else process.env.DOCHAUS_LAUNCHER_PID = saved
    owner.kill("SIGTERM")
    await owner.exited
  }
}, 15_000)

async function until(ready: () => Promise<boolean>) {
  const deadline = Date.now() + 10_000
  while (!await ready()) {
    if (Date.now() > deadline) throw new Error("Update supervisor did not become ready")
    await Bun.sleep(25)
  }
}
