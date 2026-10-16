import { afterEach, expect, test } from "bun:test"
import { mkdir, mkdtemp, readFile, rm, writeFile } from "node:fs/promises"
import os from "node:os"
import path from "node:path"

const helper = path.resolve(import.meta.dir, "../start-runtime.sh")
const fixtures: { root: string; processes: ReturnType<typeof Bun.spawn>[] }[] = []
afterEach(async () => {
  for (const fixture of fixtures.splice(0)) {
    for (const process of fixture.processes) {
      if (process.exitCode === null) process.kill("SIGTERM")
      await process.exited
    }
    await rm(fixture.root, { recursive: true, force: true })
  }
})

async function fixture() {
  const root = await mkdtemp(path.join(os.tmpdir(), "dochaus-start-runtime-"))
  const runtime = path.join(root, "runtime")
  const lock = path.join(runtime, "engine.lock")
  const processes: ReturnType<typeof Bun.spawn>[] = []
  fixtures.push({ root, processes })
  await mkdir(runtime)
  const harness = path.join(root, "claim.sh")
  await writeFile(harness, `#!/bin/bash
set -euo pipefail
cd "$1"
. "$2"
status=0
claim_engine_runtime "$1/runtime" || status=$?
if [ "$status" = 0 ]; then
  echo CLAIMED
  rm -f "$1/runtime/engine.lock/pid"
  rmdir "$1/runtime/engine.lock"
fi
exit "$status"
`)
  const launch = async () => {
    const script = path.join(root, "start.sh")
    await writeFile(script, `#!/bin/bash
set -euo pipefail
mkdir "$1/engine.lock"
echo "$$" > "$1/engine.lock/pid"
trap 'rm -f "$1/engine.lock/pid"; rmdir "$1/engine.lock"; exit 0' TERM
while true; do sleep 0.1; done
`)
    const process = Bun.spawn(["bash", script, runtime], { cwd: root, stdout: "ignore", stderr: "inherit" })
    processes.push(process)
    await until(async () => (await readFile(path.join(lock, "pid"), "utf8").catch(() => "")).trim() === String(process.pid))
    return process
  }
  const run = async (answer?: string) => {
    const process = answer === undefined
      ? Bun.spawn(["bash", harness, root, helper], { cwd: root, stdin: "ignore", stdout: "pipe", stderr: "pipe" })
      : Bun.spawn(["python3", "-c", terminal, harness, root, helper, answer], { cwd: root, stdout: "pipe", stderr: "pipe" })
    processes.push(process)
    const [stdout, stderr, status] = await Promise.all([new Response(process.stdout).text(), new Response(process.stderr).text(), process.exited])
    return { output: stdout + stderr, status }
  }
  return { root, runtime, lock, processes, launch, run }
}

test("an interactive decline leaves the existing launcher running", async () => {
  const f = await fixture()
  const owner = await f.launch()
  const result = await f.run("n")
  expect(result.status).toBe(2)
  expect(result.output).toContain("already running")
  expect(result.output).toContain("Stop that stack and restart it? [y/N]")
  expect(owner.exitCode).toBeNull()
  expect((await readFile(path.join(f.lock, "pid"), "utf8")).trim()).toBe(String(owner.pid))
})

test("an interactive yes shuts down only the launcher and reclaims its runtime", async () => {
  const f = await fixture()
  const owner = await f.launch()
  const unrelated = Bun.spawn(["sleep", "30"], { cwd: f.root })
  f.processes.push(unrelated)
  const result = await f.run("y")
  expect(result.status).toBe(0)
  expect(result.output).toContain("CLAIMED")
  expect(await owner.exited).toBe(0)
  expect(unrelated.exitCode).toBeNull()
})

test("noninteractive calls leave a running stack alone", async () => {
  const f = await fixture()
  const owner = await f.launch()
  const result = await f.run()
  expect(result.status).toBe(2)
  expect(result.output).toContain("Run ./start.sh in a terminal")
  expect(owner.exitCode).toBeNull()
})

test("stale locks are reclaimed but a surviving engine still blocks a second stack", async () => {
  const f = await fixture()
  const exited = Bun.spawn(["true"])
  await exited.exited
  await mkdir(f.lock)
  await writeFile(path.join(f.lock, "pid"), String(exited.pid))
  const engine = Bun.spawn(["sleep", "30"], { cwd: f.root })
  f.processes.push(engine)
  await writeFile(path.join(f.runtime, "engine-process.lock"), String(engine.pid))
  const blocked = await f.run()
  expect(blocked.status).toBe(1)
  expect(blocked.output).toContain("the engine is still running")
  expect(engine.exitCode).toBeNull()
  engine.kill("SIGTERM")
  await engine.exited
  const reclaimed = await f.run()
  expect(reclaimed.status).toBe(0)
  expect(reclaimed.output).toContain("clearing a stale launcher lock")
  expect(reclaimed.output).toContain("clearing a stale engine lock")
  expect(await Bun.file(path.join(f.runtime, "engine-process.lock")).exists()).toBe(false)
})

test("an abandoned engine record is removed even without a launcher lock", async () => {
  const f = await fixture()
  const exited = Bun.spawn(["true"])
  await exited.exited
  const record = path.join(f.runtime, "engine-process.lock")
  await writeFile(record, String(exited.pid))
  const result = await f.run()
  expect(result.status).toBe(0)
  expect(result.output).toContain("clearing a stale engine lock")
  // Prove the engine's actual exclusive-create operation can now succeed.
  const descriptor = await import("node:fs/promises").then(fs => fs.open(record, "wx", 0o600))
  await descriptor.close()
})

test("unverifiable engine ownership is preserved", async () => {
  const f = await fixture()
  const record = path.join(f.runtime, "engine-process.lock")
  await writeFile(record, "invalid")
  const result = await f.run()
  expect(result.status).toBe(1)
  expect(result.output).toContain("ownership could not be verified")
  expect(await readFile(record, "utf8")).toBe("invalid")
})

test("a reused PID belonging to an unrelated process is never stopped", async () => {
  const f = await fixture()
  const owner = Bun.spawn(["sleep", "30"], { cwd: f.root })
  f.processes.push(owner)
  await mkdir(f.lock)
  await writeFile(path.join(f.lock, "pid"), String(owner.pid))
  const result = await f.run("y")
  expect(result.status).toBe(1)
  expect(result.output).toContain("could not be identified")
  expect(owner.exitCode).toBeNull()
})

test("a surviving engine with no launcher lock still blocks a replacement", async () => {
  const f = await fixture()
  const engine = Bun.spawn(["sleep", "30"], { cwd: f.root })
  f.processes.push(engine)
  await writeFile(path.join(f.runtime, "engine-process.lock"), String(engine.pid))
  const result = await f.run()
  expect(result.status).toBe(1)
  expect(result.output).toContain("the engine is still running")
  expect(engine.exitCode).toBeNull()
  expect(await readFile(path.join(f.lock, "pid"), "utf8").catch(() => undefined)).toBeUndefined()
})

test("restart waits for the old engine's graceful shutdown after its launcher exits", async () => {
  const f = await fixture()
  const script = path.join(f.root, "start.sh")
  const engineScript = path.join(f.root, "engine.sh")
  await writeFile(engineScript, `#!/bin/bash
echo "$$" > "$1/engine-process.lock"
trap 'sleep 0.5; exit 0' TERM
while true; do sleep 0.1; done
`)
  const engine = Bun.spawn(["bash", engineScript, f.runtime], { cwd: f.root })
  f.processes.push(engine)
  await until(async () => (await readFile(path.join(f.runtime, "engine-process.lock"), "utf8").catch(() => "")).trim() === String(engine.pid))
  await writeFile(script, `#!/bin/bash
mkdir "$1/engine.lock"
echo "$$" > "$1/engine.lock/pid"
trap 'kill -TERM "$2"; rm -f "$1/engine.lock/pid"; rmdir "$1/engine.lock"; exit 0' TERM
while true; do sleep 0.1; done
`)
  const owner = Bun.spawn(["bash", script, f.runtime, String(engine.pid)], { cwd: f.root })
  f.processes.push(owner)
  await until(async () => (await readFile(path.join(f.lock, "pid"), "utf8").catch(() => "")).trim() === String(owner.pid))
  const result = await f.run("y")
  expect(result.status).toBe(0)
  expect(result.output).toContain("CLAIMED")
  expect(owner.exitCode).toBe(0)
  expect(engine.exitCode).toBe(0)
})

async function until(ready: () => Promise<boolean>) {
  const deadline = Date.now() + 3000
  while (!await ready()) {
    if (Date.now() > deadline) throw new Error("Fixture launcher did not start")
    await Bun.sleep(20)
  }
}

// A real terminal makes bash's interactive check and read behave like ./start.sh.
const terminal = `
import errno, os, pty, select, subprocess, sys, time
master, slave = pty.openpty()
proc = subprocess.Popen(["bash", *sys.argv[1:4]], stdin=slave, stdout=slave, stderr=slave)
os.close(slave)
output = bytearray()
answered = False
deadline = time.monotonic() + 20
try:
    while time.monotonic() < deadline:
        if select.select([master], [], [], 0.1)[0]:
            try:
                chunk = os.read(master, 65536)
            except OSError as error:
                if error.errno == errno.EIO: break
                raise
            if not chunk: break
            output.extend(chunk)
            if not answered and b"[y/N]" in output:
                os.write(master, (sys.argv[4] + "\\n").encode())
                answered = True
        elif proc.poll() is not None:
            break
    else:
        proc.terminate()
        raise RuntimeError("Launcher terminal did not finish")
    sys.stdout.buffer.write(output)
    sys.exit(proc.wait(timeout=3))
finally:
    os.close(master)
    if proc.poll() is None: proc.terminate()
`
