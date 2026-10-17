import { afterEach, expect, test } from "bun:test"
import { chmod, mkdir, mkdtemp, readFile, rm, writeFile } from "node:fs/promises"
import os from "node:os"
import path from "node:path"

const requirements = path.resolve(import.meta.dir, "../requirements.sh")
const python = Bun.which("python3")!
const roots: string[] = []
afterEach(async () => { for (const root of roots.splice(0)) await rm(root, { recursive: true, force: true }) })

// Source the launcher helper with a PATH of stubs. Every installer and package
// manager is a stub that records its arguments, so nothing touches the host.
async function fixture() {
  const root = await mkdtemp(path.join(os.tmpdir(), "dochaus-requirements-"))
  roots.push(root)
  const bin = path.join(root, "bin"), home = path.join(root, "home"), runtime = path.join(root, "runtime"), calls = path.join(root, "calls")
  for (const directory of [bin, home, runtime]) await mkdir(directory, { recursive: true })
  await writeFile(calls, "")
  const harness = path.join(root, "harness.sh")
  await writeFile(harness, `#!/bin/bash
set -euo pipefail
. "${requirements}"
if [ "$1" = bun ]; then
  ensure_bun
  echo "SELECTED:$BUN_BIN ($BUN_FOUND)"
else
  ensure_host_tools "${runtime}"
  echo TOOLS-DONE
fi
`)
  const release = path.join(root, "bun-release")
  await writeFile(release, "#!/bin/sh\n[ \"$1\" = --version ] && echo 1.4.3\n")
  await chmod(release, 0o755)
  async function stub(name: string, body = "exit 0") {
    await writeFile(path.join(bin, name), `#!/bin/sh\necho "${name} $*" >> "${calls}"\n${body}\n`)
    await chmod(path.join(bin, name), 0o755)
  }
  // The official installer stub installs the release Bun where the real one would.
  await stub("curl", `case "$*" in *bun.sh/install*) echo 'mkdir -p "\${BUN_INSTALL:-$HOME/.bun}/bin" && cp "${release}" "\${BUN_INSTALL:-$HOME/.bun}/bin/bun"' ;; esac`)
  for (const name of ["unzip", "brew", "apt-get", "pipx"]) await stub(name, name === "brew" ? '[ "$1" = list ] && exit 1; exit 0' : "exit 0")
  await stub("sudo", 'exec "$@"')
  const env = (extra: Record<string, string>) => ({ HOME: home, PATH: `${bin}:/usr/bin:/bin`, ...extra })
  async function run(mode: "bun" | "tools", extra: Record<string, string> = {}) {
    const child = Bun.spawn(["/bin/bash", harness, mode], { cwd: root, stdin: "ignore", stdout: "pipe", stderr: "pipe", env: env(extra) })
    const [status, stdout, stderr] = await Promise.all([child.exited, new Response(child.stdout).text(), new Response(child.stderr).text()])
    return { status, output: stdout + stderr }
  }
  async function answer(mode: "bun" | "tools", reply: string, extra: Record<string, string> = {}) {
    const child = Bun.spawn([python, "-c", terminal, reply, "/bin/bash", harness, mode], { cwd: root, stdout: "pipe", stderr: "pipe", env: env(extra) })
    const [status, stdout, stderr] = await Promise.all([child.exited, new Response(child.stdout).text(), new Response(child.stderr).text()])
    return { status, output: stdout + stderr }
  }
  return { root, bin, home, runtime, stub, run, answer, calls: () => readFile(calls, "utf8") }
}

// A real terminal makes the helper's interactive check and read behave like ./start.sh.
const terminal = `
import errno, os, pty, select, subprocess, sys, time
master, slave = pty.openpty()
proc = subprocess.Popen(sys.argv[2:], stdin=slave, stdout=slave, stderr=slave)
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
            if not answered and b"[Y/n]" in output:
                os.write(master, (sys.argv[1] + "\\n").encode())
                answered = True
        elif proc.poll() is not None:
            break
    else:
        proc.terminate()
        raise RuntimeError("Helper terminal did not finish")
    sys.stdout.buffer.write(output)
    sys.exit(proc.wait(timeout=3))
finally:
    os.close(master)
    if proc.poll() is None: proc.terminate()
`

test("Enter at the Bun prompt installs Bun with the official installer and selects it", async () => {
  const f = await fixture()
  const result = await f.answer("bun", "")
  expect(result.status).toBe(0)
  expect(result.output).toContain("doc.haus could not run Bun")
  expect(result.output).toContain("[Y/n]")
  expect(await f.calls()).toContain("curl -fsSL https://bun.sh/install")
  expect(result.output).toContain(`SELECTED:${path.join(f.home, ".bun/bin/bun")} (1.4.3)`)
})

test("declining the Bun prompt stops with the install hint and installs nothing", async () => {
  const f = await fixture()
  const result = await f.answer("bun", "n")
  expect(result.status).toBe(1)
  expect(result.output).toContain("Install Bun or set BUN_BIN to a working Bun executable")
  expect(await f.calls()).toBe("")
})

test("without a terminal a missing Bun is reported, not installed", async () => {
  const f = await fixture()
  const result = await f.run("bun")
  expect(result.status).toBe(1)
  expect(result.output).not.toContain("[Y/n]")
  expect(result.output).toContain("Install Bun or set BUN_BIN")
  expect(await f.calls()).toBe("")
})

test("DOCHAUS_INSTALL_TOOLS=yes replaces a too-old Bun without asking", async () => {
  const f = await fixture()
  await f.stub("bun", '[ "$1" = --version ] && echo 1.3.0')
  const result = await f.run("bun", { DOCHAUS_INSTALL_TOOLS: "yes" })
  expect(result.status).toBe(0)
  expect(result.output).toContain("requires Bun 1.4.0 or newer (found: 1.3.0)")
  expect(result.output).toContain(`SELECTED:${path.join(f.home, ".bun/bin/bun")} (1.4.3)`)
})

test("an explicit BUN_BIN is never replaced by an install", async () => {
  const f = await fixture()
  const result = await f.run("bun", { DOCHAUS_INSTALL_TOOLS: "yes", BUN_BIN: path.join(f.root, "missing-bun") })
  expect(result.status).toBe(1)
  expect(result.output).toContain("doc.haus could not run Bun")
  expect(await f.calls()).toBe("")
})

// Usable stand-ins for every optional tool.
async function installAll(f: Awaited<ReturnType<typeof fixture>>) {
  await f.stub("deno", '[ "$1" = --version ] && echo "deno 2.8.2 (stable, release)"')
  await f.stub("python3", "echo 3.12")
  await f.stub("fc-list", "echo Carlito")
  for (const name of ["markitdown", "pdftoppm", "tesseract", "soffice", "lsof", "pgrep"]) await f.stub(name)
}

test("usable tools pass without a prompt and clear an earlier decline", async () => {
  const f = await fixture()
  await installAll(f)
  await writeFile(path.join(f.runtime, "declined-tools"), "tesseract\n")
  const result = await f.answer("tools", "n")
  expect(result.status).toBe(0)
  expect(result.output).not.toContain("[Y/n]")
  expect(result.output).not.toContain("missing")
  expect(await Bun.file(path.join(f.runtime, "declined-tools")).exists()).toBe(false)
})

test("missing and outdated tools are listed, and nothing is installed without a terminal", async () => {
  const f = await fixture()
  await f.stub("deno", '[ "$1" = --version ] && echo "deno 1.46.3 (stable, release)"')
  await f.stub("python3", "echo 3.9")
  const result = await f.run("tools")
  expect(result.status).toBe(0)
  expect(result.output).toContain("Deno 2+ (runs the shipped Python tool): found 1.46.3, needs 2 or newer")
  expect(result.output).toContain("Python 3.10+ (runs markitdown): found 3.9, needs 3.10 or newer")
  expect(result.output).toContain("Tesseract (OCR for scanned PDFs): not found")
  expect(result.output).toContain("Run ./start.sh in a terminal to install them")
  expect(result.output).toContain("TOOLS-DONE")
  expect((await f.calls()).split("\n").filter(line => !line.startsWith("deno --version") && !line.startsWith("python3 ") && line)).toEqual([])
  expect(await Bun.file(path.join(f.runtime, "declined-tools")).exists()).toBe(false)
})

test("Enter installs missing tools through the package manager and continues", async () => {
  const f = await fixture()
  await f.stub("deno", '[ "$1" = --version ] && echo "deno 1.46.3 (stable, release)"')
  const result = await f.answer("tools", "")
  expect(result.status).toBe(0)
  const calls = await f.calls()
  if (process.platform === "darwin") expect(calls).toMatch(/^brew install .*tesseract/m)
  else expect(calls).toMatch(/^apt-get install -y --no-install-recommends .*tesseract-ocr/m)
  expect(calls).toContain("deno upgrade")
  expect(calls).toContain("pipx install markitdown[pdf]")
  // The stubs install nothing, so the launcher reports what remains and starts.
  expect(result.output).toContain("still missing:")
  expect(result.output).toContain("TOOLS-DONE")
})

test("declined tools are not asked about again until another tool goes missing", async () => {
  const f = await fixture()
  const declined = await f.answer("tools", "n")
  expect(declined.status).toBe(0)
  expect(declined.output).toContain("[Y/n]")
  const recorded = (await readFile(path.join(f.runtime, "declined-tools"), "utf8")).trim()
  expect(recorded.split(" ")).toContain("tesseract")

  const quiet = await f.answer("tools", "n")
  expect(quiet.output).not.toContain("[Y/n]")
  expect(quiet.output).toContain(`optional tools not installed: ${recorded} (set DOCHAUS_INSTALL_TOOLS=ask`)

  expect((await f.answer("tools", "n", { DOCHAUS_INSTALL_TOOLS: "ask" })).output).toContain("[Y/n]")
  await writeFile(path.join(f.runtime, "declined-tools"), "tesseract\n")
  expect((await f.answer("tools", "n")).output).toContain("[Y/n]")
  expect(await f.calls()).not.toMatch(/install/)
})
