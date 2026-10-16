import { afterEach, describe, expect, test } from "bun:test"
import { spawnSync } from "node:child_process"
import { copyFileSync, mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"

const temporary: string[] = []
afterEach(() => {
  for (const directory of temporary.splice(0)) rmSync(directory, { recursive: true, force: true })
})

function fixture() {
  const directory = mkdtempSync(path.join(tmpdir(), "product-boundary-"))
  temporary.push(directory)
  const script = path.join(directory, "dochaus/script")
  mkdirSync(path.join(script, "product-boundary"), { recursive: true })
  for (const file of ["guard-product-boundary.ts", "product-boundary/runtime-leak-patterns.json"]) {
    copyFileSync(path.join(import.meta.dir, "../script", file), path.join(script, file))
  }
  const git = (...args: string[]) => {
    const result = spawnSync("git", args, { cwd: directory, encoding: "utf8" })
    if (result.status !== 0) throw new Error(result.stderr)
  }
  git("init", "--quiet")
  git("config", "user.email", "test@example.invalid")
  git("config", "user.name", "Synthetic Test")
  git("add", ".")
  git("-c", "core.hooksPath=/dev/null", "commit", "--quiet", "-m", "Synthetic baseline")
  const write = (file: string, body: string) => {
    mkdirSync(path.dirname(path.join(directory, file)), { recursive: true })
    writeFileSync(path.join(directory, file), body)
  }
  const run = (...args: string[]) =>
    spawnSync(process.execPath, [path.join(script, "guard-product-boundary.ts"), ...args], {
      cwd: directory,
      encoding: "utf8",
    })
  return { git, write, run }
}

describe("runtime product diff guard", () => {
  test("rejects diagnostic answer hints in untracked runtime code", () => {
    const project = fixture()
    project.write("dochaus/tool/synthetic.ts", "Use the expected answer from the evaluation dataset.\n")
    const result = project.run()
    expect(result.status).toBe(1)
    expect(result.stderr).toContain("product-boundary guard failed")
  })

  test("rejects staged prompt contamination", () => {
    const project = fixture()
    project.write("dochaus/agent/synthetic.md", "Follow the benchmark answer-key.\n")
    project.git("add", "dochaus/agent/synthetic.md")
    expect(project.run("--staged").status).toBe(1)
  })

  test("allows ordinary legal wording and diagnostic tests outside runtime paths", () => {
    const project = fixture()
    project.write("dochaus/tool/synthetic.ts", "Apply the client's criterion to the facts.\n")
    project.write("dochaus/test/synthetic.test.ts", "The expected answer is a synthetic fixture.\n")
    const result = project.run()
    expect(result.status).toBe(0)
    expect(result.stderr).toBe("")
  })
})
