import { afterEach, expect, test } from "bun:test"
import { mkdtempSync, mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import os from "node:os"
import path from "node:path"
import { checkPublication, publishRelease, type Runner } from "../publish-release"
import { prepareRelease } from "../release"
import { RepoUpdater } from "../updater"

const directories: string[] = []
afterEach(() => {
  for (const directory of directories.splice(0)) rmSync(directory, { recursive: true, force: true })
})

const local: Runner = (command, root) => {
  const result = Bun.spawnSync(command, { cwd: root, stdout: "pipe", stderr: "pipe" })
  return { exitCode: result.exitCode, stdout: result.stdout.toString(), stderr: result.stderr.toString() }
}

function fixture() {
  const root = mkdtempSync(path.join(os.tmpdir(), "dochaus-publish-test-"))
  directories.push(root)
  const git = (...args: string[]) => {
    const result = local(["git", ...args], root)
    if (result.exitCode) throw new Error(result.stderr)
    return result.stdout.trim()
  }
  const write = (file: string, value: unknown) => writeFileSync(path.join(root, file), typeof value === "string" ? value : `${JSON.stringify(value)}\n`)
  git("init", "--quiet")
  git("config", "user.name", "Release Test")
  git("config", "user.email", "release@example.invalid")
  git("remote", "add", "origin", "https://github.com/example/product.git")
  write("release.json", { version: 1, name: "1.0.0", notes: "Initial release." })
  write("package.json", { version: "1.0.0" })
  git("add", ".")
  git("commit", "--quiet", "-m", "initial")
  const base = git("rev-parse", "HEAD")
  function release() {
    write("release.json", { version: 2, name: "1.1.0", notes: "A useful improvement.\n\nUse Check for updates." })
    write("package.json", { version: "1.1.0" })
    mkdirSync(path.join(root, "releases"))
    write("releases/1.1.0.md", "A useful improvement.\n\nUse Check for updates.\n")
    git("add", ".")
    git("commit", "--quiet", "-m", "release")
    const target = git("rev-parse", "HEAD")
    git("update-ref", "refs/remotes/origin/dochaus", target)
    return target
  }
  return { root, base, git, write, release }
}

function network(root: string, target: string, options: { tags?: string; existing?: unknown; lookupError?: string } = {}) {
  const writes: string[][] = []
  let publishedNotes = ""
  const run: Runner = (command, cwd) => {
    if (command[0] === "git" && command[1] === "fetch") return { exitCode: 0, stdout: "", stderr: "" }
    if (command[0] === "git" && command[1] === "ls-remote") return { exitCode: 0, stdout: options.tags ?? "", stderr: "" }
    if (command[0] === "git" && command[1] === "push") {
      writes.push(command)
      return { exitCode: 0, stdout: "", stderr: "" }
    }
    if (command[0] === "gh" && command[1] === "api") {
      if (options.existing) return { exitCode: 0, stdout: JSON.stringify(options.existing), stderr: "" }
      return { exitCode: 1, stdout: "", stderr: options.lookupError ?? "gh: Not Found (HTTP 404)" }
    }
    if (command[0] === "gh" && command[1] === "release") {
      writes.push(command)
      const file = command[command.indexOf("--notes-file") + 1]!
      publishedNotes = readFileSync(file, "utf8")
      return { exitCode: 0, stdout: "https://github.com/example/product/releases/tag/v1.1.0", stderr: "" }
    }
    if (command[0] !== "git") throw new Error(`Unexpected command: ${command[0]}`)
    return local(command, cwd)
  }
  return { run, writes, notes: () => publishedNotes, publish: () => publishRelease(root, target, "example/product", run) }
}

const released = { tag_name: "v1.1.0", draft: false, prerelease: false, body: "A useful improvement.\n\nUse Check for updates." }

test("ordinary feature commits never request publication", () => {
  const f = fixture()
  f.write("feature.txt", "An ordinary change")
  f.git("add", ".")
  f.git("commit", "--quiet", "-m", "feature")
  expect(checkPublication(f.root, f.base)).toEqual({ publish: false })
})

test("release preparation identifies exactly the updater marker, including an identical merge tree", () => {
  const f = fixture()
  const target = f.release()
  f.git("commit", "--quiet", "--allow-empty", "-m", "merge with identical tree")
  expect(checkPublication(f.root, f.base)).toEqual({ publish: true, target, tag: "v1.1.0" })
})

test("a feature added after the marker cannot be claimed as tested release content", () => {
  const f = fixture()
  f.release()
  expect(checkPublication(f.root, f.base).publish).toBe(true)
  f.write("later.txt", "Not in the release")
  f.git("add", ".")
  f.git("commit", "--quiet", "-m", "later")
  expect(() => checkPublication(f.root, f.base)).toThrow("complete tested tree")
})

test("canonical notes cannot disagree with the updater notes", () => {
  const f = fixture()
  f.release()
  f.write("releases/1.1.0.md", "Different notes\n")
  f.git("add", ".")
  f.git("commit", "--quiet", "--amend", "--no-edit")
  expect(() => checkPublication(f.root, f.base)).toThrow("notes do not match")
})

test("a new release pushes only the exact marker and publishes its committed notes", () => {
  const f = fixture()
  const target = f.release()
  const n = network(f.root, target)
  expect(n.publish()).toEqual({ tag: "v1.1.0", target, created: true })
  expect(n.writes[0]).toEqual(["git", "push", "origin", `${target}:refs/tags/v1.1.0`])
  expect(n.writes[1]).toContain("--verify-tag")
  expect(n.writes[1]).toContain("--latest")
  expect(n.notes()).toBe(`${released.body}\n`)
})

test("an existing annotated tag at the same commit makes retries idempotent", () => {
  const f = fixture()
  const target = f.release()
  const n = network(f.root, target, {
    tags: `${"a".repeat(40)}\trefs/tags/v1.1.0\n${target}\trefs/tags/v1.1.0^{}\n`,
    existing: released,
  })
  expect(n.publish()).toEqual({ tag: "v1.1.0", target, created: false })
  expect(n.writes).toEqual([])
})

test.each([false, true])("wrong remote tag target is never replaced (annotated=%s)", (annotated) => {
  const f = fixture()
  const target = f.release()
  const tags = annotated
    ? `${"a".repeat(40)}\trefs/tags/v1.1.0\n${f.base}\trefs/tags/v1.1.0^{}\n`
    : `${f.base}\trefs/tags/v1.1.0\n`
  const n = network(f.root, target, { tags })
  expect(() => n.publish()).toThrow("different commit")
  expect(n.writes).toEqual([])
})

test("a retry after tag push completes the missing GitHub release without moving the tag", () => {
  const f = fixture()
  const target = f.release()
  const n = network(f.root, target, { tags: `${target}\trefs/tags/v1.1.0\n` })
  expect(n.publish().created).toBe(true)
  expect(n.writes).toHaveLength(1)
  expect(n.writes[0]!.slice(0, 3)).toEqual(["gh", "release", "create"])
})

test("drafts and mismatching published notes require review without mutations", () => {
  const f = fixture()
  const target = f.release()
  for (const existing of [{ ...released, draft: true }, { ...released, body: "Different" }]) {
    const n = network(f.root, target, { tags: `${target}\trefs/tags/v1.1.0\n`, existing })
    expect(() => n.publish()).toThrow()
    expect(n.writes).toEqual([])
  }
})

test("an API outage does not count as a missing release", () => {
  const f = fixture()
  const target = f.release()
  const n = network(f.root, target, { lookupError: "gh: Bad Gateway (HTTP 502)" })
  expect(() => n.publish()).toThrow("Could not inspect")
  expect(n.writes).toEqual([])
})

test("a superseded release run does not publish", () => {
  const f = fixture()
  const target = f.release()
  f.git("update-ref", "refs/remotes/origin/dochaus", f.base)
  const n = network(f.root, target)
  expect(() => n.publish()).toThrow("stale run")
  expect(n.writes).toEqual([])
})

test("a mismatching gh repository is rejected before remote mutations", () => {
  const f = fixture()
  const target = f.release()
  const n = network(f.root, target)
  expect(() => publishRelease(f.root, target, "other/product", n.run)).toThrow("same GitHub repository")
  expect(n.writes).toEqual([])
})

test("daily feature batches publish once and later features do not republish", async () => {
  const f = fixture()
  for (let index = 0; index < 3; index++) {
    const before = f.git("rev-parse", "HEAD")
    f.write(`feature-${index}.txt`, `Feature ${index}`)
    f.git("add", ".")
    f.git("commit", "--quiet", "-m", `feature ${index}`)
    expect(checkPublication(f.root, before)).toEqual({ publish: false })
  }
  const beforeRelease = f.git("rev-parse", "HEAD")
  f.git("branch", "-M", "dochaus")
  const installationParent = mkdtempSync(path.join(os.tmpdir(), "dochaus-release-install-"))
  directories.push(installationParent)
  const installation = path.join(installationParent, "product")
  expect(local(["git", "clone", "--quiet", f.root, installation], f.root).exitCode).toBe(0)
  const updater = new RepoUpdater(installation, path.join(installationParent, "runtime"))
  expect((await updater.check(true)).available).toBe(false)
  f.write("draft-notes.md", `${released.body}\n`)
  prepareRelease(f.root, "minor", "draft-notes.md")
  rmSync(path.join(f.root, "draft-notes.md"))
  f.git("add", ".")
  f.git("commit", "--quiet", "-m", "prepare daily release")
  const target = f.git("rev-parse", "HEAD")
  expect(await updater.check(true)).toMatchObject({ available: true, target, latest: { name: "1.1.0", version: 2 } })
  expect(checkPublication(f.root, beforeRelease)).toEqual({ publish: true, target, tag: "v1.1.0" })
  const output = path.join(f.root, "actions-output")
  const result = Bun.spawnSync([process.execPath, path.resolve(import.meta.dir, "../publish-release.ts"), "check", "--base", beforeRelease], {
    cwd: f.root,
    env: { ...process.env, GITHUB_OUTPUT: output },
    stdout: "pipe",
    stderr: "pipe",
  })
  expect(result.exitCode).toBe(0)
  expect(JSON.parse(result.stdout.toString())).toEqual({ publish: true, target, tag: "v1.1.0" })
  expect(readFileSync(output, "utf8")).toBe(`publish=true\ntarget=${target}\ntag=v1.1.0\n`)
  rmSync(output)
  f.write("later.txt", "Next release's work")
  f.git("add", ".")
  f.git("commit", "--quiet", "-m", "next feature")
  expect(checkPublication(f.root, target)).toEqual({ publish: false })
  expect(await updater.check(true)).toMatchObject({ available: true, target, latest: { name: "1.1.0", version: 2 } })
})
