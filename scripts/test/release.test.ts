import { afterEach, expect, test } from "bun:test"
import { mkdtempSync, mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import os from "node:os"
import path from "node:path"
import { checkRelease, prepareRelease } from "../release"

const roots: string[] = []
afterEach(() => { for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true }) })

function git(root: string, ...args: string[]) {
  const result = Bun.spawnSync(["git", ...args], { cwd: root, stdout: "pipe", stderr: "pipe" })
  if (result.exitCode !== 0) throw new Error(result.stderr.toString())
  return result.stdout.toString().trim()
}

function write(root: string, file: string, value: unknown) {
  writeFileSync(path.join(root, file), `${JSON.stringify(value, null, 2)}\n`)
}

function fixture(packageVersion = "2.2.0") {
  const root = mkdtempSync(path.join(os.tmpdir(), "dochaus-release-"))
  roots.push(root)
  git(root, "init", "--quiet")
  git(root, "config", "user.email", "release@example.invalid")
  git(root, "config", "user.name", "Release test")
  write(root, "release.json", { version: 22, name: "2.2.0", notes: "Existing release" })
  write(root, "package.json", { name: "synthetic-app", version: packageVersion, scripts: { example: "preserved" } })
  git(root, "add", ".")
  git(root, "commit", "--quiet", "-m", "baseline")
  writeFileSync(path.join(root, "notes.md"), "\nNew user-facing feature.\n\nUpgrade normally.\n")
  return root
}

function marker(root: string, changes: Record<string, unknown>) {
  const file = path.join(root, "release.json")
  write(root, "release.json", { ...JSON.parse(readFileSync(file, "utf8")), ...changes })
}

test("existing marker can be reconciled without declaring a new release", () => {
  const root = fixture("2.0.19")
  expect(() => checkRelease(root)).toThrow("package.json version")
  write(root, "package.json", { version: "2.2.0" })
  expect(checkRelease(root, "HEAD").version).toBe(22)
})

test("prepare synchronizes metadata and archived notes, preserves package fields, and refuses a second bump", () => {
  const root = fixture()
  const release = prepareRelease(root, "minor", "notes.md")
  expect(release).toEqual({ version: 23, name: "2.3.0", notes: "New user-facing feature.\n\nUpgrade normally." })
  expect(checkRelease(root, "HEAD")).toEqual(release)
  expect(JSON.parse(readFileSync(path.join(root, "package.json"), "utf8"))).toMatchObject({ version: "2.3.0", scripts: { example: "preserved" } })
  expect(readFileSync(path.join(root, "releases/2.3.0.md"), "utf8")).toBe(`${release.notes}\n`)
  expect(() => prepareRelease(root, "patch", "notes.md")).toThrow("already has uncommitted changes")
})

test.each([ ["patch", "2.2.1"], ["major", "3.0.0"], ["2.4.0", "2.4.0"] ])("prepare accepts %s", (bump, expected) => {
  const root = fixture()
  expect(prepareRelease(root, bump, "notes.md").name).toBe(expected)
  expect(checkRelease(root, "HEAD").version).toBe(23)
})

test("notes-only and formatting-only marker edits cannot silently change the updater target", () => {
  const root = fixture()
  expect(checkRelease(root, "HEAD").name).toBe("2.2.0")
  marker(root, { notes: "Changed wording" })
  expect(() => checkRelease(root, "HEAD")).toThrow("higher semantic version")
  git(root, "restore", "release.json")
  writeFileSync(path.join(root, "release.json"), `${readFileSync(path.join(root, "release.json"), "utf8")}\n`)
  expect(() => checkRelease(root, "HEAD")).toThrow("higher semantic version")
})

test("checks reject skipped counters, missing or divergent notes, and package drift", () => {
  const root = fixture()
  prepareRelease(root, "patch", "notes.md")
  expect(checkRelease(root, "HEAD").name).toBe("2.2.1")
  marker(root, { version: 24 })
  expect(() => checkRelease(root, "HEAD")).toThrow("exactly once")
  marker(root, { version: 23 })
  const notes = path.join(root, "releases/2.2.1.md")
  writeFileSync(notes, "Other notes\n")
  expect(() => checkRelease(root, "HEAD")).toThrow("must match")
  rmSync(notes)
  expect(() => checkRelease(root, "HEAD")).toThrow("Missing canonical")
  write(root, "package.json", { version: "2.2.0" })
  expect(() => checkRelease(root, "HEAD")).toThrow("package.json version")
})

test.each(["2.2.0", "2.1.9", "2.3.0-rc.1", "02.3.0", "2.3", "9007199254740992.0.0"])("invalid or reused version %s fails without writes", (version) => {
  const root = fixture()
  const before = readFileSync(path.join(root, "release.json"), "utf8")
  expect(() => prepareRelease(root, version, "notes.md")).toThrow()
  expect(readFileSync(path.join(root, "release.json"), "utf8")).toBe(before)
  expect(git(root, "diff", "--stat")).toBe("")
})

test("blank notes, existing archives and existing tags fail before metadata writes", () => {
  const root = fixture()
  writeFileSync(path.join(root, "empty.md"), " \n")
  expect(() => prepareRelease(root, "patch", "empty.md")).toThrow("must not be empty")
  mkdirSync(path.join(root, "releases"))
  writeFileSync(path.join(root, "releases/2.2.1.md"), "Historical\n")
  expect(() => prepareRelease(root, "patch", "notes.md")).toThrow("already exist")
  git(root, "tag", "v2.3.0")
  expect(() => prepareRelease(root, "minor", "notes.md")).toThrow("Tag v2.3.0 already exists")
  expect(git(root, "diff", "--stat")).toBe("")
})

test("actual CLI prepares and checks a release, and reports invalid references and arguments", () => {
  const root = fixture()
  const script = path.resolve(import.meta.dir, "../release.ts")
  const cli = (...args: string[]) => Bun.spawnSync([process.execPath, script, ...args], { cwd: root, stdout: "pipe", stderr: "pipe" })
  expect(cli("prepare", "patch", "--notes", "notes.md").exitCode).toBe(0)
  const checked = cli("check", "--base", "HEAD")
  expect(checked.exitCode).toBe(0)
  expect(checked.stdout.toString()).toContain("2.2.1 (counter 23)")
  expect(cli("check", "--base", "missing-branch").exitCode).toBe(1)
  expect(cli("prepare", "patch").exitCode).toBe(1)
})

test("check rejects reused tags but permits rerunning validation of the same published marker", () => {
  const root = fixture()
  prepareRelease(root, "patch", "notes.md")
  git(root, "tag", "v2.2.1")
  expect(() => checkRelease(root, "HEAD")).toThrow("another release")
  git(root, "tag", "--delete", "v2.2.1")
  git(root, "add", "release.json", "package.json", "releases/2.2.1.md")
  git(root, "commit", "--quiet", "-m", "release")
  git(root, "tag", "v2.2.1")
  expect(checkRelease(root, "HEAD^").name).toBe("2.2.1")
})

test("an unchanged marker cannot lose its previously archived notes", () => {
  const root = fixture()
  prepareRelease(root, "patch", "notes.md")
  git(root, "add", "release.json", "package.json", "releases/2.2.1.md")
  git(root, "commit", "--quiet", "-m", "release")
  expect(checkRelease(root, "HEAD").name).toBe("2.2.1")
  rmSync(path.join(root, "releases/2.2.1.md"))
  expect(() => checkRelease(root, "HEAD")).toThrow("Missing canonical release notes")
})
