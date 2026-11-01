import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs"
import path from "node:path"

export type Release = { version: number; name: string; notes: string }

function git(root: string, ...args: string[]) {
  const result = Bun.spawnSync(["git", ...args], { cwd: root, stdout: "pipe", stderr: "pipe" })
  if (result.exitCode !== 0) throw new Error(`Git could not ${args[0]}. Check the repository and supplied reference.`)
  return result.stdout.toString()
}

function components(name: string): number[] {
  if (!/^(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)$/.test(name))
    throw new Error("Release name must be a stable semantic version, such as 2.3.0.")
  const parts = name.split(".").map(Number)
  if (!parts.every(Number.isSafeInteger)) throw new Error("Release version components must be safe integers.")
  return parts
}

function newer(next: string, previous: string) {
  const left = components(next)
  const right = components(previous)
  for (let index = 0; index < 3; index++) {
    if (left[index] !== right[index]) return left[index]! > right[index]!
  }
  return false
}

export function parseRelease(raw: string): Release {
  const value = JSON.parse(raw)
  if (!value || !Number.isSafeInteger(value.version) || value.version < 1)
    throw new Error("Release counter must be a positive safe integer.")
  if (typeof value.name !== "string") throw new Error("Release name is required.")
  components(value.name)
  if (typeof value.notes !== "string" || !value.notes.trim()) throw new Error("Release notes must not be empty.")
  return value
}

function read(root: string, file: string) {
  return readFileSync(path.join(root, file), "utf8")
}

function archive(root: string, release: Release, required: boolean) {
  const file = path.join(root, "releases", `${release.name}.md`)
  if (!existsSync(file)) {
    if (required) throw new Error(`Missing canonical release notes: releases/${release.name}.md`)
    return
  }
  if (readFileSync(file, "utf8") !== `${release.notes.trim()}\n`)
    throw new Error(`releases/${release.name}.md must match release.json notes.`)
}

export function checkRelease(root: string, base?: string): Release {
  const raw = read(root, "release.json")
  const current = parseRelease(raw)
  if (JSON.parse(read(root, "package.json")).version !== current.name)
    throw new Error("package.json version must match release.json name.")
  let changed = false
  let previouslyArchived = false
  if (base) {
    const revision = git(root, "rev-parse", "--verify", "--end-of-options", `${base}^{commit}`).trim()
    // Compare bytes too: even formatting-only marker edits change the updater's target.
    const previousRaw = git(root, "show", `${revision}:release.json`)
    changed = raw !== previousRaw
    previouslyArchived = Boolean(git(root, "ls-tree", "--name-only", revision, "--", `releases/${current.name}.md`).trim())
    if (changed) {
      const previous = parseRelease(previousRaw)
      if (!newer(current.name, previous.name)) throw new Error("A changed release marker requires a higher semantic version.")
      if (current.version !== previous.version + 1)
        throw new Error("A new release must increment the updater counter exactly once.")
    }
  }
  if (changed && git(root, "tag", "--list", `v${current.name}`)) {
    const tag = git(root, "rev-parse", `refs/tags/v${current.name}^{commit}`).trim()
    const marker = git(root, "log", "-1", "--format=%H", "HEAD", "--", "release.json").trim()
    if (tag !== marker || git(root, "show", `${tag}:release.json`) !== raw)
      throw new Error(`Tag v${current.name} already identifies another release; never reuse a release version.`)
  }
  archive(root, current, changed || previouslyArchived)
  return current
}

export function prepareRelease(root: string, bump: string, notesPath: string): Release {
  const raw = read(root, "release.json")
  const current = parseRelease(raw)
  if (raw !== git(root, "show", "HEAD:release.json"))
    throw new Error("release.json already has uncommitted changes. Review the prepared release before preparing another.")
  const parts = components(current.name)
  const index = ["major", "minor", "patch"].indexOf(bump)
  if (index >= 0) {
    parts[index] = parts[index]! + 1
    for (let reset = index + 1; reset < 3; reset++) parts[reset] = 0
  }
  const name = index >= 0 ? parts.join(".") : bump
  components(name)
  if (!newer(name, current.name)) throw new Error("The next release must have a higher semantic version.")
  const notes = readFileSync(path.resolve(root, notesPath), "utf8").trim()
  const next = parseRelease(JSON.stringify({ ...current, version: current.version + 1, name, notes }))
  const notesFile = path.join(root, "releases", `${name}.md`)
  if (existsSync(notesFile)) throw new Error(`Release notes already exist for ${name}; never reuse a release version.`)
  if (git(root, "tag", "--list", `v${name}`)) throw new Error(`Tag v${name} already exists; never reuse a release version.`)
  const pkg = JSON.parse(read(root, "package.json"))
  // Finish validation before writing any files. Preparation never commits or publishes.
  pkg.version = name
  mkdirSync(path.dirname(notesFile), { recursive: true })
  writeFileSync(notesFile, `${notes}\n`, { flag: "wx" })
  writeFileSync(path.join(root, "package.json"), `${JSON.stringify(pkg, null, 2)}\n`)
  writeFileSync(path.join(root, "release.json"), `${JSON.stringify(next, null, 2)}\n`)
  return next
}

if (import.meta.main) {
  try {
    const [command, ...args] = process.argv.slice(2)
    let release: Release
    if (command === "check" && (args.length === 0 || (args.length === 2 && args[0] === "--base"))) {
      release = checkRelease(process.cwd(), args[1])
      console.log(`Release ${release.name} (counter ${release.version}) is consistent.`)
    } else if (command === "prepare" && args.length === 3 && args[1] === "--notes") {
      release = prepareRelease(process.cwd(), args[0]!, args[2]!)
      console.log(`Prepared ${release.name} (counter ${release.version}). Review release.json, package.json and releases/${release.name}.md before committing.`)
    } else {
      throw new Error("Usage: bun scripts/release.ts check [--base REF] | prepare <patch|minor|major|x.y.z> --notes FILE")
    }
  } catch (error) {
    console.error(error instanceof Error ? error.message : "Release command failed.")
    process.exitCode = 1
  }
}
