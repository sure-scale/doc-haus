import { appendFileSync, mkdtempSync, rmSync, writeFileSync } from "node:fs"
import os from "node:os"
import path from "node:path"
import { parseRelease } from "./release"

export type CommandResult = { exitCode: number; stdout: string; stderr: string }
export type Runner = (command: string[], root: string) => CommandResult
export type Publication = { publish: false } | { publish: true; target: string; tag: string }

const execute: Runner = (command, root) => {
  const result = Bun.spawnSync(command, { cwd: root, stdout: "pipe", stderr: "pipe" })
  return { exitCode: result.exitCode, stdout: result.stdout.toString(), stderr: result.stderr.toString() }
}

function command(run: Runner, root: string, ...args: string[]) {
  const result = run(args, root)
  // Git error output can include credential-bearing remote URLs.
  if (result.exitCode !== 0) throw new Error(`${args[0]} ${args[1]} failed. Check authentication and repository state.`)
  return result.stdout.trim()
}

function commit(run: Runner, root: string, ref: string) {
  return command(run, root, "git", "rev-parse", "--verify", "--end-of-options", `${ref}^{commit}`)
}

function candidate(root: string, target: string, run: Runner) {
  if (!/^[a-f0-9]{40,64}$/.test(target)) throw new Error("The release target must be a full commit SHA.")
  command(run, root, "git", "merge-base", "--is-ancestor", target, "HEAD")
  const marker = command(run, root, "git", "log", "-1", "--format=%H", "HEAD", "--", "release.json")
  if (marker !== target) throw new Error("The target must be the updater's latest release-marker commit.")
  const releaseTree = command(run, root, "git", "rev-parse", `${target}^{tree}`)
  const testedTree = command(run, root, "git", "rev-parse", "HEAD^{tree}")
  if (releaseTree !== testedTree)
    throw new Error("The release marker must contain the complete tested tree. Prepare it after all release changes.")
  const release = parseRelease(command(run, root, "git", "show", `${target}:release.json`))
  const pkg = JSON.parse(command(run, root, "git", "show", `${target}:package.json`))
  if (pkg.version !== release.name) throw new Error("The committed package version does not match the release.")
  const notes = command(run, root, "git", "show", `${target}:releases/${release.name}.md`)
  if (notes !== release.notes.trim()) throw new Error("The committed release notes do not match release.json.")
  return { release, notes, tag: `v${release.name}` }
}

export function checkPublication(root: string, base: string, run: Runner = execute): Publication {
  const previous = commit(run, root, base)
  const before = command(run, root, "git", "show", `${previous}:release.json`)
  const after = command(run, root, "git", "show", "HEAD:release.json")
  if (before === after) return { publish: false }
  const prior = parseRelease(before)
  const current = parseRelease(after)
  const left = current.name.split(".").map(Number)
  const right = prior.name.split(".").map(Number)
  const index = left.findIndex((value, i) => value !== right[i])
  if (index < 0 || left[index]! < right[index]! || current.version !== prior.version + 1)
    throw new Error("A release needs a higher semantic version and exactly one counter increment.")
  const target = command(run, root, "git", "log", "-1", "--format=%H", "HEAD", "--", "release.json")
  const { tag } = candidate(root, target, run)
  return { publish: true, target, tag }
}

export function publishRelease(root: string, target: string, repository: string, run: Runner = execute) {
  if (!/^[\w.-]+\/[\w.-]+$/.test(repository)) throw new Error("Set GH_REPO to the GitHub owner/repository.")
  const { release, notes, tag } = candidate(root, target, run)
  const remote = command(run, root, "git", "remote", "get-url", "origin")
  const expected = [
    `https://github.com/${repository}`,
    `https://github.com/${repository}.git`,
    `git@github.com:${repository}.git`,
    `ssh://git@github.com/${repository}.git`,
  ]
  if (!expected.includes(remote)) throw new Error("Git origin and GH_REPO must identify the same GitHub repository.")
  command(run, root, "git", "fetch", "--no-tags", "--", "origin", "refs/heads/dochaus:refs/remotes/origin/dochaus")
  const current = command(run, root, "git", "log", "-1", "--format=%H", "refs/remotes/origin/dochaus", "--", "release.json")
  if (current !== target) throw new Error("A different release is now current on dochaus; do not publish this stale run.")

  const remoteTags = command(run, root, "git", "ls-remote", "--tags", "origin", `refs/tags/${tag}`, `refs/tags/${tag}^{}`)
  const refs = new Map(remoteTags.split("\n").filter(Boolean).map((line) => {
    const [sha, ref] = line.split(/\s+/)
    return [ref!, sha!] as const
  }))
  const remoteTarget = refs.get(`refs/tags/${tag}^{}`) ?? refs.get(`refs/tags/${tag}`)
  if (remoteTarget && remoteTarget !== target) throw new Error(`${tag} already identifies a different commit; never move a release tag.`)

  const existing = run(["gh", "api", `repos/${repository}/releases/tags/${tag}`], root)
  if (existing.exitCode === 0) {
    const found = JSON.parse(existing.stdout)
    if (!remoteTarget || found.tag_name !== tag || found.draft || found.prerelease)
      throw new Error("An existing draft, prerelease or missing tag needs review before publication.")
    if (typeof found.body !== "string" || found.body.trim() !== notes)
      throw new Error("The published release notes differ from the approved notes; review them before retrying.")
    return { tag, target, created: false }
  }
  if (!/\(HTTP 404\)/.test(existing.stderr)) throw new Error("Could not inspect the GitHub release; no tag or release was created.")

  if (!remoteTarget) command(run, root, "git", "push", "origin", `${target}:refs/tags/${tag}`)
  const temporary = mkdtempSync(path.join(os.tmpdir(), "dochaus-release-"))
  try {
    const file = path.join(temporary, "notes.md")
    writeFileSync(file, `${notes}\n`)
    command(run, root, "gh", "release", "create", tag, "--repo", repository, "--verify-tag", "--title", `doc.haus ${release.name}`, "--notes-file", file, "--latest")
  } finally {
    rmSync(temporary, { recursive: true, force: true })
  }
  return { tag, target, created: true }
}

if (import.meta.main) {
  try {
    const [operation, option, value, ...extra] = process.argv.slice(2)
    if (extra.length || !value) throw new Error("Use check --base REF or publish --target FULL_SHA.")
    if (operation === "check" && option === "--base") {
      const result = checkPublication(process.cwd(), value)
      if (process.env.GITHUB_OUTPUT) {
        const lines = Object.entries(result).map(([key, entry]) => `${key}=${entry}`).join("\n")
        appendFileSync(process.env.GITHUB_OUTPUT, `${lines}\n`)
      }
      console.log(JSON.stringify(result))
    } else if (operation === "publish" && option === "--target") {
      console.log(JSON.stringify(publishRelease(process.cwd(), value, process.env.GH_REPO ?? "")))
    } else throw new Error("Use check --base REF or publish --target FULL_SHA.")
  } catch (error) {
    console.error(error instanceof Error ? error.message : "Release publication failed.")
    process.exitCode = 1
  }
}
