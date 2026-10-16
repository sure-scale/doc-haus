#!/usr/bin/env bun
import { spawnSync } from "node:child_process"

const runtimeLeakPatterns = (await Bun.file(
  new URL("./product-boundary/runtime-leak-patterns.json", import.meta.url),
).json()) as string[]
const staged = Bun.argv.includes("--staged")
const diffArgs = staged ? ["diff", "--cached", "--unified=0"] : ["diff", "HEAD", "--unified=0"]
const untracked = staged
  ? []
  : git(["ls-files", "--others", "--exclude-standard"])
      .split(/\r?\n/)
      .map((line) => line.trim())
      .filter(Boolean)
const diff = [git([...diffArgs, "--", "dochaus/"]), await untrackedDiff(untracked)].filter(Boolean).join("\n")

const errors = [
  ...runtimePromptErrors(addedLinesForPath(runtimeProductPathPattern())),
]

if (!errors.length) process.exit(0)

console.error(
  [
    "doc.haus product-boundary guard failed:",
    ...errors.map((error) => `- ${error}`),
    "",
    "Keep runtime product instructions independent of diagnostic fixtures.",
  ].join("\n"),
)
process.exit(1)

function runtimePromptErrors(addedLines: string[]) {
  if (!addedLines) return []

  return runtimeLeakPatterns.some((pattern) => new RegExp(pattern, "i").test(addedLines))
    ? ["runtime product diff contains benchmark-shaped or fixture-specific wording"]
    : []
}

function git(args: string[]) {
  const result = spawnSync("git", args, { encoding: "utf8", maxBuffer: 16 * 1024 * 1024 })
  if (result.status !== 0) throw new Error(result.stderr || `git ${args.join(" ")} failed`)
  return result.stdout
}

function addedLinesForPath(pattern: RegExp) {
  let current = ""
  const lines: string[] = []
  for (const line of diff.split(/\r?\n/)) {
    const header = line.match(/^diff --git a\/(.+) b\/(.+)$/)
    if (header) {
      current = header[2]
      continue
    }
    if (!pattern.test(current)) continue
    if (line.startsWith("+") && !line.startsWith("+++")) lines.push(line)
  }
  return lines.join("\n")
}

async function untrackedDiff(files: string[]) {
  const relevant = files.filter(
    (file) => runtimeProductPathPattern().test(file),
  )
  const patches = await Promise.all(
    relevant.map(async (file) => {
      const body = Bun.file(file)
      if (body.size > 2 * 1024 * 1024) return ""
      const content = await body.text().catch(() => "")
      if (!content) return ""
      return [`diff --git a/${file} b/${file}`, ...content.split(/\r?\n/).map((line) => `+${line}`)].join("\n")
    }),
  )
  return patches.filter(Boolean).join("\n")
}

function runtimeProductPathPattern() {
  return /^dochaus\/(?:agent|mode|skill|tool|plugin|lib|command|jurisdiction)\/|^dochaus\/opencode\.json$/
}
