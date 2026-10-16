#!/usr/bin/env bun
import { existsSync, readFileSync, readdirSync, statSync } from "node:fs"
import path from "node:path"
import { spawnSync } from "node:child_process"
import { lintProseFrequency, lintSkill, parseSkillFile } from "../lib/skill-authoring-rules"

type Options = {
  staged: boolean
  since?: string
  verbose: boolean
  nonBlocking: boolean
  postCommit: boolean
  ignoreNoisyPaths: boolean
}

const options = parseOptions()
const errors: string[] = []
const changed = collectTargetFiles(options)

const promptFiles = changed.filter(isPromptRuntimeFile)

if (!changed.length || (options.ignoreNoisyPaths && promptFiles.length === 0)) {
  if (options.verbose) {
    console.log("doc.haus legal-runtime lint skipped: no runtime prompt files in scope")
  }
  process.exit(0)
}

if (promptFiles.length === 0) {
  if (options.verbose) {
    console.log("doc.haus legal-runtime lint: no changed legal runtime prompt files")
  }
  process.exit(0)
}

for (const file of promptFiles) {
  if (!existsSync(file)) continue
  const content = readFileSync(file, "utf8")
  if (!content) continue
  const frontmatter = extractFrontmatter(content)

  if (!frontmatter) {
    errors.push(`${file}: missing YAML/TOML-style frontmatter`)
    continue
  }

  const keys = frontmatterToKeys(frontmatter)
  if (!keys.has("description") && !isSkillFile(file)) {
    errors.push(`${file}: missing description in frontmatter`)
  }

  if (isSkillFile(file)) {
    for (const finding of lintSkill(skillInputFor(file, content))) {
      const tag = `${file}: [${finding.rule}] ${finding.message}`
      if (finding.severity === "error") errors.push(tag)
      else if (options.verbose) console.warn(`- ${tag}`)
    }
  } else {
    for (const finding of lintProseFrequency(parseSkillFile(content).body)) {
      if (options.verbose) console.warn(`- ${file}: [${finding.rule}] ${finding.message}`)
    }
  }

  if (isSubagentModeAgent(file, frontmatter)) {
    ensureSubagentHidden(file, frontmatter)
  }
}

if (!options.nonBlocking) {
  if (errors.length) reportAndExit()
  if (options.verbose) {
    console.log("doc.haus legal-runtime lint: no issues found")
  }
  process.exit(0)
}

if (!errors.length && options.verbose) {
  console.log("doc.haus legal-runtime lint: no issues found")
}

if (errors.length) {
  console.error("doc.haus legal-runtime lint found issues (non-blocking mode):")
  for (const error of errors) {
    console.error(`- ${error}`)
  }
}

process.exit(0)

function parseOptions(): Options {
  const options: Options = {
    staged: Bun.argv.includes("--staged"),
    verbose: Bun.argv.includes("--verbose"),
    nonBlocking: Bun.argv.includes("--post-commit") || Bun.argv.includes("--non-blocking"),
    postCommit: Bun.argv.includes("--post-commit"),
    ignoreNoisyPaths: Bun.argv.includes("--ignore-noisy-paths") || Bun.argv.includes("--skip-nonruntime"),
  }

  const sinceArg = Bun.argv.find((arg) => arg.startsWith("--since="))
  if (sinceArg) options.since = sinceArg.slice("--since=".length)
  if (Bun.argv.includes("--last") && !options.since) options.since = "HEAD~1..HEAD"
  if (options.postCommit && !options.since) options.since = "HEAD~1..HEAD"

  return options
}

function collectTargetFiles(options: Options) {
  let args: string[]
  if (options.staged) {
    args = ["diff", "--cached", "--name-only"]
  } else if (options.since) {
    args = ["diff", "--name-only", options.since]
  } else {
    args = ["diff", "--name-only"]
  }
  const result = git(args)
  return result
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)
}

function isPromptRuntimeFile(file: string) {
  if (!existsSync(file)) return false
  if (!file.startsWith("dochaus/")) return false
  return (
    /^dochaus\/mode\/[a-z0-9-]+\.md$/.test(file)
    || /^dochaus\/agent\/[a-z0-9-]+\.md$/.test(file)
    || /^dochaus\/command\/[a-z0-9-]+\.md$/.test(file)
    || /^dochaus\/skill\/[a-z0-9-]+\/SKILL\.md$/.test(file)
  )
}

function isSubagentModeAgent(file: string, frontmatter: string) {
  if (!file.startsWith("dochaus/agent/") || !file.endsWith(".md")) return false
  return /\nmode:\s*subagent\s*(\n|$)/.test(frontmatter)
}

function ensureSubagentHidden(file: string, frontmatter: string) {
  if (!/\nhidden:\s*true\s*(\n|$)/.test(frontmatter)) {
    errors.push(`${file}: subagent must set hidden: true`)
  }
}

function extractFrontmatter(content: string) {
  const match = content.match(/^---\n([\s\S]*?)\n---/)
  return match ? match[1] : null
}

function frontmatterToKeys(frontmatter: string) {
  const entries = frontmatter
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line && !line.startsWith("#"))
  return new Set(entries.map((line) => {
    const key = line.match(/^([A-Za-z0-9_-]+)\s*:/)?.[1]
    return key ?? ""
  }))
}

function normalizeDraftModeForLint(content: string) {
  const normalized = content.replace(/\r\n/g, "\n").trim()
  const match = normalized.match(/^---\n([\s\S]*?)\n---\n?([\s\S]*)$/)
  if (!match) return normalized
  const frontmatter = match[1]
    .split("\n")
    .map((line) => line.trimEnd())
    .filter((line) => {
      const trimmed = line.trim()
      return trimmed !== "hidden: true" && trimmed !== "hidden: false"
    })
    .join("\n")
    .trim()
  const body = match[2].trim()
  return `${frontmatter}\n---\n${body}`
}

function reportAndExit() {
  console.error([
    "doc.haus legal-runtime lint failed:",
    ...errors.map((error) => `- ${error}`),
    "",
    "This guard enforces reusable-flow structure and conservative mode/agent constraints.",
    "Use --non-blocking to review issues without blocking commits.",
  ].join("\n"))
  process.exit(1)
}

function isSkillFile(file: string) {
  return /^dochaus\/skill\/[a-z0-9-]+\/SKILL\.md$/.test(file)
}

function skillInputFor(file: string, content: string) {
  const parsed = parseSkillFile(content)
  const name = file.match(/^dochaus\/skill\/([a-z0-9-]+)\/SKILL\.md$/)?.[1] ?? ""
  return {
    name,
    path: file,
    frontmatter: parsed.frontmatter,
    description: parsed.description,
    body: parsed.body,
    references: collectReferencePaths(file),
    builtins: collectBuiltinSkillNames(name),
  }
}

function collectReferencePaths(skillFile: string): string[] {
  const dir = path.dirname(skillFile)
  const refsDir = path.join(dir, "references")
  if (!existsSync(refsDir)) return []
  const out: string[] = []
  const walk = (root: string, currentDir: string) => {
    for (const entry of readdirSync(currentDir)) {
      const full = path.join(currentDir, entry)
      if (statSync(full).isDirectory()) walk(root, full)
      else out.push(`references/${path.relative(root, full)}`)
    }
  }
  walk(dir, refsDir)
  return out
}

function collectBuiltinSkillNames(exclude: string): Set<string> {
  const root = "dochaus/skill"
  if (!existsSync(root)) return new Set()
  const out = new Set<string>()
  for (const entry of readdirSync(root)) {
    if (entry === exclude) continue
    if (existsSync(path.join(root, entry, "SKILL.md"))) out.add(entry)
  }
  return out
}

function git(args: string[]) {
  const result = spawnSync("git", args, { encoding: "utf8" })
  if (result.status !== 0) throw new Error(result.stderr || `git ${args.join(" ")} failed`)
  return result.stdout
}
