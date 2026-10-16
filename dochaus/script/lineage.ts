#!/usr/bin/env bun
// Read-only lineage query CLI (#96). Operator/dev-facing — NOT an agent tool.
// Run from the matter dir: `bun run script/lineage.ts <artifact> [--current|--manifest|--run|--verify] [--json]`
import { lastEntryFor, readLineage, resolveTier, sha256File, type LineageRow, type Tier } from "../lib/lineage"
import path from "node:path"

export function chainFor(rows: readonly LineageRow[], artifact: string): LineageRow[] {
  return rows.filter((r) => r.artifact === artifact)
}

export function currentEntry(rows: readonly LineageRow[], artifact: string): LineageRow | undefined {
  return lastEntryFor(rows, artifact)
}

export function entriesForRun(rows: readonly LineageRow[], runId: string): LineageRow[] {
  return rows.filter((r) => r.run === runId)
}

function scanTieredFiles(matterDir: string): string[] {
  const pats = ["*.md", "*.json", "*.docx", "source-chunks/*.txt", "sources/*.txt", ".dochaus/*.json", ".dochaus/*.db", ".dochaus/*.jsonl"]
  const found = new Set<string>()
  for (const pat of pats) {
    for (const p of new Bun.Glob(pat).scanSync({ cwd: matterDir, onlyFiles: true, dot: true })) {
      found.add(p.replace(/\\/g, "/"))
    }
  }
  return [...found].filter((p) => resolveTier(p) !== "FRESH")
}

export function verifyDrift(
  matterDir: string,
  rows: readonly LineageRow[],
): {
  drifted: { artifact: string; ledger_sha: string; live_sha: string }[]
  unledgered: { artifact: string }[]
} {
  const drifted: { artifact: string; ledger_sha: string; live_sha: string }[] = []
  const unledgered: { artifact: string }[] = []
  for (const artifact of scanTieredFiles(matterDir)) {
    const live = sha256File(path.join(matterDir, artifact))
    if (!live) continue
    const last = lastEntryFor(rows, artifact)
    if (!last) unledgered.push({ artifact })
    else if (last.content_sha256 && last.content_sha256 !== live.sha256)
      drifted.push({ artifact, ledger_sha: last.content_sha256, live_sha: live.sha256 })
  }
  return { drifted, unledgered }
}

export function buildManifest(matterDir: string): { artifact: string; tier: Tier }[] {
  return scanTieredFiles(matterDir)
    .map((artifact) => ({ artifact, tier: resolveTier(artifact) }))
    .sort((a, b) => a.tier.localeCompare(b.tier) || a.artifact.localeCompare(b.artifact))
}

function fmtRow(r: LineageRow): string {
  return [
    r.ts,
    r.entry_id,
    r.op.padEnd(6),
    `run=${r.run}`,
    r.stage ? `stage=${r.stage}` : "stage=-",
    `agent=${r.subagent}`,
    r.supersedes ? `supersedes=${r.supersedes}` : "supersedes=-",
    r.content_sha256 ? `sha=${r.content_sha256.slice(0, 12)}` : "sha=-",
  ].join("  ")
}

export function main(argv: string[]): void {
  const matter = process.cwd()
  const flags = new Set(argv.filter((a) => a.startsWith("--")))
  const json = flags.has("--json")
  const positional = argv.filter((a) => !a.startsWith("--"))
  const valueOf = (name: string) => {
    const i = argv.indexOf(`--${name}`)
    return i >= 0 ? argv[i + 1] : undefined
  }
  const rows = readLineage(matter)
  const out = (text: string) => (json ? console.log(JSON.stringify(text)) : console.log(text))

  if (flags.has("--manifest")) {
    const m = buildManifest(matter)
    if (json) return void console.log(JSON.stringify(m))
    for (const { artifact, tier } of m) console.log(`${tier.padEnd(9)}  ${artifact}`)
    return
  }
  if (flags.has("--verify")) {
    const d = verifyDrift(matter, rows)
    if (json) return void console.log(JSON.stringify(d))
    if (!d.drifted.length && !d.unledgered.length) return void console.log("no drift detected")
    for (const x of d.drifted) console.log(`DRIFTED    ${x.artifact}  (ledger ${x.ledger_sha.slice(0, 12)} != live ${x.live_sha.slice(0, 12)})`)
    for (const a of d.unledgered) console.log(`UNLEDGERED ${a.artifact}`)
    return
  }
  const run = valueOf("run")
  if (run) {
    const entries = entriesForRun(rows, run)
    if (json) return void console.log(JSON.stringify(entries))
    if (!entries.length) return void out(`no entries for run ${run}`)
    for (const r of entries) console.log(fmtRow(r))
    return
  }
  const artifact = flags.has("--current") ? valueOf("current") : positional[0]
  if (!artifact) return void out("usage: lineage <artifact> | --current <artifact> | --manifest | --run <id> | --verify  [--json]")
  const chain = chainFor(rows, artifact)
  if (json) return void console.log(JSON.stringify(chain))
  if (!chain.length) return void out(`no lineage recorded for ${artifact}`)
  for (const r of chain) console.log(fmtRow(r))
  if (flags.has("--current")) {
    const live = sha256File(path.join(matter, artifact))
    const last = chain[chain.length - 1]
    if (live && last.content_sha256 === live.sha256) console.log(`(live file matches latest entry)`)
    else if (live) console.log(`(LIVE FILE DIFFERS from latest ledger entry — external edit)`)
  }
}

if (import.meta.main) main(process.argv.slice(2))
