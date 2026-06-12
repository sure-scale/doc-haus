#!/usr/bin/env bun
// Retrieval quality benchmark for the search-document tool (issue #67).
//
// Measures whether queries containing exact tokens (section numbers, defined
// terms, party names) surface the chunks containing those literals, and whether
// pure semantic queries hold up, before and after retrieval changes. Run once on
// the current code to record a baseline, change the retrieval code, run again,
// then compare:
//
//   bun eval/retrieval-eval.ts build            # ingest corpus, derive exact-token gold
//   bun eval/retrieval-eval.ts run baseline     # score current code
//   ... change retrieval ...
//   bun eval/retrieval-eval.ts run hybrid
//   bun eval/retrieval-eval.ts compare baseline hybrid
//
// The corpus is a fixed set of harvey-labs task document folders ingested into a
// throwaway matter at eval/.retrieval-corpus through the real ingest pipeline,
// so chunking/embedding match production exactly. Queries go through the real
// tool's execute() so the benchmark exercises the exact code path agents use.
//
// Gold sets:
// - eval/retrieval-gold-exact.json  — derived deterministically from corpus text
//   at build time (section numbers, defined terms, party names). A query is a
//   hit at rank r when the r-th citation's excerpt or section contains the
//   literal needle.
// - eval/retrieval-gold-semantic.json — hand-curated natural-language questions
//   with gold (document, substring) answers; guards against lexical changes
//   degrading semantic retrieval.

import { existsSync, mkdirSync, readFileSync, rmSync } from "node:fs"
import path from "node:path"
import searchDocument from "../tool/search-document"
import { ingestDocument } from "../../services/ingest/src/ingest"

const HARVEY_ROOT = process.env.HARVEY_ROOT ?? "/Users/nickwatson/Documents/GitHub/harvey-labs"
const CORPUS_TASKS = [
  "corporate-ma/analyze-change-of-control-provisions-across-targets-material-contracts",
  "banking-finance/compare-credit-agreement-against-term-sheet",
  "intellectual-property/triage-counterparty-redlines-to-company-dpa-template",
  "employment-labor/offer-letter-to-employment-agreement",
]
const CORPUS_DIR = path.join(import.meta.dir, ".retrieval-corpus")
const RESULTS_DIR = path.join(import.meta.dir, "results")
const GOLD_EXACT = path.join(import.meta.dir, "retrieval-gold-exact.json")
const GOLD_SEMANTIC = path.join(import.meta.dir, "retrieval-gold-semantic.json")
const K = 5

type ExactGold = { query: string; needle: string; kind: "section" | "term" | "party" }
type SemanticGold = { query: string; doc: string; substring: string }
type Citation = { documentName: string; section: string; excerpt: string; score: number }

const command = Bun.argv[2]
if (command === "build") await build()
else if (command === "run") await run(Bun.argv[3])
else if (command === "compare") compare(Bun.argv[3], Bun.argv[4])
else {
  console.error("usage: bun eval/retrieval-eval.ts build | run <label> | compare <a> <b>")
  process.exit(1)
}

async function build() {
  rmSync(CORPUS_DIR, { recursive: true, force: true })
  mkdirSync(CORPUS_DIR, { recursive: true })
  for (const task of CORPUS_TASKS) {
    const docsDir = path.join(HARVEY_ROOT, "tasks", task, "documents")
    const prefix = path.basename(task).slice(0, 24)
    for (const rel of [...new Bun.Glob("**/*.docx").scanSync({ cwd: docsDir })].sort()) {
      const name = `${prefix}--${path.basename(rel)}`
      const result = await ingestDocument(CORPUS_DIR, name, Buffer.from(await Bun.file(path.join(docsDir, rel)).arrayBuffer()))
      console.log(`ingested ${name}: ${result.chunks} chunks`)
    }
  }

  const { Database } = await import("bun:sqlite")
  const db = new Database(path.join(CORPUS_DIR, ".dochaus", "legal.db"), { readonly: true })
  const chunks = db.query("SELECT doc_name, section, text FROM chunks").all() as Array<{
    doc_name: string
    section: string
    text: string
  }>
  db.close()
  console.log(`corpus: ${chunks.length} chunks`)

  await Bun.write(GOLD_EXACT, JSON.stringify(deriveExactGold(chunks), null, 2))
  console.log(`wrote ${GOLD_EXACT}`)
  // Compact text dump for hand-curating the semantic gold set.
  await Bun.write(
    path.join(CORPUS_DIR, "chunks-dump.json"),
    JSON.stringify(chunks.map((c) => ({ doc: c.doc_name, section: c.section, text: c.text.slice(0, 700) })), null, 2),
  )
}

// Deterministic exact-token gold: every rule scans the corpus text itself, so a
// needle is guaranteed to exist in at least one chunk; retrieval failing to
// surface it is purely a ranking failure.
function deriveExactGold(chunks: Array<{ doc_name: string; section: string; text: string }>): ExactGold[] {
  const corpus = chunks.map((c) => c.text).join("\n")

  const sections = [...new Set(chunks.map((c) => c.section).filter((s) => /^\d+\.\d+$/.test(s)))].sort(
    (a, b) => a.localeCompare(b, undefined, { numeric: true }),
  )
  const sectionGold = sample(sections, 12).map((s) => ({
    query: `Section ${s}`,
    needle: s,
    kind: "section" as const,
  }))

  // Defined terms introduced as (the "Term") / ("Term") — multi-word only, so the
  // query is a phrase embeddings plausibly fumble but lexical search nails.
  const termCounts = new Map<string, number>()
  for (const m of corpus.matchAll(/\(\s*(?:the\s+|each an?\s+)?[“"]([A-Z][A-Za-z]+(?: [A-Z][A-Za-z]+){1,3})[”"]\s*\)/g))
    termCounts.set(m[1], (termCounts.get(m[1]) ?? 0) + 1)
  const terms = [...termCounts.keys()].sort()
  const termGold = sample(terms, 10).map((t) => ({ query: t, needle: t, kind: "term" as const }))

  const partySet = new Set<string>()
  for (const m of corpus.matchAll(
    /\b([A-Z][A-Za-z&.]+(?: [A-Z][A-Za-z&.]+){0,3},? (?:Holdings|LLC|LLP|Inc\.|Ltd\.|Limited|Corporation|Corp\.|GmbH))(?=[\s,.;)])/g,
  ))
    partySet.add(m[1])
  const parties = [...partySet].sort()
  const partyGold = sample(parties, 8).map((p) => ({ query: p, needle: p, kind: "party" as const }))

  return [...sectionGold, ...termGold, ...partyGold]
}

// Evenly-strided deterministic sample so the gold set spans the corpus instead
// of clustering in whatever document sorts first.
function sample<T>(items: T[], n: number): T[] {
  if (items.length <= n) return items
  return Array.from({ length: n }, (_, i) => items[Math.floor((i * items.length) / n)])
}

async function run(label: string | undefined) {
  if (!label) throw new Error("run requires a label, e.g. `run baseline`")
  if (!existsSync(GOLD_EXACT)) throw new Error("no gold set — run `build` first")
  const exact = (await Bun.file(GOLD_EXACT).json()) as ExactGold[]
  const semantic = existsSync(GOLD_SEMANTIC) ? ((await Bun.file(GOLD_SEMANTIC).json()) as SemanticGold[]) : []

  const queries = [
    ...exact.map((g) => ({ ...g, match: (c: Citation) => containsNeedle(c, g) })),
    ...semantic.map((g) => ({
      query: g.query,
      kind: "semantic" as const,
      match: (c: Citation) => c.documentName === g.doc && c.excerpt.toLowerCase().includes(g.substring.toLowerCase()),
    })),
  ]

  const rows = []
  for (const [i, q] of queries.entries()) {
    const result = await searchDocument.execute({ query: q.query, k: K }, evalContext(`retrieval-eval-${label}-${i}`))
    const citations = (typeof result === "string" ? [] : ((result.metadata?.citations ?? []) as Citation[]))
    const rank = citations.findIndex(q.match) + 1 // 0 → miss
    rows.push({ kind: q.kind, query: q.query, rank, top: citations.map((c) => `${c.documentName}§${c.section}`) })
    console.log(`${rank ? `rank ${rank}` : "MISS  "}  [${q.kind}] ${q.query}`)
  }

  mkdirSync(RESULTS_DIR, { recursive: true })
  const out = path.join(RESULTS_DIR, `retrieval-${label}.json`)
  await Bun.write(out, JSON.stringify({ label, k: K, rows }, null, 2))
  console.log(`\nwrote ${out}\n`)
  printSummary(label, rows)
}

// Section-number needles get digit boundaries so "8.3" never scores a hit off
// "18.3" or "8.30"; terms and parties are case-sensitive literals.
function containsNeedle(c: Citation, g: ExactGold) {
  const haystack = `${c.section}\n${c.excerpt}`
  if (g.kind !== "section") return haystack.includes(g.needle)
  return new RegExp(`(^|[^\\d.])${g.needle.replaceAll(".", "\\.")}($|[^\\d])`, "m").test(haystack)
}

function evalContext(sessionID: string) {
  return {
    sessionID,
    messageID: "eval",
    agent: "eval",
    directory: CORPUS_DIR,
    worktree: CORPUS_DIR,
    abort: new AbortController().signal,
    metadata: () => {},
    ask: async () => {},
  }
}

type Row = { kind: string; query: string; rank: number }

function printSummary(label: string, rows: Row[]) {
  console.log(`=== ${label} ===`)
  for (const kind of ["section", "term", "party", "semantic"]) {
    const group = rows.filter((r) => r.kind === kind)
    if (!group.length) continue
    const hit1 = group.filter((r) => r.rank === 1).length
    const hit5 = group.filter((r) => r.rank >= 1).length
    const mrr = group.reduce((s, r) => s + (r.rank ? 1 / r.rank : 0), 0) / group.length
    console.log(
      `${kind.padEnd(9)} n=${String(group.length).padEnd(3)} hit@1=${pct(hit1, group.length)} hit@${K}=${pct(hit5, group.length)} mrr=${mrr.toFixed(3)}`,
    )
  }
}

function pct(n: number, total: number) {
  return `${((100 * n) / total).toFixed(0).padStart(3)}%`
}

function compare(a: string | undefined, b: string | undefined) {
  if (!a || !b) throw new Error("compare requires two labels")
  const load = (label: string) =>
    JSON.parse(readFileSync(path.join(RESULTS_DIR, `retrieval-${label}.json`), "utf8")) as { rows: Row[] }
  const ra = load(a).rows
  const rb = load(b).rows
  printSummary(a, ra)
  printSummary(b, rb)
  console.log(`\nper-query changes (${a} → ${b}):`)
  // Keyed by query, not position: if the gold set was rebuilt between runs the
  // comparison is invalid, so drifted queries are called out instead of skipped.
  const byQuery = new Map(rb.map((r) => [`${r.kind}|${r.query}`, r]))
  for (const row of ra) {
    const after = byQuery.get(`${row.kind}|${row.query}`)
    if (!after) {
      console.log(`  [${row.kind}] ${row.query}: missing from ${b} — gold set drifted, rebuild + rerun both labels`)
      continue
    }
    if (after.rank === row.rank) continue
    const fmt = (r: number) => (r ? `rank ${r}` : "MISS")
    console.log(`  [${row.kind}] ${row.query}: ${fmt(row.rank)} → ${fmt(after.rank)}`)
  }
}
