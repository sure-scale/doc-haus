import { Database } from "bun:sqlite"
import { afterEach, expect, spyOn, test } from "bun:test"
import path from "node:path"
import { monitorEventLoopDelay } from "node:perf_hooks"
import { readLeanEvents } from "../../../dochaus/workflows/lean/events"
import { enqueueLeanRequest } from "../../../dochaus/workflows/lean/requests"
import type { LeanDeps } from "../../../dochaus/workflows/lean/stages"
import { SOURCE, SOURCE_TEXT, cleanupMatters, happyPath, matter, scriptedModel } from "../../../dochaus/workflows/lean/test-support"
import { leanDraftingIdle, leanDraftingRoutes } from "./lean-drafting"

// Model-free responsiveness baseline for the in-process lean coordinator. A
// scripted run (each reply held for a short sleep) works over a large synthetic
// source while the state route is probed about every 50 ms on the same event
// loop. It reports route latency, event-loop delay and how many times chunk text
// is read in full. Measurement only: it asserts no latency budget.
//
// Default size is small. Scale up with, from services/ingest:
//   DOCHAUS_LEAN_LATENCY_MB=1,4,16 bun test src/lean-drafting-latency.test.ts --timeout 1800000
// DOCHAUS_LEAN_LATENCY_MODEL_MS (default 25) sets each scripted reply's hold.

const SIZES_MB = (process.env.DOCHAUS_LEAN_LATENCY_MB ?? "0.25").split(",").map(Number).filter(size => size > 0)
const MODEL_MS = Number(process.env.DOCHAUS_LEAN_LATENCY_MODEL_MS ?? 25)
const PROBE_MS = 50
const CHUNK_CHARS = 2_000
const WORDS = ["alpha", "bravo", "charlie", "delta", "echo", "foxtrot", "golf", "hotel", "india", "juliet", "kilo", "lima", "mike", "november", "oscar", "papa"]
// Statements that read chunk text, the full-scan shape (all chunks, or all of one document).
const TEXT_SCAN = /^\s*SELECT\b[^;]*\btext\b[^;]*\bFROM chunks\b/i

afterEach(cleanupMatters)

/** Generated filler only; the fixture's supply terms are the last chunk, so a
 * quotation check walks every chunk before it finds them. */
function syntheticSource(directory: string, bytes: number) {
  const db = new Database(path.join(directory, ".dochaus", "legal.db"))
  try {
    db.run("DELETE FROM chunks")
    const insert = db.query("INSERT INTO chunks VALUES (?, ?, ?, ?, ?)")
    db.transaction(() => {
      let at = 0
      for (let index = 0; at < bytes; index += 1) {
        const words = [`Synthetic filler paragraph ${index}.`]
        for (let word = 0; words.join(" ").length < CHUNK_CHARS; word += 1) words.push(WORDS[(index * 7 + word * 3) % WORDS.length]!)
        const text = words.join(" ")
        insert.run(SOURCE, String(index + 1), text, at, at + text.length)
        at += text.length + 1
      }
      insert.run(SOURCE, "final", SOURCE_TEXT, at, at + SOURCE_TEXT.length)
    })()
  } finally { db.close() }
}

function percentile(values: number[], rank: number) {
  const sorted = values.toSorted((a, b) => a - b)
  return sorted[Math.min(sorted.length - 1, Math.max(0, Math.ceil(rank / 100 * sorted.length) - 1))] ?? 0
}

const ms = (value: number) => Math.round(value * 10) / 10

for (const size of SIZES_MB) {
  test(`state route latency while a ${size} MB synthetic-source run checks its steps`, async () => {
    const directory = matter()
    syntheticSource(directory, Math.round(size * 1024 * 1024))
    const scripted = scriptedModel(happyPath)
    const deps: LeanDeps = { sessions: { ...scripted.port, prompt: async (input) => {
      await Bun.sleep(MODEL_MS)
      return scripted.port.prompt(input)
    } } }
    const routes = leanDraftingRoutes({ directory: () => directory, deps: () => deps })
    // An included instruction makes every status poll report its progress.
    enqueueLeanRequest(directory, { kind: "update", text: "Keep the synthetic headings short." })
    const queries = spyOn(Database.prototype, "query")
    const scans = (from: number, to = queries.mock.calls.length) =>
      queries.mock.calls.slice(from, to).filter(([sql]) => TEXT_SCAN.test(String(sql))).length
    const loop = monitorEventLoopDelay({ resolution: 10 })
    try {
      const startedAt = performance.now()
      const firstQuery = queries.mock.calls.length
      loop.enable()
      expect((await routes.request("/m1/lean/resume", { method: "POST" })).status).toBe(202)
      const run = { done: false }
      const idle = leanDraftingIdle(directory, 5).then(() => { run.done = true })
      // Each probe is due 50 ms after the previous one finished. Latency counts
      // from that due time, so a probe held behind a blocked loop shows the delay.
      const probes: { observed: number; service: number; scans: number }[] = []
      for (let due = performance.now() + PROBE_MS; !run.done;) {
        await Bun.sleep(Math.max(0, due - performance.now()))
        if (run.done) break
        const before = queries.mock.calls.length
        const sent = performance.now()
        // The state handler is synchronous: its reads happen inside this call.
        const pending = routes.request("/m1/lean/state")
        const statusScans = scans(before)
        const response = await pending
        expect(response.status).toBe(200)
        await response.arrayBuffer()
        const finished = performance.now()
        probes.push({ observed: finished - due, service: finished - sent, scans: statusScans })
        due = finished + PROBE_MS
      }
      await idle
      loop.disable()
      const elapsed = performance.now() - startedAt
      const totalScans = scans(firstQuery)
      const statusScans = probes.reduce((sum, probe) => sum + probe.scans, 0)
      const idleBefore = queries.mock.calls.length
      const idleSent = performance.now()
      const settled = await (await routes.request("/m1/lean/state")).json() as { status: string }
      const idleService = performance.now() - idleSent
      const events = readLeanEvents(directory)
      const total = (key: "wait_ms" | "check_ms", names: string[]) =>
        events.filter(event => names.includes(event.event)).reduce((sum, event) => sum + (event[key] ?? 0), 0)
      const report = {
        source_mb: size,
        run_ms: ms(elapsed),
        probes: probes.length,
        route_observed_ms: { p50: ms(percentile(probes.map(row => row.observed), 50)), p99: ms(percentile(probes.map(row => row.observed), 99)), max: ms(Math.max(0, ...probes.map(row => row.observed))) },
        route_service_ms: { p50: ms(percentile(probes.map(row => row.service), 50)), p99: ms(percentile(probes.map(row => row.service), 99)), max: ms(Math.max(0, ...probes.map(row => row.service))) },
        loop_delay_ms: { p50: ms(loop.percentile(50) / 1e6), p99: ms(loop.percentile(99) / 1e6), max: ms(loop.max / 1e6) },
        chunk_text_scans: {
          run_total: totalScans,
          coordinator: totalScans - statusScans,
          per_status_call: probes.length ? ms(statusScans / probes.length) : 0,
          per_status_call_max: Math.max(0, ...probes.map(row => row.scans)),
          settled_status_call: scans(idleBefore),
        },
        settled_status_ms: ms(idleService),
        events_ms: { model_wait: ms(total("wait_ms", ["checker"])), checker_checks: ms(total("check_ms", ["checker"])), step_boundaries: ms(total("check_ms", ["started", "skipped"])) },
      }
      console.log(`lean-latency ${JSON.stringify(report)}`)
      expect(settled.status).toBe("completed")
      expect(probes.length).toBeGreaterThan(0)
      expect(totalScans).toBeGreaterThan(0)
    } finally {
      loop.disable()
      queries.mockRestore()
      await leanDraftingIdle(directory)
    }
  }, 1_800_000)
}
