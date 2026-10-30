import { expect, test } from "bun:test"
import { runConcurrencyTest } from "./concurrency-test"

test("counts only admitted overlapping requests and groups failures", async () => {
  let release!: () => void
  const barrier = new Promise<void>(resolve => { release = resolve })
  let started = 0
  const result = runConcurrencyTest("synthetic/model", 3, async observer => {
    observer.start()
    if (++started === 3) release()
    await barrier
    observer.end()
    return started-- === 3 ? { ok: false, kind: "empty" } : { ok: true }
  })
  expect(await result).toMatchObject({ requested: 3, succeeded: 2, failed: 1, peak: 3, errors: [{ kind: "empty", count: 1 }] })
})
