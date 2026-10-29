import { expect, test } from "bun:test"
import { dashboardRefreshController } from "../src/components/dashboard/data"

const tick = () => new Promise<void>((resolve) => setTimeout(resolve, 0))

test("invalidations coalesce, do not overlap, and discard superseded results", async () => {
  const pending: Array<(value: number) => void> = []
  const values: number[] = []
  let reads = 0
  const reader = dashboardRefreshController({
    load: () => { reads++; return new Promise<number>((resolve) => pending.push(resolve)) },
    pending: () => {}, publish: (value) => values.push(value), visible: () => true,
  })
  reader.request()
  reader.request()
  reader.request()
  expect(reads).toBe(1)
  pending.shift()!(1)
  await tick()
  expect(values).toEqual([])
  expect(reads).toBe(2)
  pending.shift()!(2)
  await tick()
  expect(values).toEqual([2])
  reader.request() // manual or changed inputs
  expect(reads).toBe(3)
  reader.stop()
  pending.shift()!(3)
  await tick()
  expect(values).toEqual([2])
})

test("polling pauses when hidden, resumes immediately when visible, and stops cleanly", async () => {
  let visible = true
  let reads = 0
  const reader = dashboardRefreshController({
    load: async () => ++reads, pending: () => {}, publish: () => {},
    visible: () => visible, intervalMs: 15,
  })
  reader.request()
  await tick()
  visible = false
  await new Promise((resolve) => setTimeout(resolve, 30))
  expect(reads).toBe(1)
  reader.onVisible()
  expect(reads).toBe(1)
  visible = true
  reader.onVisible()
  await tick()
  expect(reads).toBe(2)
  reader.stop()
  reader.onVisible()
  await new Promise((resolve) => setTimeout(resolve, 30))
  expect(reads).toBe(2)
})
