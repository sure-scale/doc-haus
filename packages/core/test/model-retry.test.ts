import { expect } from "bun:test"
import { AIError, RateLimitError, TransportError } from "@opencode/ai/effect"
import { ModelRetry } from "@opencode/core/model-retry"
import { Effect, Exit, Fiber, Queue } from "effect"
import { TestClock } from "effect/testing"
import { it } from "./lib/effect"

const capacity = new AIError({ reason: new RateLimitError({ message: "Temporary capacity exhausted", retryAfterMs: 5_000 }) })
const reset = new AIError({ reason: new TransportError({ message: "Socket reset", transport: "http", operation: "read", code: "ECONNRESET" }) })

it.effect("buffered native work survives sustained capacity refusals and returns only its successful result", () =>
  Effect.gen(function* () {
    const called = yield* Queue.unbounded<number>()
    let count = 0
    const run = yield* Effect.suspend(() => {
      count++
      return Queue.offer(called, count).pipe(Effect.andThen(count <= 12 ? Effect.fail(capacity) : Effect.succeed("Finished")))
    }).pipe(ModelRetry.retry, Effect.forkChild({ startImmediately: true }))
    for (let i = 1; i <= 12; i++) {
      expect(yield* Queue.take(called)).toBe(i)
      yield* TestClock.adjust("12 seconds")
    }
    expect(yield* Fiber.join(run)).toBe("Finished")
    expect(count).toBe(13)
  }),
)

it.effect("interrupting a capacity wait stops the same pending work without another attempt", () =>
  Effect.gen(function* () {
    const called = yield* Queue.unbounded<number>()
    let count = 0
    const run = yield* Effect.suspend(() => {
      count++
      return Queue.offer(called, count).pipe(Effect.andThen(Effect.fail(capacity)))
    }).pipe(ModelRetry.retry, Effect.forkChild({ startImmediately: true }))
    yield* Queue.take(called)
    yield* TestClock.adjust("4999 millis")
    expect(count).toBe(1)
    yield* Fiber.interrupt(run)
    yield* TestClock.adjust("1 hour")
    expect(Exit.hasInterrupts(yield* Fiber.await(run))).toBe(true)
    expect(count).toBe(1)
  }),
)

it.effect("one-shot capacity refusals retain the three-retry billing bound", () =>
  Effect.gen(function* () {
    const called = yield* Queue.unbounded<number>()
    let count = 0
    const run = yield* Effect.suspend(() => {
      count++
      return Queue.offer(called, count).pipe(Effect.andThen(Effect.fail(capacity)))
    }).pipe(ModelRetry.retryOneShot, Effect.flip, Effect.forkChild({ startImmediately: true }))
    yield* Queue.take(called)
    yield* TestClock.adjust("2 minutes")
    expect(yield* Fiber.join(run)).toBe(capacity)
    expect(count).toBe(4)
  }),
)

it.effect("twelve typed socket resets retain native work until its thirteenth attempt succeeds", () =>
  Effect.gen(function* () {
    let count = 0
    const run = yield* Effect.suspend(() => {
      count++
      return count <= 12 ? Effect.fail(reset) : Effect.succeed("Recovered")
    }).pipe(ModelRetry.retry, Effect.forkChild({ startImmediately: true }))
    yield* TestClock.adjust("3 minutes")
    expect(yield* Fiber.join(run)).toBe("Recovered")
    expect(count).toBe(13)
  }),
)

it.effect("cancelling a native connection wait starts no further attempt", () =>
  Effect.gen(function* () {
    const called = yield* Queue.unbounded<number>()
    let count = 0
    const run = yield* Effect.suspend(() => {
      count++
      return Queue.offer(called, count).pipe(Effect.andThen(Effect.fail(reset)))
    }).pipe(ModelRetry.retry, Effect.forkChild({ startImmediately: true }))
    yield* Queue.take(called)
    yield* TestClock.adjust("1599 millis")
    expect(count).toBe(1)
    yield* Fiber.interrupt(run)
    yield* TestClock.adjust("1 hour")
    expect(Exit.hasInterrupts(yield* Fiber.await(run))).toBe(true)
    expect(count).toBe(1)
  }),
)

for (const delivery of [undefined, "not-sent"] as const) it.effect(`one-shot socket reset preserves billing bounds for ${delivery ?? "uncertain"} delivery`, () =>
  Effect.gen(function* () {
    let count = 0
    const failure = new AIError({ reason: new TransportError({ message: "Socket reset", transport: "http", operation: "request", code: "ECONNRESET", delivery }) })
    const run = yield* Effect.suspend(() => { count++; return Effect.fail(failure) })
      .pipe(ModelRetry.retryOneShot, Effect.flip, Effect.forkChild({ startImmediately: true }))
    yield* TestClock.adjust("2 minutes")
    expect(yield* Fiber.join(run)).toBe(failure)
    expect(count).toBe(delivery === "not-sent" ? 4 : 1)
  }),
)
