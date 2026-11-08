import { expect } from "bun:test"
import { AIError, RateLimitError } from "@opencode/ai/effect"
import { ModelRetry } from "@opencode/core/model-retry"
import { Effect, Exit, Fiber, Queue } from "effect"
import { TestClock } from "effect/testing"
import { it } from "./lib/effect"

const capacity = new AIError({ reason: new RateLimitError({ message: "Temporary capacity exhausted", retryAfterMs: 5_000 }) })

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
