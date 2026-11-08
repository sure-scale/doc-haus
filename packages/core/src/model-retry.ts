export * as ModelRetry from "./model-retry.js"

import { AIError, isRetryable } from "@opencode/ai/effect"
import { Clock, Duration, Effect, Pull, Schedule } from "effect"

interface Input {
  readonly cause: AIError
  readonly retry: boolean
}

/** Bound provider-requested delays so a hostile or buggy retry-after cannot stall a call for hours. */
const RETRY_AFTER_MAX = Duration.toMillis("15 minutes")

const retryAfter = (input: Input) => {
  if (input.cause.reason._tag === "RateLimit" || input.cause.reason._tag === "ProviderInternal")
    return input.cause.reason.retryAfterMs === undefined
      ? undefined
      : Math.min(input.cause.reason.retryAfterMs, RETRY_AFTER_MAX)
  return undefined
}

// Capacity waits and ordinary faults share the delay rules, not the retry count.
// Exponential from 2s capped at 10s per gap (before jitter); provider hints can
// raise the minimum, up to RETRY_AFTER_MAX. Every wait remains interruptible.
const backoff = Schedule.min([Schedule.exponential("2 seconds"), Schedule.spaced("10 seconds")]).pipe(
  Schedule.jittered,
  Schedule.setInputType<Input>(),
  Schedule.modifyDelay(({ input, duration: delay }) => {
    const minimum = retryAfter(input)
    const duration = minimum === undefined ? delay : Duration.max(delay, Duration.millis(minimum))
    return Effect.succeed(Duration.millis(Math.ceil(Duration.toMillis(duration))))
  }),
)
const schedule = Schedule.max([backoff, Schedule.recurs(10)])

// A timed-out attempt already waited minutes before failing, so the general allowance would let a
// dead provider hold a call for most of an hour. Cap those attempts well below it.
const MAX_TIMEOUT_RETRIES = 3

const makePolicy = (capacityWait: boolean) =>
  Effect.gen(function* () {
    const step = yield* Schedule.toStep(schedule)
    const capacity = yield* Schedule.toStep(backoff)
    let attempt = 1
    let timeouts = 0
    return (input: Input) =>
      Effect.gen(function* () {
        if (!input.retry || input.cause.reason.http?.headers["x-should-retry"] === "false")
          return { attempt: attempt + 1, delay: 0, retry: false }
        const now = yield* Clock.currentTimeMillis
        // Soft throttles leave the same native turn pending. They must not spend
        // the fault allowance shared by pre-output Retry and partial Continue.
        const waiting = capacityWait && input.cause.reason._tag === "RateLimit"
        const next = yield* (waiting ? capacity : step)(now, input).pipe(Pull.catchDone(() => Effect.succeed(undefined)))
        if (!next) return undefined
        const [, duration] = next
        attempt++
        if (input.cause.reason._tag === "Transport" && input.cause.reason.code === "Timeout") timeouts++
        return {
          attempt,
          delay: Math.ceil(Duration.toMillis(duration)),
          retry: waiting || timeouts <= MAX_TIMEOUT_RETRIES,
        }
      })
  })

export const policy = makePolicy(true)

/** Retry a buffered model call from scratch; failed attempts never contribute to its returned output. */
export const retry = <A, R>(effect: Effect.Effect<A, AIError, R>) =>
  Effect.gen(function* () {
    const decide = yield* policy
    while (true) {
      const result = yield* Effect.result(effect)
      if (result._tag === "Success") return result.success
      const decision = yield* decide({ cause: result.failure, retry: isRetryable(result.failure) })
      if (!decision?.retry) return yield* result.failure
      yield* Effect.sleep(Duration.millis(decision.delay))
    }
  })

// One-shot generations (stateless text, images, decisions) can be billed per attempt, so they retry only failures
// the provider never ran, a few times, with the same backoff and retry-after handling.
const ONE_SHOT_RETRIES = 3
const NOT_SENT_CODES = new Set(["ECONNREFUSED", "ENOTFOUND", "EAI_AGAIN"])

const notRun = (error: AIError) => {
  const override = error.reason.http?.headers["x-should-retry"]
  if (override === "true") return true
  if (override === "false") return false
  const reason = error.reason
  if (reason._tag === "RateLimit") return true
  if (reason._tag === "ProviderInternal") return reason.http?.status === 503 || reason.http?.status === 529
  if (reason._tag === "Transport")
    return reason.delivery === "not-sent" || (reason.code !== undefined && NOT_SENT_CODES.has(reason.code))
  return false
}

/** Retry a one-shot call only when the provider clearly didn't run it: rate limits, overload, unreachable hosts. */
export const retryOneShot = <A, R>(effect: Effect.Effect<A, AIError, R>) =>
  Effect.gen(function* () {
    // Stateless generations keep their billing bound, including soft capacity.
    const decide = yield* makePolicy(false)
    let retries = 0
    while (true) {
      const result = yield* Effect.result(effect)
      if (result._tag === "Success") return result.success
      const decision = yield* decide({ cause: result.failure, retry: notRun(result.failure) })
      if (!decision?.retry || retries++ >= ONE_SHOT_RETRIES) return yield* result.failure
      yield* Effect.sleep(Duration.millis(decision.delay))
    }
  })
