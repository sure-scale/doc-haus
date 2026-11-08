import { describe, expect } from "bun:test"
import { AIError, AuthenticationError, HttpContext, InvalidRequestError, ProviderInternalError, QuotaExceededError, RateLimitError, TransportError } from "@opencode/ai/effect"
import { Agent } from "@opencode/schema/agent"
import { Model } from "@opencode/schema/model"
import { Provider } from "@opencode/core/provider"
import { SessionSchema } from "@opencode/core/session/schema"
import { SessionRunnerRetry } from "@opencode/core/session/runner/retry"
import { toSessionError } from "@opencode/core/session/to-session-error"
import { Effect } from "effect"
import { it } from "./lib/effect"

const timeout = new AIError({
  reason: new TransportError({ message: "Timed out", transport: "http", operation: "request", code: "Timeout" }),
})
const internal = new AIError({ reason: new ProviderInternalError({ message: "internal" }) })
const capacity = new AIError({ reason: new RateLimitError({ message: "Temporary token capacity exhausted" }) })

const input = (cause: AIError) => ({
  cause,
  error: toSessionError(cause),
  agent: Agent.ID.make("build"),
  model: Model.Ref.make({ id: Model.ID.make("model"), providerID: Provider.ID.make("provider") }),
  hook: () => Effect.void,
  retry: SessionRunnerRetry.isRetryable(cause),
})

// Decisions for `count` consecutive failures of the same cause within one step.
const decisions = (cause: AIError, count: number) =>
  Effect.gen(function* () {
    const policy = yield* SessionRunnerRetry.policy(SessionSchema.ID.make("ses_retry"))
    const results: boolean[] = []
    for (let i = 0; i < count; i++) results.push((yield* policy(input(cause))).retry)
    return results
  })

describe("SessionRunnerRetry.policy", () => {
  it.effect("stops retrying transport timeouts after three attempts", () =>
    Effect.gen(function* () {
      expect(yield* decisions(timeout, 4)).toEqual([true, true, true, false])
    }),
  )

  it.effect("keeps the general allowance for other transient failures", () =>
    Effect.gen(function* () {
      expect(yield* decisions(internal, 11)).toEqual([...Array(10).fill(true), false])
    }),
  )

  it.effect("keeps soft capacity pending beyond the ordinary retry allowance", () =>
    Effect.gen(function* () {
      expect(yield* decisions(capacity, 25)).toEqual(Array(25).fill(true))
    }),
  )

  it.effect("capacity waits neither spend nor reset the ordinary fault and timeout allowances", () =>
    Effect.gen(function* () {
      const policy = yield* SessionRunnerRetry.policy(SessionSchema.ID.make("ses_mixed_retry"))
      const attempts: number[] = []
      for (let i = 0; i < 10; i++) {
        const wait = yield* policy(input(capacity))
        expect(wait.retry).toBe(true)
        if (wait.retry) attempts.push(wait.attempt)
        const fault = yield* policy(input(internal))
        expect(fault.retry).toBe(true)
        if (fault.retry) attempts.push(fault.attempt)
      }
      expect((yield* policy(input(internal))).retry).toBe(false)
      expect((yield* policy(input(capacity))).retry).toBe(true)
      expect(attempts).toEqual(Array.from({ length: 20 }, (_, i) => i + 2))

      const timed = yield* SessionRunnerRetry.policy(SessionSchema.ID.make("ses_mixed_timeout"))
      for (let i = 0; i < 3; i++) {
        expect((yield* timed(input(timeout))).retry).toBe(true)
        expect((yield* timed(input(capacity))).retry).toBe(true)
      }
      expect((yield* timed(input(timeout))).retry).toBe(false)
    }),
  )

  it.effect("capacity uses bounded backoff and honours provider retry-after minimum and cap", () =>
    Effect.gen(function* () {
      const policy = yield* SessionRunnerRetry.policy(SessionSchema.ID.make("ses_capacity_delay"))
      for (let i = 0; i < 15; i++) {
        const result = yield* policy(input(capacity))
        expect(result.retry).toBe(true)
        if (!result.retry) throw new Error("Capacity unexpectedly stopped")
        expect(result.delay).toBeGreaterThan(0)
        expect(result.delay).toBeLessThanOrEqual(12_000)
      }
      for (const [requested, expected] of [[20_000, 20_000], [3_600_000, 900_000]]) {
        const result = yield* policy(input(new AIError({ reason: new RateLimitError({ message: "Wait", retryAfterMs: requested }) })))
        expect(result).toMatchObject({ retry: true, delay: expected })
      }
    }),
  )

  it.effect("never converts explicit refusals or permanent failures into capacity waits", () =>
    Effect.gen(function* () {
      const policy = yield* SessionRunnerRetry.policy(SessionSchema.ID.make("ses_refused_retry"))
      expect((yield* policy({ ...input(capacity), retry: false })).retry).toBe(false)
      const stopped = new AIError({ reason: new RateLimitError({ message: "Stop", http: new HttpContext({ url: "https://provider.test", status: 429, headers: { "x-should-retry": "false" } }) }) })
      expect((yield* policy({ ...input(stopped), retry: true })).retry).toBe(false)
      for (const reason of [new AuthenticationError({ message: "Denied" }), new QuotaExceededError({ message: "Credit exhausted" }), new InvalidRequestError({ message: "Invalid payload" })])
        expect((yield* policy(input(new AIError({ reason })))).retry).toBe(false)
    }),
  )
})
