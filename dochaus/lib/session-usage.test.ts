import { expect, test } from "bun:test"
import { sessionUsageEntries, wholeSessionUsage, type UsageSession } from "./session-usage"
import { UPKEEP_AGENT, WORD_SUGGESTIONS_AGENT } from "./usage-agents"

const session: UsageSession = {
  id: "ses_usage", agent: "qa", cost: 0.05,
  tokens: { input: 100, output: 20, reasoning: 0, cache: { read: 5, write: 0 } },
  time: { created: 1_000, updated: 9_000 },
}
const step = (cost: number, created: number, extra: object = {}) =>
  ({ info: { cost, time: { created }, providerID: "fixture", modelID: "note", tokens: { input: 10, output: 2 }, ...extra } })

test("each priced model step is an entry under its agent and model; the remainder is upkeep", () => {
  expect(sessionUsageEntries(session, [step(0.02, 2_000), step(0, 3_000), step(0.01, 4_000, { agent: "lean-advise" })])).toEqual([
    { time: 2_000, cost: 0.02, session: "ses_usage", agent: "qa", model: "fixture/note", tokens: { input: 10, output: 2, reasoning: 0, cache: { read: 0, write: 0 } } },
    { time: 4_000, cost: 0.01, session: "ses_usage", agent: "lean-advise", model: "fixture/note", tokens: { input: 10, output: 2, reasoning: 0, cache: { read: 0, write: 0 } } },
    { time: 9_000, cost: expect.closeTo(0.02, 10), session: "ses_usage", agent: UPKEEP_AGENT },
  ])
})

test("a title costing under a hundredth of a cent is still upkeep", () => {
  const titled = { ...session, cost: 0.02005 }
  expect(sessionUsageEntries(titled, [step(0.02, 2_000)]).at(-1)).toEqual({ time: 9_000, cost: expect.closeTo(0.00005, 12), session: "ses_usage", agent: UPKEEP_AGENT })
})

test("a session without cost is one whole entry; Word suggestions resolve by metadata", () => {
  const unpriced = { ...session, cost: 0, metadata: { kind: "suggestions" } }
  expect(sessionUsageEntries(unpriced, [step(0.02, 2_000)])).toEqual([wholeSessionUsage(unpriced)])
  expect(wholeSessionUsage(unpriced)).toEqual({ time: 1_000, cost: 0, session: "ses_usage", agent: WORD_SUGGESTIONS_AGENT,
    tokens: { input: 100, output: 20, reasoning: 0, cache: { read: 5, write: 0 } } })
  expect(sessionUsageEntries({ ...session, metadata: { kind: "suggestions" } }, [step(0.05, 2_000)]).map(entry => entry.agent)).toEqual([WORD_SUGGESTIONS_AGENT])
})
