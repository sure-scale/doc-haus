import { afterEach, expect, test } from "bun:test"
import { saveCandidate, SaveUnconfirmedError, testCandidate } from "../src/api/opencode/provider-candidate"

const original = globalThis.fetch
afterEach(() => { globalThis.fetch = original })

test("a lost Save response cannot invite another credential activation", async () => {
  globalThis.fetch = (async () => { throw new TypeError("Network disconnected") }) as unknown as typeof fetch
  await expect(saveCandidate("draft", "primary", "fast")).rejects.toBeInstanceOf(SaveUnconfirmedError)
})

test("a confirmed rollback permits retry; partial state requires inspection", async () => {
  globalThis.fetch = (async () => Response.json({ error: "Previous connection restored.", code: "connection_save_failed", partialState: false }, { status: 503 })) as unknown as typeof fetch
  let failure: unknown
  try { await saveCandidate("draft", "primary", "fast") } catch (error) { failure = error }
  expect(failure).toBeInstanceOf(Error)
  expect(failure).not.toBeInstanceOf(SaveUnconfirmedError)
  globalThis.fetch = (async () => Response.json({ error: "The previous connection could not be restored.", code: "setup_partial_state", partialState: true }, { status: 503 })) as unknown as typeof fetch
  await expect(saveCandidate("draft", "primary", "fast")).rejects.toBeInstanceOf(SaveUnconfirmedError)
})

test("successful Save remains successful when candidate cleanup is pending", async () => {
  globalThis.fetch = (async () => Response.json({ ok: true, sessionClosePending: true })) as unknown as typeof fetch
  expect(await saveCandidate("draft", "primary", "fast")).toEqual({ ok: true, sessionClosePending: true })
  globalThis.fetch = (async () => Response.json({})) as unknown as typeof fetch
  await expect(saveCandidate("draft", "primary", "fast")).rejects.toBeInstanceOf(SaveUnconfirmedError)
})

test("expected model rejection and interrupted transport remain different recovery outcomes", async () => {
  const rejected = { ok: false, kind: "model", error: "Choose an available model." } as const
  globalThis.fetch = (async () => Response.json(rejected, { status: 400 })) as unknown as typeof fetch
  expect(await testCandidate("draft", "fast")).toEqual(rejected)
  globalThis.fetch = (async () => { throw new TypeError("Network disconnected") }) as unknown as typeof fetch
  expect(await testCandidate("draft", "fast")).toMatchObject({ ok: false, kind: "unavailable" })
})
