import { afterEach, describe, expect, test } from "bun:test"
import { parseRoute, quickDraftEligible, type DraftScope } from "../src/api/opencode/routing"
import { AUTO_CANDIDATES, CHAT_ASSISTANTS, QUICK_DRAFT } from "../src/agents"
import {
  QUICK_BANNER,
  quickDraftLimit,
  rememberChoice,
  rememberedChoice,
  routeChoice,
  routeForm,
  routeRequest,
  upgradeInput,
} from "../src/components/chat/quick-draft"
import type { Turn } from "../src/components/chat/turn-model"

// The quick-draft path (#282): the router's closed eligibility estimates, the
// route form, the remembered pre-selection and what an upgrade carries.

const NAMES = AUTO_CANDIDATES.map((a) => a.name)
const ELIGIBLE: DraftScope = { documents: "one", legal_effect: "no", outside_law: "no", firm_material: "no", sources: "few" }
const reply = (fields: Record<string, unknown>) => JSON.stringify({ route: "draft", ...fields })
// The route form's one choice field.
const choice = (form: ReturnType<typeof routeForm>) => form.fields[0] as { key: string; default?: unknown; options?: { value: string }[] }

describe("router reply", () => {
  test("a draft route with all five closed fields keeps its scope", () => {
    expect(parseRoute(reply(ELIGIBLE), NAMES)).toEqual({ name: "draft", scope: ELIGIBLE })
    expect(parseRoute("```json\n" + reply(ELIGIBLE) + "\n```", NAMES)).toEqual({ name: "draft", scope: ELIGIBLE })
    expect(parseRoute(JSON.stringify({ route: "qa" }), NAMES)).toEqual({ name: "qa" })
    expect(parseRoute("redliner", NAMES)).toEqual({ name: "redliner" })
  })

  test("a missing, unknown or extra scope field reads as no scope; the route stands", () => {
    const { sources: _sources, ...four } = ELIGIBLE
    for (const fields of [
      four,
      { ...ELIGIBLE, sources: "some" },
      { ...ELIGIBLE, legal_effect: "maybe" },
      { ...ELIGIBLE, legal_effect: false },
      { ...ELIGIBLE, extra: "no" },
      {},
    ])
      expect(parseRoute(reply(fields), NAMES)).toEqual({ name: "draft", scope: undefined })
  })

  test("malformed or off-list output is no route at all, never a substring match", () => {
    for (const text of [
      "",
      "draft please",
      "I would pick draft.",
      `Here you go: ${reply(ELIGIBLE)}`,
      "{route: draft}",
      JSON.stringify({ route: "quick-draft", ...ELIGIBLE }),
      JSON.stringify({ route: "drafting" }),
      JSON.stringify([{ route: "draft" }]),
      JSON.stringify({ documents: "one" }),
      "null",
    ])
      expect(parseRoute(text, NAMES)).toBeUndefined()
  })

  test("only one/no/no/no/few is eligible; each other value, and no scope, is not", () => {
    expect(quickDraftEligible(ELIGIBLE)).toBe(true)
    expect(quickDraftEligible(undefined)).toBe(false)
    const others: Partial<DraftScope>[] = [
      { documents: "several" },
      { legal_effect: "yes" },
      { outside_law: "yes" },
      { firm_material: "yes" },
      { sources: "many" },
    ]
    for (const change of others) {
      expect(quickDraftEligible({ ...ELIGIBLE, ...change })).toBe(false)
      expect(quickDraftLimit({ ...ELIGIBLE, ...change })).toBeTruthy()
    }
    expect(quickDraftLimit(undefined)).toBeTruthy()
    expect(quickDraftLimit(ELIGIBLE)).toBeUndefined()
  })
})

describe("candidates", () => {
  test("Quick draft is a picker entry Auto never routes to", () => {
    expect(CHAT_ASSISTANTS.some((a) => a.name === QUICK_DRAFT)).toBe(true)
    expect(NAMES).not.toContain(QUICK_DRAFT)
    expect(NAMES).toContain("draft")
    expect(NAMES.length).toBe(CHAT_ASSISTANTS.length - 1)
  })
})

describe("route form", () => {
  test("carries the request as route-only metadata and pre-selects the given choice", () => {
    const form = routeForm("ses_1", "Draft a covering email.", "quick")
    expect(form.metadata).toEqual({ kind: "route", route: "draft", request: "Draft a covering email." })
    expect(form.fields).toHaveLength(1)
    expect(form.fields[0]).toMatchObject({ key: "route", required: true, default: "quick" })
    expect(choice(form).options!.map((option) => option.value)).toEqual(["quick", "full"])
    expect(choice(routeForm("ses_1", "x", "full")).default).toBe("full")
    expect(routeRequest({ metadata: form.metadata })).toBe("Draft a covering email.")
    expect(routeRequest({ metadata: { kind: "question" } })).toBeUndefined()
    expect(routeChoice({ route: "full" })).toBe("full")
    expect(routeChoice({ route: "other" })).toBeUndefined()
    expect(routeChoice({})).toBeUndefined()
  })

  describe("remembered pick", () => {
    const store = new Map<string, string>()
    const saved = (globalThis as { window?: unknown }).window
    afterEach(() => {
      store.clear()
      ;(globalThis as { window?: unknown }).window = saved
    })

    test("is per matter, pre-selection only, and Quick until the lawyer picks otherwise", () => {
      ;(globalThis as { window?: unknown }).window = {
        localStorage: { getItem: (key: string) => store.get(key) ?? null, setItem: (key: string, value: string) => void store.set(key, value) },
      }
      expect(rememberedChoice("m1")).toBe("quick")
      rememberChoice("m1", "full")
      expect(rememberedChoice("m1")).toBe("full")
      expect(rememberedChoice("m2")).toBe("quick")
      expect(choice(routeForm("ses_1", "x", rememberedChoice("m1"))).default).toBe("full")
    })

    test("unreadable storage falls back to Quick", () => {
      ;(globalThis as { window?: unknown }).window = undefined
      expect(rememberedChoice("m1")).toBe("quick")
    })
  })
})

describe("upgrade", () => {
  const turn = (role: Turn["role"], text: string, agent?: string, id?: string): Turn => ({
    role,
    text,
    agent,
    id,
    citations: [],
    redlines: [],
    drafts: [],
    todos: [],
    steps: [],
  })

  test("carries the request with later instructions, and the latest quick text as the working draft", () => {
    const turns = [
      turn("user", "What is the governing law?", undefined, "u0"),
      turn("assistant", "English law.", "qa", "a0"),
      turn("user", "Draft a covering email on the two invoices.", undefined, "u1"),
      turn("assistant", "Dear client, first version.", QUICK_DRAFT, "a1"),
      turn("user", "Make it shorter.", undefined, "u2"),
      turn("assistant", "Dear client, short version.", QUICK_DRAFT, "a2"),
    ]
    expect(upgradeInput(turns, 5)).toEqual({
      request: "Draft a covering email on the two invoices.\n\nLater messages from the lawyer, in order:\n- Make it shorter.",
      workingDraft: "Dear client, short version.",
    })
    expect(upgradeInput(turns, 3)).toEqual({
      request: "Draft a covering email on the two invoices.",
      workingDraft: "Dear client, first version.",
    })
    expect(upgradeInput(turns, 1)).toBeUndefined()
    expect(upgradeInput(turns, 4)).toBeUndefined()
  })

  test("a quick answer that declined the request carries no working draft", () => {
    const turns = [
      turn("user", "Draft the termination notice.", undefined, "u1"),
      turn("assistant", "Quick draft can’t cover this request: it would end the lease.", QUICK_DRAFT, "a1"),
    ]
    expect(upgradeInput(turns, 1)).toEqual({ request: "Draft the termination notice." })
  })

  test("the banner is one line naming what did not run", () => {
    expect(QUICK_BANNER).toBe("Quick draft — full drafting checks not run.")
  })
})
