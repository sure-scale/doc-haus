import { afterEach, describe, expect, test } from "bun:test"
import { parseRoute, quickDraftEligible, type DraftScope } from "../src/api/opencode/routing"
import { AUTO_CANDIDATES, CHAT_ASSISTANTS, CLIENT_SUMMARY, QUICK_DRAFT } from "../src/agents"
import { furnitureLegend } from "../../../dochaus/lib/furniture-legend"
import {
  QUICK_BANNER,
  quickDraftAnswer,
  quickDraftLimit,
  rememberChoice,
  rememberedChoice,
  routeChoice,
  routeForm,
  routeRequest,
  upgradeInput,
  withMatterSources,
} from "../src/components/chat/quick-draft"
import { toTurns, type Turn } from "../src/components/chat/turn-model"

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

  test("a matter with three or fewer source documents can only need few; more leaves the router's judgment", () => {
    const many: DraftScope = { ...ELIGIBLE, sources: "many" }
    for (const count of [0, 1, 3]) {
      expect(withMatterSources(many, count)).toEqual(ELIGIBLE)
      expect(quickDraftEligible(withMatterSources(many, count))).toBe(true)
      expect(quickDraftLimit(withMatterSources(many, count))).toBeUndefined()
    }
    expect(withMatterSources(many, 4)).toEqual(many)
    expect(withMatterSources(many, undefined)).toEqual(many)
    // The count settles sources only; every other field stays the router's.
    expect(quickDraftEligible(withMatterSources({ ...many, legal_effect: "yes" }, 1))).toBe(false)
    expect(withMatterSources(undefined, 1)).toBeUndefined()
  })
})

describe("candidates", () => {
  test("Quick draft is a picker entry Auto never routes to", () => {
    expect(CHAT_ASSISTANTS.some((a) => a.name === QUICK_DRAFT)).toBe(true)
    expect(NAMES).not.toContain(QUICK_DRAFT)
    expect(NAMES).toContain("draft")
    expect(NAMES.length).toBe(CHAT_ASSISTANTS.length - 2)
  })

  test("Client summary is a picker entry Auto never routes to", () => {
    expect(CHAT_ASSISTANTS.find((a) => a.name === CLIENT_SUMMARY)?.label).toBe("Client summary")
    expect(NAMES).not.toContain(CLIENT_SUMMARY)
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

  test("the form line is not working-draft text", () => {
    const turns = [
      turn("user", "Draft a covering email on the two invoices.", undefined, "u1"),
      turn("assistant", "Form: advice letter\n\nDear client, both remain unpaid.", QUICK_DRAFT, "a1"),
    ]
    expect(upgradeInput(turns, 1)).toEqual({ request: "Draft a covering email on the two invoices.", workingDraft: "Dear client, both remain unpaid." })
  })

  test("the banner is one line naming what did not run", () => {
    expect(QUICK_BANNER).toBe("Quick draft — full drafting checks not run.")
  })
})

describe("quick answer", () => {
  const LEGENDS = { agreementLegend: "Subject to contract", adviceLegend: "Privileged & confidential" }

  test("its form line picks the firm's legend, as a rendered document's form does; the line is not draft text", () => {
    for (const line of ["Form: advice letter", "**Form:** advice letter", "**Form: Advice letter**"]) {
      const answer = quickDraftAnswer(`${line}\n\nDear client,\n\nBoth invoices remain unpaid.`)
      expect(answer.form).toBe("advice letter")
      expect(answer.text).toBe("Dear client,\n\nBoth invoices remain unpaid.")
    }
    expect(furnitureLegend("advice letter", LEGENDS)).toBe("Privileged & confidential")
    expect(furnitureLegend("advice", LEGENDS)).toBe("Privileged & confidential")
    expect(furnitureLegend("letter", LEGENDS)).toBe("")
    expect(furnitureLegend("other", LEGENDS)).toBe("")
    // Legends are off by default: no firm legend, no legend.
    expect(furnitureLegend("advice letter", { agreementLegend: "", adviceLegend: "" })).toBe("")
    // No form line: no form, so no legend, and the text stands as written.
    expect(quickDraftAnswer("Dear client,\n\nText.")).toEqual({ form: undefined, text: "Dear client,\n\nText.", placeholders: [] })
  })

  test("bracketed placeholders left in the draft are its open items; citations, links and the lists after it are not", () => {
    const answer = [
      "Form: letter",
      "",
      "Dear [Client contact],",
      "",
      "The cap is 150% of the fees [Cedar Vale proposal.docx § 5]. See [the portal](https://example.test). We meet on [open item: meeting date].",
      "",
      "Kind regards, [Client contact]",
      "",
      "**Sources:**",
      "- cap at 150% [Cedar Vale proposal.docx § 5]",
      "",
      "Open items:",
      "None — [Signatory] confirmed.",
    ].join("\n")
    expect(quickDraftAnswer(answer).placeholders).toEqual(["[Client contact]", "[open item: meeting date]"])
    expect(quickDraftAnswer("Dear client,\n\nSee [Letter.docx § 1].\n\nSources:\n- [Letter.docx § 1]").placeholders).toEqual([])
  })
})

describe("quick text", () => {
  const text = (messageID: string, value: string) => ({ type: "text" as const, id: `${messageID}:text`, sessionID: "ses_1", messageID, text: value })
  const tool = (messageID: string) => ({
    type: "tool" as const,
    id: `${messageID}:tool`,
    sessionID: "ses_1",
    messageID,
    tool: "search-document",
    state: { status: "completed" as const, input: { query: "payment terms" }, output: "", title: "search-document" },
  })
  const NARRATION = "Quick draft on your supplier comparison — I'll pull the caps and payment terms."
  const LATER_NARRATION = "I've got the key figures — now I'll verify the details."
  const DRAFT = "Dear client,\n\nThe caps and payment terms are below."
  const conversation = (agent: string, steps: { info: { id: string; role: "assistant" }; parts: ReturnType<typeof text | typeof tool>[] }[]) =>
    toTurns([{ info: { id: "msg_u", role: "user", agent }, parts: [text("msg_u", "Draft a short email on the caps.")] }, ...steps], "")

  test("a quick draft's text leaves out what it wrote in steps that call a tool; the working draft follows", () => {
    const steps = [
      { info: { id: "msg_1", role: "assistant" as const }, parts: [text("msg_1", NARRATION), tool("msg_1")] },
      { info: { id: "msg_2", role: "assistant" as const }, parts: [text("msg_2", LATER_NARRATION), tool("msg_2")] },
      { info: { id: "msg_3", role: "assistant" as const }, parts: [text("msg_3", DRAFT)] },
    ]
    const turns = conversation(QUICK_DRAFT, steps)
    expect(turns[1]!.text).toBe(DRAFT)
    expect(turns[1]!.steps.length).toBeGreaterThan(0)
    expect(upgradeInput(turns, 1)?.workingDraft).toBe(DRAFT)
    // A client summary is copied out too, so it drops narration the same way.
    expect(conversation(CLIENT_SUMMARY, steps)[1]!.text).toBe(DRAFT)
    // Other assistants keep each step's text, as before.
    expect(conversation("qa", steps)[1]!.text).toBe(`${NARRATION}\n\n${LATER_NARRATION}\n\n${DRAFT}`)
  })

  test("an answer with no tool call is all answer; a turn with text only in working steps keeps it", () => {
    expect(conversation(QUICK_DRAFT, [{ info: { id: "msg_1", role: "assistant" as const }, parts: [text("msg_1", DRAFT)] }])[1]!.text).toBe(DRAFT)
    expect(conversation(QUICK_DRAFT, [{ info: { id: "msg_1", role: "assistant" as const }, parts: [text("msg_1", NARRATION), tool("msg_1")] }])[1]!.text).toBe(NARRATION)
  })
})
