// Run from apps/web: bun test test/key-dates-register.test.tsx
// Mocked ingest only: the register reads stored rows and writes nothing.
import { expect, test } from "bun:test"
import { Window } from "happy-dom"
import { act, createElement } from "react"

type Row = {
  id: number
  document: string
  kind: string
  date: string | null
  formula: string | null
  noticeDays: number | null
  actBy: string | null
  party: string | null
  quote: string
  reason: string
  author: string
  recordedAt: number
}

const rows: Row[] = [
  {
    id: 1,
    document: "MSA.docx",
    kind: "renewal",
    date: "2027-01-01",
    formula: null,
    noticeDays: 90,
    actBy: "2026-10-03",
    party: "Customer",
    quote: "This Agreement renews for successive one-year terms.",
    reason: "Auto-renewal anniversary",
    author: "Synthetic Counsel",
    recordedAt: Date.UTC(2026, 9, 3),
  },
  {
    id: 2,
    document: "SOW.docx",
    kind: "deliverable",
    date: null,
    formula: "30 days after acceptance",
    noticeDays: null,
    actBy: null,
    party: null,
    quote: "Delivery follows <img src=x onerror=alert(1)> acceptance.",
    reason: "Delivery window with no anchor date",
    author: "Synthetic Counsel",
    recordedAt: Date.UTC(2026, 9, 4),
  },
]

async function render(response: () => Promise<Response> | Response) {
  const browser = new Window({ url: "http://localhost:5275" })
  const names = ["window", "document", "HTMLElement", "Node", "Event", "IS_REACT_ACT_ENVIRONMENT"] as const
  const saved = names.map((name) => [name, Object.getOwnPropertyDescriptor(globalThis, name)] as const)
  Object.defineProperties(
    globalThis,
    Object.fromEntries(
      names.map((name) => [
        name,
        {
          configurable: true,
          writable: true,
          value: name === "IS_REACT_ACT_ENVIRONMENT" ? true : browser[name as keyof Window],
        },
      ]),
    ),
  )
  const originalFetch = globalThis.fetch
  const requests: string[] = []
  globalThis.fetch = (async (input) => {
    requests.push(String(input))
    return response()
  }) as typeof fetch
  const opened: string[] = []
  let asked = 0
  const { createRoot } = await import("react-dom/client")
  const { default: KeyDatesRegister } = await import("../src/components/key-dates/KeyDatesRegister")
  const container = document.createElement("div")
  document.body.append(container)
  const root = createRoot(container)
  const mount = async () =>
    act(async () => {
      root.render(
        createElement(KeyDatesRegister, {
          matterId: "matter",
          documents: ["MSA.docx"],
          onClose: () => {},
          onOpenDocument: (name: string) => opened.push(name),
          onAsk: () => asked++,
        }),
      )
    })
  await mount()
  const cleanup = async () => {
    await act(async () => root.unmount())
    container.remove()
    globalThis.fetch = originalFetch
    await browser.happyDOM.close()
    for (const [name, descriptor] of saved) {
      if (descriptor) Object.defineProperty(globalThis, name, descriptor)
      else delete (globalThis as Record<string, unknown>)[name]
    }
  }
  return {
    mount,
    cleanup,
    requests,
    opened,
    asked: () => asked,
    text: () => container.textContent ?? "",
    html: () => container.innerHTML,
    button: (label: string) =>
      Array.from(container.querySelectorAll<HTMLButtonElement>("button")).find(
        (item) => item.textContent?.trim() === label,
      ),
  }
}

test("shows stored dates soonest first and labels a relative formula as unresolved", async () => {
  const view = await render(() => Response.json({ rows }))
  try {
    expect(view.requests).toHaveLength(1)
    expect(view.requests[0]!.endsWith("/matters/matter/key-dates")).toBe(true)
    const text = view.text()
    expect(text).toContain("2 entries · 1 entry has no calendar date")
    expect(text).toContain("Auto-renewal anniversary")
    expect(text).toContain("Renewal")
    expect(text).toContain("90 days")
    expect(text).toContain("No calendar date")
    expect(text).toContain("Unresolved formula")
    expect(text).toContain("30 days after acceptance")
    expect(text).toContain("not a confirmed deadline")
    expect(text.indexOf("Auto-renewal anniversary")).toBeLessThan(text.indexOf("Delivery window"))
    // A stored quote is displayed as text, never parsed as markup.
    expect(text).toContain("Delivery follows <img src=x onerror=alert(1)> acceptance.")
    expect(view.html()).not.toContain("<img")
    await act(async () => view.button("MSA.docx")!.click())
    expect(view.opened).toEqual(["MSA.docx"])
    expect(view.text()).toContain("SOW.docx (not in this matter now)")
  } finally {
    await view.cleanup()
  }
})

test("an empty register says none are recorded and offers the deadline question", async () => {
  const view = await render(() => Response.json({ rows: [] }))
  try {
    const text = view.text()
    expect(text).toContain("No key dates recorded yet")
    expect(text).toContain("An empty register means none have been recorded, not that the matter has no deadlines.")
    await act(async () => view.button("Ask about deadlines")!.click())
    expect(view.asked()).toBe(1)
  } finally {
    await view.cleanup()
  }
})

test("a failed read shows the error and retries the same read", async () => {
  let attempt = 0
  const view = await render(() => {
    attempt++
    return attempt === 1
      ? Response.json({ error: "The register is unavailable." }, { status: 503 })
      : Response.json({ rows: [rows[0]] })
  })
  try {
    expect(view.text()).toContain("The register is unavailable.")
    await act(async () => view.button("Try again")!.click())
    expect(view.requests).toHaveLength(2)
    expect(view.text()).toContain("Auto-renewal anniversary")
    expect(view.text()).not.toContain("The register is unavailable.")
  } finally {
    await view.cleanup()
  }
})
