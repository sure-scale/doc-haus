import { expect, test } from "bun:test"

test("a cached unversioned pane stops before Office access or sync startup", async () => {
  const globals = ["window", "document"] as const
  const previous = globals.map((key) => Object.getOwnPropertyDescriptor(globalThis, key))
  const elements = new Map<string, { textContent: string; hidden: boolean; dataset: object }>()
  let officeReads = 0
  try {
    Object.defineProperty(globalThis, "window", {
      configurable: true,
      value: {
        location: { href: "https://localhost:4612/word-addin/pane.html" },
        get Office() {
          officeReads++
          return undefined
        },
      },
    })
    Object.defineProperty(globalThis, "document", {
      configurable: true,
      value: {
        getElementById(id: string) {
          if (!elements.has(id)) elements.set(id, { textContent: "", hidden: false, dataset: {} })
          return elements.get(id)
        },
      },
    })
    await import("./pane")
    expect(officeReads).toBe(0)
    expect(elements.get("status")?.textContent).toBe("Update needed")
    expect(elements.get("message")?.textContent).toContain("out of date")
    expect(elements.get("actions")?.hidden).toBe(true)
  } finally {
    globals.forEach((key, index) => {
      if (previous[index]) Object.defineProperty(globalThis, key, previous[index]!)
      else Reflect.deleteProperty(globalThis, key)
    })
  }
})
