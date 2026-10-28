import { expect, test } from "bun:test"
import { Window } from "happy-dom"
import { act, createElement, useRef, useState } from "react"
import { useLibraryChatKey } from "../src/components/chat/use-library-chat-key"

test("a newly created session keeps its composer, while navigating to a new chat resets it", async () => {
  const browser = new Window({ url: "http://localhost" })
  const names = ["window", "document", "HTMLElement", "Node", "Event", "IS_REACT_ACT_ENVIRONMENT"] as const
  const saved = names.map(name => [name, Object.getOwnPropertyDescriptor(globalThis, name)] as const)
  Object.defineProperties(globalThis, Object.fromEntries(names.map(name => [name, {
    configurable: true, writable: true, value: name === "IS_REACT_ACT_ENVIRONMENT" ? true : browser[name as keyof Window],
  }])))
  const { createRoot } = await import("react-dom/client")
  const container = document.createElement("div")
  const root = createRoot(container)
  const created = { current: undefined as string | undefined }
  function Composer() {
    const [text, setText] = useState("draft attachment")
    return createElement("button", { onClick: () => setText("unsent attachment") }, text)
  }
  function Page(props: { session?: string; prompt?: string }) {
    const ref = useRef(created.current)
    ref.current = created.current
    const key = useLibraryChatKey(props.session, props.prompt, ref)
    created.current = ref.current
    return createElement(Composer, { key })
  }
  try {
    await act(async () => root.render(createElement(Page, { prompt: "import" })))
    act(() => container.querySelector("button")!.click())
    created.current = "session-1"
    await act(async () => root.render(createElement(Page, { session: "session-1" })))
    expect(container.textContent).toBe("unsent attachment")
    await act(async () => root.render(createElement(Page, {})))
    expect(container.textContent).toBe("draft attachment")
    await act(async () => root.render(createElement(Page, { session: "session-1" })))
    expect(created.current).toBeUndefined()
  } finally {
    await act(async () => root.unmount())
    for (const [name, descriptor] of saved) {
      if (descriptor) Object.defineProperty(globalThis, name, descriptor)
      else delete (globalThis as Record<string, unknown>)[name]
    }
  }
})
