import { expect, test } from "bun:test"
import { Window } from "happy-dom"
import { act, createElement, useRef, useState } from "react"
import { useLibraryAttachments } from "../src/components/chat/use-library-attachments"
import { useLibraryChatKey } from "../src/components/chat/use-library-chat-key"

async function withDom<T>(run: (createRoot: typeof import("react-dom/client").createRoot) => Promise<T>) {
  const browser = new Window({ url: "http://localhost" })
  const names = ["window", "document", "HTMLElement", "Node", "Event", "sessionStorage", "IS_REACT_ACT_ENVIRONMENT"] as const
  const saved = names.map(name => [name, Object.getOwnPropertyDescriptor(globalThis, name)] as const)
  Object.defineProperties(globalThis, Object.fromEntries(names.map(name => [name, {
    configurable: true, writable: true, value: name === "IS_REACT_ACT_ENVIRONMENT" ? true : browser[name as keyof Window],
  }])))
  try {
    const { createRoot } = await import("react-dom/client")
    return await run(createRoot)
  } finally {
    await browser.happyDOM.close()
    for (const [name, descriptor] of saved) {
      if (descriptor) Object.defineProperty(globalThis, name, descriptor)
      else delete (globalThis as Record<string, unknown>)[name]
    }
  }
}

test("pending files return to their conversation after navigation and clear after sending", async () => {
  await withDom(async (createRoot) => {
    const container = document.createElement("div")
    const root = createRoot(container)
    const source = { id: "one", name: "guide.docx", source: "one.txt", original: "one.docx" }
    function Composer({ session }: { session?: string }) {
      const [files, setFiles] = useLibraryAttachments("/library", session)
      return createElement("div", null,
        createElement("span", null, files.map(file => file.name).join(",")),
        createElement("button", { onClick: () => setFiles([source]) }, "add"),
        createElement("button", { onClick: () => setFiles([]) }, "clear"))
    }
    try {
      await act(async () => root.render(createElement(Composer, { key: "new" })))
      act(() => container.querySelectorAll("button")[0]!.click())
      await act(async () => root.render(createElement(Composer, { key: "new", session: "session-1" })))
      await act(async () => root.render(createElement(Composer, { key: "other", session: "session-2" })))
      expect(container.querySelector("span")!.textContent).toBe("")
      await act(async () => root.render(createElement(Composer, { key: "session-1", session: "session-1" })))
      expect(container.querySelector("span")!.textContent).toBe("guide.docx")
      act(() => container.querySelectorAll("button")[1]!.click())
      await act(async () => root.render(createElement(Composer, { key: "other", session: "session-2" })))
      await act(async () => root.render(createElement(Composer, { key: "session-1", session: "session-1" })))
      expect(container.querySelector("span")!.textContent).toBe("")
    } finally {
      await act(async () => root.unmount())
    }
  })
})

test("a newly created session keeps its composer, while navigating to a new chat resets it", async () => {
  await withDom(async (createRoot) => {
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
    }
  })
})
