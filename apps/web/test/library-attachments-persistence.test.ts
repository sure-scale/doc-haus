import { expect, test } from "bun:test"
import { Window } from "happy-dom"
import { act, createElement } from "react"
import { useLibraryAttachments } from "../src/components/chat/use-library-attachments"

test("pending files return to their conversation after navigation and clear after sending", async () => {
  const browser = new Window({ url: "http://localhost" })
  const names = ["window", "document", "HTMLElement", "Node", "Event", "sessionStorage", "IS_REACT_ACT_ENVIRONMENT"] as const
  const saved = names.map(name => [name, Object.getOwnPropertyDescriptor(globalThis, name)] as const)
  Object.defineProperties(globalThis, Object.fromEntries(names.map(name => [name, {
    configurable: true, writable: true, value: name === "IS_REACT_ACT_ENVIRONMENT" ? true : browser[name as keyof Window],
  }])))
  const { createRoot } = await import("react-dom/client")
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
    for (const [name, descriptor] of saved) {
      if (descriptor) Object.defineProperty(globalThis, name, descriptor)
      else delete (globalThis as Record<string, unknown>)[name]
    }
  }
})
