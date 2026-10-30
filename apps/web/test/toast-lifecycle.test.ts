import { expect, spyOn, test } from "bun:test"
import { Window } from "happy-dom"
import { act, createElement } from "react"
import { ToastProvider, useToast } from "../src/components/Toast"

test("unmount cancels pending toast expiry and exit transition timers", async () => {
  const browser = new Window({url:"http://localhost"})
  const names = ["window", "document", "HTMLElement", "Node", "Event", "IS_REACT_ACT_ENVIRONMENT"] as const
  const saved = names.map(name => [name,Object.getOwnPropertyDescriptor(globalThis,name)] as const)
  Object.defineProperties(globalThis,Object.fromEntries(names.map(name => [name,{configurable:true,writable:true,value:name === "IS_REACT_ACT_ENVIRONMENT" ? true : browser[name as keyof Window]}])))
  const scheduled = spyOn(globalThis,"setTimeout"), cleared = spyOn(globalThis,"clearTimeout")
  const {createRoot} = await import("react-dom/client")
  const container = document.createElement("div"); document.body.append(container)
  const root = createRoot(container)
  function Notify() { const toast = useToast(); return createElement("button", {onClick:() => {toast("info","Info");toast("error","Error")}}, "Notify") }
  try {
    await act(async () => root.render(createElement(ToastProvider,null,createElement(Notify))))
    act(() => (container.querySelector("button") as HTMLButtonElement).click())
    act(() => (container.querySelector('.toast-close') as HTMLButtonElement).click())
    expect(container.querySelector('.toast-leaving')).not.toBeNull()
    const timers = scheduled.mock.calls.flatMap((call,i) => [5000,8000,180].includes(call[1] as number) ? [scheduled.mock.results[i]!.value] : [])
    expect(timers).toHaveLength(3)
    await act(async () => root.unmount())
    for (const timer of timers) expect(cleared.mock.calls.some(call => call[0] === timer)).toBe(true)
  } finally {
    scheduled.mockRestore(); cleared.mockRestore(); container.remove()
    for (const [name,descriptor] of saved) if (descriptor) Object.defineProperty(globalThis,name,descriptor); else delete (globalThis as Record<string,unknown>)[name]
    await browser.happyDOM.close()
  }
})
