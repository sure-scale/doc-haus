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
    act(() => (document.querySelector('.toast-close') as HTMLButtonElement).click())
    expect(document.querySelector('.toast-leaving')).not.toBeNull()
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

test("a toast fired while a modal dialog is open shows inside it, then returns to the page when it closes", async () => {
  const browser = new Window({url:"http://localhost"})
  const names = ["window", "document", "HTMLElement", "HTMLDialogElement", "Node", "Event", "IS_REACT_ACT_ENVIRONMENT"] as const
  const saved = names.map(name => [name,Object.getOwnPropertyDescriptor(globalThis,name)] as const)
  Object.defineProperties(globalThis,Object.fromEntries(names.map(name => [name,{configurable:true,writable:true,value:name === "IS_REACT_ACT_ENVIRONMENT" ? true : browser[name as keyof Window]}])))
  const {createRoot} = await import("react-dom/client")
  const container = document.createElement("div"); document.body.append(container)
  const dialog = document.createElement("dialog"); document.body.append(dialog)
  // happy-dom does not implement :modal; a real browser matches it while showModal() holds the dialog open.
  const matches = dialog.matches.bind(dialog)
  dialog.matches = (selector: string) => selector === ":modal" ? dialog.open : matches(selector)
  const root = createRoot(container)
  let notify: (type: "success", message: string) => void = () => {}
  function Notify() { notify = useToast(); return null }
  try {
    await act(async () => root.render(createElement(ToastProvider,null,createElement(Notify))))
    dialog.showModal()
    act(() => notify("success","Saved."))
    expect(dialog.querySelector(".toast")?.textContent).toContain("Saved.")
    act(() => dialog.close())
    expect(dialog.querySelector(".toast")).toBeNull()
    expect(document.body.querySelector(":scope > .toast-stack .toast")?.textContent).toContain("Saved.")
    await act(async () => root.unmount())
  } finally {
    container.remove(); dialog.remove()
    for (const [name,descriptor] of saved) if (descriptor) Object.defineProperty(globalThis,name,descriptor); else delete (globalThis as Record<string,unknown>)[name]
    await browser.happyDOM.close()
  }
})
